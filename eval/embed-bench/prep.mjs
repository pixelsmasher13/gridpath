// Per file: Excel's cached formula results, GridPath's engine values against
// them, and an edit target (an input with many dependents, and one
// dependent to watch). Run before run.mjs:  node prep.mjs [fileId ...]
import fs from "node:fs";
import { createRequire } from "node:module";
import { FILES } from "./files.mjs";
import { a1, cachedFormulaValues, norm, same } from "./lib.mjs";

const require = createRequire(import.meta.url);
const engine = require("../../packages/gridpath/engine/gridpath_engine.js");
fs.mkdirSync("results/prep", { recursive: true });

const ids = process.argv.slice(2).length ? process.argv.slice(2) : Object.keys(FILES);
for (const id of ids) {
  const { path, editBudgetMs } = FILES[id];
  console.log(`\n== ${id}`);
  const cached = await cachedFormulaValues(path);
  const formulaCount = cached.reduce((n, s) => n + s.values.length, 0);
  console.log(`cached formula results: ${formulaCount}`);

  const bytes = fs.readFileSync(path);
  let t = performance.now();
  const wb = engine.Workbook.open(bytes);
  const openMs = performance.now() - t;
  const names = wb.sheetNames();
  const snap = JSON.parse(wb.snapshot());
  const value = new Map();
  const constants = [];
  snap.sheets.forEach((sh, si) => {
    for (const c of sh.cells) {
      value.set(`${si}|${c.r}|${c.c}`, c.v);
      if (!c.f && typeof c.v === "number") constants.push([si, c.r, c.c, c.v]);
    }
  });

  // GridPath's engine against Excel's cached results.
  let agree = 0;
  const misses = [];
  for (const s of cached) {
    const si = names.indexOf(s.name);
    for (let i = 0; i < s.values.length; i++) {
      const r = s.cells[2 * i], c = s.cells[2 * i + 1];
      const got = value.get(`${si}|${r}|${c}`);
      if (same(got, s.values[i])) agree++;
      else if (misses.length < 25) misses.push({ sheet: s.name, cell: a1(r, c), excel: norm(s.values[i]), got: norm(got) });
    }
  }
  console.log(`gridpath engine: open+evaluate ${(openMs / 1000).toFixed(1)} s (node), agreement ${agree}/${formulaCount}`);
  fs.writeFileSync(`results/prep/${id}.gridpath-calc.json`, JSON.stringify({ openMs, agree, total: formulaCount, misses }, null, 2));

  // Edit target: the sampled numeric input whose change moves the most formulas.
  let seed = 12345;
  const rand = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff);
  for (let i = constants.length - 1; i > 0; i--) {
    const j = Math.floor(rand() * (i + 1));
    [constants[i], constants[j]] = [constants[j], constants[i]];
  }
  let best = null;
  t = performance.now();
  let tried = 0;
  for (const [si, r, c, v] of constants) {
    if (performance.now() - t > editBudgetMs) break;
    tried++;
    const nv = v === 0 ? 1 : Math.round(v * 1.1 * 1000) / 1000 + 1;
    wb.setInput(si, r, c, String(nv));
    const res = JSON.parse(wb.evaluateEdited());
    wb.setInput(si, r, c, String(v));
    wb.evaluateEdited();
    if (res.full) continue;
    const moved = res.cells.filter((x) => x.f && !same(x.v, value.get(`${x.sheet}|${x.r}|${x.c}`)));
    if (!best || moved.length > best.moved.length) best = { si, r, c, v, nv, moved };
  }
  if (!best || best.moved.length === 0) throw new Error(`${id}: no input with dependents found in ${tried} tries`);
  // Watch a dependent on another sheet when there is one (cross-sheet chains).
  const far = best.moved.find((x) => x.sheet !== best.si) ?? best.moved.reduce((a, b) => (b.r > a.r ? b : a));
  const prep = {
    id,
    input: { sheet: names[best.si], ref: a1(best.r, best.c), original: best.v, changed: best.nv },
    dep: { sheet: names[far.sheet], ref: a1(far.r, far.c), before: value.get(`${far.sheet}|${far.r}|${far.c}`), after: far.v },
    dependentsMoved: best.moved.length,
    candidatesTried: tried,
    sheets: cached.map((s) => ({ name: s.name, cells: s.cells })),
  };
  console.log(`edit target: ${prep.input.sheet}!${prep.input.ref} (${tried} tried) moves ${best.moved.length} formulas; watching ${prep.dep.sheet}!${prep.dep.ref}`);
  fs.writeFileSync(`results/prep/${id}.json`, JSON.stringify(prep));
  fs.writeFileSync(`results/prep/${id}.cached.json`, JSON.stringify(cached.map((s) => s.values)));
  wb.free();
}
