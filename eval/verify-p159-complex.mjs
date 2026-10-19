#!/usr/bin/env node
/**
 * Strict checks for the two harder p159 tasks, beyond what grade.mjs's
 * label assertions can express. Loads the original and the output into the
 * gridpath engine (same evaluator for both lanes) and diffs cell by cell.
 *
 *   node eval/verify-p159-complex.mjs <run-dir> [<run-dir>...]
 */
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evalDir = path.dirname(fileURLToPath(import.meta.url));
// The engine build: a checkout of the product repo by default, or point
// GRIDPATH_ENGINE_DIR at node_modules/gridpath/engine from the npm package.
const engineDir = path.resolve(process.env.GRIDPATH_ENGINE_DIR ?? path.join(evalDir, "..", "packages", "gridpath", "engine")) + path.sep;
const eng = createRequire(engineDir)(engineDir + "gridpath_engine.js");

const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ""; while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); } return s; };
const snap = (file) => {
  const j = JSON.parse(eng.Workbook.open(fs.readFileSync(file)).snapshot());
  const m = new Map();
  for (const sh of j.sheets) for (const c of sh.cells) m.set(`${sh.name}!${colName(c.c)}${c.r}`, c);
  return m;
};
const close = (a, b) => (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(a), Math.abs(b)) : a === b);
const A = (ref) => `Assumptions!${ref}`;
const QCOLS = ["BX", "BY", "BZ", "CA"];

function scenario(orig, out, meta) {
  const want = new Map();
  for (const c of QCOLS) { want.set(A(`${c}54`), 1); want.set(A(`${c}58`), 0.93); want.set(A(`${c}60`), 11000); }
  let inputsOk = 0;
  for (const [k, v] of want) if (!out.get(k)?.f && close(out.get(k)?.v, v)) inputsOk++;
  // Any other cell whose input (formula text, or literal when not a formula) changed.
  const stray = [];
  for (const k of new Set([...orig.keys(), ...out.keys()])) {
    if (want.has(k)) continue;
    const o = orig.get(k), n = out.get(k);
    const oi = o?.f ?? o?.v ?? null, ni = n?.f ?? n?.v ?? null;
    if (!close(oi, ni) && !(oi == null && (ni == null || ni === ""))) stray.push(k);
  }
  const fy = out.get(A("CB65"))?.v;
  const said = String(meta.result ?? "").replace(/,/g, "");
  const nums = [...said.matchAll(/\d+(?:\.\d+)?/g)].map((x) => Number(x[0]));
  const reported = nums.some((x) => Math.abs(x - 135720.44) <= 135720.44 * 0.001 || Math.abs(x - 135.72044) <= 0.14);
  return { inputs_ok: `${inputsOk}/12`, stray_changes: stray.length, stray_sample: stray.slice(0, 5), fy2027_recalculated: fy, fy2027_ok: close(Math.round(fy), 135720), reported_right_answer: reported };
}

function addRow(orig, out) {
  const shift = (k) => { const m = k.match(/^Assumptions!([A-Z]+)(\d+)$/); return m && Number(m[2]) >= 62 ? `Assumptions!${m[1]}${Number(m[2]) + 1}` : k; };
  let formulas = 0, valueMismatch = 0, missing = 0; const sample = [];
  for (const [k, o] of orig) {
    if (!o.f) continue;
    formulas++;
    const n = out.get(shift(k));
    if (!n || !n.f) { missing++; if (sample.length < 5) sample.push(`${k} formula gone`); continue; }
    if (!close(o.v, n.v)) { valueMismatch++; if (sample.length < 5) sample.push(`${k}: ${o.v} -> ${n.v}`); }
  }
  const label = out.get(A("A62"))?.v;
  let newOk = 0, newCols = 0;
  for (let c = colNum("BE"); c <= colNum("CG"); c++) {
    const col = colName(c), rev = out.get(A(`${col}61`))?.v, cap = out.get(A(`${col}55`))?.v;
    if (typeof rev !== "number") continue;
    newCols++;
    const n = out.get(A(`${col}62`));
    if (n?.f && typeof cap === "number" && cap !== 0 && close(n.v, rev / cap)) newOk++;
  }
  return { label_at_row_62: /per ending GW/i.test(String(label ?? "")), new_row_formulas_ok: `${newOk}/${newCols}`, existing_formulas: formulas, formulas_lost: missing, values_changed: valueMismatch, sample };
}

const origCache = new Map();
for (const dir of process.argv.slice(2)) {
  const meta = JSON.parse(fs.readFileSync(path.join(dir, "meta.json"), "utf8"));
  const cli = JSON.parse(fs.readFileSync(path.join(dir, "cli-stdout.json"), "utf8"));
  const origPath = path.join(dir, "original.xlsx");
  if (!origCache.size) origCache.set("o", snap(origPath));
  const orig = origCache.get("o"), out = snap(path.join(dir, "output.xlsx"));
  const res = meta.task === "p159-scenario" ? scenario(orig, out, { result: cli.result }) : addRow(orig, out);
  console.log(JSON.stringify({ run: path.basename(dir), task: meta.task, harness: meta.harness, wall_s: Math.round(meta.wall_ms / 100) / 10, turns: meta.num_turns, out_tokens: meta.usage?.output_tokens, cost: Math.round(meta.total_cost_usd * 100) / 100, ...res }));
}
