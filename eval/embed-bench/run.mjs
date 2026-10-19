// Embed benchmark runner. Each product × file gets a fresh headless
// Chromium; memory is the RSS of the whole browser process tree (page,
// workers, GPU), sampled every 250 ms.
//   node prep.mjs && npx vite build && node run.mjs [product:file ...]
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { createRequire } from "node:module";
import { chromium } from "playwright";
import { FILES } from "./files.mjs";
import { fidelity, norm, same } from "./lib.mjs";

const require = createRequire(import.meta.url);
const engine = require("../../packages/gridpath/engine/gridpath_engine.js");
// product → page (and query) it runs on.
const PAGES = { gridpath: "gridpath.html", spreadjs: "spreadjs.html", "spreadjs-lazy": "spreadjs.html?mode=lazy", hot: "hot.html" };
const PRODUCTS = ["gridpath", "spreadjs", "hot"];
const TIMEOUT = { "rich-model": 120_000, analyst: 15 * 60_000, large: 30 * 60_000 };
const EDITS = 5;

const MIME = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".wasm": "application/wasm", ".json": "application/json", ".xlsx": "application/octet-stream" };
const server = http.createServer((req, res) => {
  const url = new URL(req.url, "http://x");
  let file;
  if (url.pathname.startsWith("/files/")) file = FILES[path.basename(url.pathname, ".xlsx")]?.path;
  else if (url.pathname.startsWith("/prep/")) file = path.resolve("results/prep", path.basename(url.pathname));
  else file = path.resolve("dist", "." + url.pathname);
  if (!file || !fs.existsSync(file)) {
    res.writeHead(404).end();
    return;
  }
  res.writeHead(200, { "content-type": MIME[path.extname(file)] ?? "application/octet-stream" });
  fs.createReadStream(file).pipe(res);
});
await new Promise((r) => server.listen(4391, r));

function treeRssMb(tag) {
  const ps = execFileSync("ps", ["-A", "-o", "pid=,ppid=,rss=,command="], { encoding: "utf8", maxBuffer: 64 << 20 });
  const rows = ps.trim().split("\n").map((l) => {
    const m = /^\s*(\d+)\s+(\d+)\s+(\d+)\s+(.*)$/.exec(l);
    return { pid: +m[1], ppid: +m[2], rss: +m[3], cmd: m[4] };
  });
  const keep = new Set(rows.filter((r) => r.cmd.includes(tag)).map((r) => r.pid));
  for (let grew = true; grew; ) {
    grew = false;
    for (const r of rows) if (!keep.has(r.pid) && keep.has(r.ppid)) (keep.add(r.pid), (grew = true));
  }
  return rows.filter((r) => keep.has(r.pid)).reduce((n, r) => n + r.rss, 0) / 1024;
}

const median = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];

async function runOne(product, id) {
  const prep = JSON.parse(fs.readFileSync(`results/prep/${id}.json`, "utf8"));
  const tag = `bench-${product}-${id}-${Date.now()}`;
  const out = { product, file: id };
  const browser = await chromium.launch({ headless: true, args: [`--bench-tag=${tag}`, "--js-flags=--expose-gc --max-old-space-size=16384"] });
  let peak = 0;
  const sampler = setInterval(() => {
    try {
      peak = Math.max(peak, treeRssMb(tag));
    } catch {}
  }, 250);
  const errors = [];
  try {
    const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
    page.on("pageerror", (e) => errors.push(e.message.slice(0, 300)));
    page.on("console", (m) => m.type() === "error" && errors.push(m.text().slice(0, 300)));
    const idleRss = treeRssMb(tag);
    const pg = PAGES[product];
    await page.goto(`http://localhost:4391/${pg}${pg.includes("?") ? "&" : "?"}file=${id}`);
    await page.waitForFunction(() => ["calculated", "error"].includes(window.__bench?.phase), null, { timeout: TIMEOUT[id], polling: 250 });
    const b = await page.evaluate(() => ({ phase: window.__bench.phase, error: window.__bench.error, t: window.__bench.t }));
    if (b.phase === "error") throw new Error(b.error);
    out.interactiveS = (b.t.interactive - b.t.start) / 1000;
    out.calculatedS = (b.t.calculated - b.t.start) / 1000;
    await page.evaluate(() => window.gc?.());
    await page.waitForTimeout(2000);
    out.settledMb = Math.round(treeRssMb(tag) - idleRss);
    out.peakMb = Math.round(peak - idleRss);

    const times = [];
    try {
      for (let k = 0; k < EDITS; k++) {
        const v = k % 2 === 0 ? prep.input.changed : prep.input.original;
        times.push(await page.evaluate((v) => window.__bench.edit(v), v));
      }
      // An odd number of edits: the input ends at its changed value.
      out.editMs = Math.round(median(times) * 10) / 10;
      out.editRuns = times.map((t) => Math.round(t));
    } catch (e) {
      // The engine took the input but the watched dependent never moved.
      out.editError = String(e.message ?? e).split("\n")[0].slice(0, 300);
      await page.evaluate((v) => window.__bench.edit(v).catch(() => 0), prep.input.changed);
    }

    const dump = await page.evaluate(() => window.__bench.dump());
    if (dump) {
      // Edits moved some formulas: score against the state Excel saved by
      // reading the input back to its original value first.
      await page.evaluate((v) => window.__bench.edit(v).catch(() => 0), prep.input.original);
      const fresh = await page.evaluate(() => window.__bench.dump());
      const cached = JSON.parse(fs.readFileSync(`results/prep/${id}.cached.json`, "utf8"));
      let agree = 0, total = 0;
      const misses = [];
      cached.forEach((vals, si) =>
        vals.forEach((v, k) => {
          total++;
          if (same(fresh[si][k], v)) agree++;
          else if (misses.length < 15) misses.push({ sheet: prep.sheets[si].name, i: k, excel: norm(v), got: norm(fresh[si][k]) });
        }),
      );
      out.calc = { agree, total, misses };
      await page.evaluate((v) => window.__bench.edit(v).catch(() => 0), prep.input.changed);
    } else {
      const g = JSON.parse(fs.readFileSync(`results/prep/${id}.gridpath-calc.json`, "utf8"));
      out.calc = { agree: g.agree, total: g.total, misses: g.misses.slice(0, 15), note: "same engine build, scored in Node" };
    }

    const t = performance.now();
    const b64 = await page.evaluate(() => window.__bench.save());
    out.saveS = (performance.now() - t) / 1000;
    const bytes = Buffer.from(b64, "base64");
    fs.mkdirSync("results/out", { recursive: true });
    fs.writeFileSync(`results/out/${product}-${id}.xlsx`, bytes);
    out.fidelity = await fidelity(FILES[id].path, bytes);
    // The saved file reopens, and carries the edit and the formulas.
    try {
      const wb = engine.Workbook.openDeferred(bytes);
      const names = wb.sheetNames();
      const snap = JSON.parse(wb.snapshot());
      const count = (s) => s.sheets.reduce((n, sh) => n + sh.cells.filter((c) => c.f).length, 0);
      const orig = JSON.parse(engine.Workbook.openDeferred(fs.readFileSync(FILES[id].path)).snapshot());
      const si = names.indexOf(prep.input.sheet);
      const m = /^([A-Z]+)(\d+)$/.exec(prep.input.ref);
      let col = 0;
      for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
      const cell = JSON.parse(wb.cell(si, Number(m[2]), col));
      out.reopen = { ok: true, editKept: same(cell.value, prep.input.changed), formulas: { before: count(orig), after: count(snap) } };
    } catch (e) {
      out.reopen = { ok: false, error: String(e).slice(0, 200) };
    }
  } catch (e) {
    out.error = String(e.message ?? e).slice(0, 500);
  } finally {
    clearInterval(sampler);
    out.peakMb ??= Math.round(peak);
    out.consoleErrors = errors.slice(0, 5);
    await browser.close();
  }
  return out;
}

const jobs = process.argv.slice(2).length
  ? process.argv.slice(2).map((j) => j.split(":"))
  : Object.keys(FILES).flatMap((id) => PRODUCTS.map((p) => [p, id]));
fs.mkdirSync("results/runs", { recursive: true });
for (const [product, id] of jobs) {
  process.stdout.write(`${product} × ${id} … `);
  const r = await runOne(product, id);
  fs.writeFileSync(`results/runs/${product}-${id}.json`, JSON.stringify(r, null, 2));
  console.log(r.error ? `ERROR ${r.error.split("\n")[0]}` : `open ${r.interactiveS?.toFixed(1)}s, calc ${r.calculatedS?.toFixed(1)}s, ${r.settledMb} MB (peak ${r.peakMb}), edit ${r.editMs} ms, calc ${r.calc?.agree}/${r.calc?.total}, parts identical ${r.fidelity?.identical}/${r.fidelity?.parts}`);
}
server.close();
