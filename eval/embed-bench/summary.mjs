// results/runs/*.json → results/summary.md
import fs from "node:fs";
import { FILES } from "./files.mjs";

const NAMES = { gridpath: "GridPath 0.2.0", spreadjs: "SpreadJS 19.2 (+charts, shapes, pivot, slicers)", hot: "Handsontable 18.1 + HyperFormula 3.5 + ExcelJS 4.4", "spreadjs-lazy": "SpreadJS 19.2, lazy open mode" };
const runs = fs.readdirSync("results/runs").map((f) => JSON.parse(fs.readFileSync(`results/runs/${f}`, "utf8")));
const s = (x, d = 1) => (x === undefined || x === null ? "—" : Number(x).toFixed(d));
const pct = (c) => (c && c.total ? `${((100 * c.agree) / c.total).toFixed(2)}% (${c.total - c.agree} differ)` : "—");

let md = `# Embed benchmark\n\nHeadless Chromium, production builds, ${new Date().toISOString().slice(0, 10)}. One run per cell; edit = median of 5.\n`;
for (const [id, f] of Object.entries(FILES)) {
  const rows = runs.filter((r) => r.file === id);
  if (!rows.length) continue;
  md += `\n## ${f.label}\n\n| | ${rows.map((r) => NAMES[r.product]).join(" | ")} |\n|---|${rows.map(() => "---").join("|")}|\n`;
  const line = (label, fn) => (md += `| ${label} | ${rows.map((r) => (r.error && fn(r) === "—" ? (/crashed/.test(r.error) ? "tab crashed" : "failed") : fn(r))).join(" | ")} |\n`);
  line("Open → grid usable (s)", (r) => s(r.interactiveS));
  line("Open → fully recalculated (s)", (r) => s(r.calculatedS));
  line("Memory after open (MB)", (r) => s(r.settledMb, 0));
  line("Peak memory (MB)", (r) => s(r.peakMb, 0));
  line("Edit → dependent updated (ms)", (r) => (r.editError ? (/crashed/.test(r.editError) ? "tab crashed" : "never updated") : s(r.editMs, 1)));
  line("Agrees with Excel's results", (r) => pct(r.calc));
  line("Save (s)", (r) => s(r.saveS));
  line("Parts byte-identical", (r) => (r.fidelity ? `${r.fidelity.identical}/${r.fidelity.parts}` : "—"));
  // By feature kind: products that renumber parts (sheet2.xml → sheet3.xml) lose nothing by renaming alone.
  line("Lost on save", (r) => { if (!r.fidelity) return "—"; const lost = Object.entries(r.fidelity.features).filter(([, v]) => v.after < v.before).map(([k, v]) => `${k} ${v.before}→${v.after}`); return lost.length ? lost.join(", ") : "nothing"; });
  line("Defined names kept", (r) => (r.fidelity ? `${r.fidelity.definedNames.after}/${r.fidelity.definedNames.before}` : "—"));
  line("Formulas in saved file", (r) => (r.reopen?.formulas ? `${r.reopen.formulas.after}/${r.reopen.formulas.before}` : "—"));
  line("Error", (r) => (r.error ? r.error.split("\n")[0].slice(0, 90) : ""));
}
fs.writeFileSync("results/summary.md", md);
console.log(md);
