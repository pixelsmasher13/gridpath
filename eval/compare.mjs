#!/usr/bin/env node
/**
 * Sweep eval/runs/ into one scoreboard: task × harness × N runs →
 * assertion pass rates, fully-clean-run rate, median duration, median
 * output tokens, summed metered cost.
 *
 *   node eval/compare.mjs [task-id ...]     (default: every task with runs)
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evalDir = path.dirname(fileURLToPath(import.meta.url));
const runsDir = path.join(evalDir, "runs");

const wanted = process.argv.slice(2);
const tasks = (fs.existsSync(runsDir) ? fs.readdirSync(runsDir) : []).filter(
  (t) => (!wanted.length || wanted.includes(t)) && fs.statSync(path.join(runsDir, t)).isDirectory(),
);

const median = (xs) => {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b);
  return s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2;
};
const fmtMs = (ms) => (ms == null ? "—" : ms >= 10_000 ? `${Math.round(ms / 1000)}s` : `${(ms / 1000).toFixed(1)}s`);

for (const task of tasks) {
  const taskDir = path.join(runsDir, task);
  const groups = new Map(); // harness -> rows
  for (const runName of fs.readdirSync(taskDir)) {
    const runDir = path.join(taskDir, runName);
    if (!fs.statSync(runDir).isDirectory()) continue;
    const harness = runName.endsWith("-gridpath") ? "gridpath" : runName.endsWith("-cc") ? "claude-code" : null;
    if (!harness) continue;
    const row = { runName };
    const metaPath = path.join(runDir, "meta.json");
    if (fs.existsSync(metaPath)) {
      const meta = JSON.parse(fs.readFileSync(metaPath, "utf8"));
      row.duration = meta.duration_ms ?? meta.wall_ms ?? null;
      row.outTokens = meta.output_tokens ?? meta.usage?.output_tokens ?? null;
      row.cost = meta.total_cost_usd ?? null;
    }
    const gradePath = path.join(runDir, "output.grade.json");
    if (fs.existsSync(gradePath)) {
      const g = JSON.parse(fs.readFileSync(gradePath, "utf8"));
      row.pass = g.counts.pass;
      row.total = g.results.length;
      row.clean = g.counts.fail === 0;
      row.failed = g.results.filter((r) => r.status === "fail").map((r) => r.type + (r.label ? `/${r.label}/` : ""));
    }
    (groups.get(harness) ?? groups.set(harness, []).get(harness)).push(row);
  }
  if (!groups.size) continue;

  console.log(`\n== ${task} ==`);
  for (const [harness, rows] of [...groups.entries()].sort()) {
    const graded = rows.filter((r) => r.total != null);
    const passSum = graded.reduce((n, r) => n + r.pass, 0);
    const totalSum = graded.reduce((n, r) => n + r.total, 0);
    const clean = graded.filter((r) => r.clean).length;
    const cost = rows.reduce((n, r) => n + (r.cost ?? 0), 0);
    console.log(
      `  ${harness.padEnd(12)} runs=${rows.length}  clean=${clean}/${graded.length}  ` +
        `assertions=${passSum}/${totalSum}  median_dur=${fmtMs(median(rows.map((r) => r.duration).filter((x) => x != null)))}  ` +
        `median_out_tok=${median(rows.map((r) => r.outTokens).filter((x) => x != null)) ?? "—"}` +
        (cost ? `  cost=$${cost.toFixed(2)}` : ""),
    );
    for (const r of graded.filter((r) => !r.clean)) {
      console.log(`      ✗ ${r.runName}: failed ${r.failed.join(", ")}`);
    }
  }
}
console.log();
