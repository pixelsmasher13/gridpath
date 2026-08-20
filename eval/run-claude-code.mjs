#!/usr/bin/env node
/**
 * Run one eval task through Claude Code headless, then grade the output.
 *
 *   node eval/run-claude-code.mjs <task-id> [--model claude-opus-4-8] [--yolo]
 *
 * Creates eval/runs/<task>/<timestamp>-cc/ containing the workdir, the
 * output workbook, timing + token usage (meta.json), and the grade report.
 *
 * --yolo passes --dangerously-skip-permissions so the run is unattended
 * (Claude Code needs Bash for python/openpyxl work). The workdir is
 * isolated, but only use this on tasks/prompts you wrote yourself.
 */

import { createRequire } from "node:module";
import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evalDir = path.dirname(fileURLToPath(import.meta.url));

const args = process.argv.slice(2);
const taskId = args.find((a) => !a.startsWith("--"));
const model = args.includes("--model") ? args[args.indexOf("--model") + 1] : "claude-opus-4-8";
const yolo = args.includes("--yolo");
if (!taskId) {
  console.error("usage: node eval/run-claude-code.mjs <task-id> [--model <id>] [--yolo]");
  process.exit(2);
}

const task = JSON.parse(fs.readFileSync(path.join(evalDir, "tasks", `${taskId}.json`), "utf8"));
const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
const runDir = path.join(evalDir, "runs", taskId, `${stamp}-cc`);
const workDir = path.join(runDir, "work");
fs.mkdirSync(workDir, { recursive: true });

// Stage the starting file (if any) and pin the output filename.
let startCopy = null;
const outputName = "output.xlsx";
if (task.start_file) {
  const src = path.join(evalDir, task.start_file);
  if (!fs.existsSync(src)) {
    console.error(`start file missing: ${src} — drop the corpus file into eval/corpus/ first`);
    process.exit(2);
  }
  fs.copyFileSync(src, path.join(workDir, outputName));
  startCopy = path.join(runDir, "original.xlsx");
  fs.copyFileSync(src, startCopy);
}

const prompt = [
  task.prompt,
  "",
  task.start_file
    ? `Work on the Excel file ./${outputName} in this directory and save your changes to the same file.`
    : `Create the result as an Excel file at ./${outputName} in this directory.`,
  // Neutral on mechanism, absolute on preservation: the old wording ("MUST
  // contain cached values — recalculate before save") induced one run to
  // satisfy it the only way openpyxl can — by REPLACING formulas with
  // values, flattening the model. Cached values help label_value grading,
  // but never at the cost of steering the harness into destroying formulas.
  "Formula cells must KEEP their formulas — never replace a formula with its computed value. " +
    "Where possible also store calculated (cached) results for formula cells without altering the formulas " +
    "(e.g. a formula-preserving recalculation pass); if that isn't possible, leave cached values absent.",
].join("\n");

console.log(`task: ${taskId} | model: ${model} | run: ${runDir}`);
const t0 = Date.now();
const cli = spawnSync(
  "claude",
  [
    "-p",
    prompt,
    "--model",
    model,
    "--output-format",
    "json",
    ...(yolo ? ["--dangerously-skip-permissions"] : []),
  ],
  { cwd: workDir, encoding: "utf8", maxBuffer: 64 * 1024 * 1024, timeout: 30 * 60 * 1000 },
);
const wallMs = Date.now() - t0;

fs.writeFileSync(path.join(runDir, "cli-stdout.json"), cli.stdout ?? "");
if (cli.stderr) fs.writeFileSync(path.join(runDir, "cli-stderr.txt"), cli.stderr);

let usage = null;
try {
  const parsed = JSON.parse(cli.stdout);
  usage = {
    num_turns: parsed.num_turns,
    duration_ms: parsed.duration_ms,
    total_cost_usd: parsed.total_cost_usd,
    usage: parsed.usage,
    is_error: parsed.is_error,
  };
} catch {
  console.warn("could not parse claude CLI JSON output — see cli-stdout.json");
}
fs.writeFileSync(
  path.join(runDir, "meta.json"),
  JSON.stringify({ task: taskId, harness: "claude-code", model, wall_ms: wallMs, ...usage }, null, 2),
);

const outPath = path.join(workDir, outputName);
if (!fs.existsSync(outPath)) {
  console.error(`\nNO OUTPUT FILE — Claude Code did not produce ${outputName}. See ${runDir}`);
  process.exit(1);
}
fs.copyFileSync(outPath, path.join(runDir, "output.xlsx"));
console.log(`output: ${path.join(runDir, "output.xlsx")} (${Math.round(wallMs / 1000)}s)`);

// Grade.
const gradeArgs = [path.join(evalDir, "grade.mjs"), taskId, path.join(runDir, "output.xlsx"), "--recalc"];
if (startCopy) gradeArgs.push("--original", startCopy);
const grade = spawnSync("node", gradeArgs, { stdio: "inherit" });
process.exit(grade.status ?? 0);
