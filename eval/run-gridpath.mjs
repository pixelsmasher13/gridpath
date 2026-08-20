#!/usr/bin/env node
/**
 * Run one eval task through GridPath's self-driving eval mode, then grade.
 *
 *   node eval/run-gridpath.mjs <task-id> [--timeout <min>] [--binary <path>]
 *
 * Mirrors run-claude-code.mjs: creates eval/runs/<task>/<timestamp>-gridpath/
 * with output.xlsx, meta.json (model, effort, duration, tokens, batches)
 * and the grade report.
 *
 * Mechanics: the wrapper pre-creates the output file IN the run dir (blank
 * template for build tasks, corpus copy for edit tasks), then launches the
 * GridPath binary with GRIDPATH_EVAL_* env vars. The app opens that file,
 * submits the prompt, auto-accepts every batch, saves IN PLACE (so a plain
 * save lands exactly where the grader looks — no dialog automation), writes
 * meta.json and exits 0/1. See src-tauri/src/engine/eval_mode.rs and the
 * eval driver block in SpreadsheetScreen.tsx.
 *
 * Prereqs:
 *   npx vite build && (cd src-tauri && cargo build)
 *   GridPath must NOT be running (single-instance lock), and the app DB
 *   must already have credentials + the model/effort you want to measure.
 */

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const evalDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.dirname(evalDir);

const args = process.argv.slice(2);
const taskId = args.find((a) => !a.startsWith("--"));
const timeoutMin = args.includes("--timeout") ? Number(args[args.indexOf("--timeout") + 1]) : 20;
const binaryOverride = args.includes("--binary") ? args[args.indexOf("--binary") + 1] : null;
if (!taskId) {
  console.error("usage: node eval/run-gridpath.mjs <task-id> [--timeout <min>] [--binary <path>]");
  process.exit(2);
}

const task = JSON.parse(fs.readFileSync(path.join(evalDir, "tasks", `${taskId}.json`), "utf8"));
const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
const runDir = path.join(evalDir, "runs", taskId, `${stamp}-gridpath`);
fs.mkdirSync(runDir, { recursive: true });

// --- locate the binary + built frontend -----------------------------------
// RELEASE + custom-protocol ONLY (unless --binary): without the
// custom-protocol cargo feature (which `tauri build` passes for you), even
// a release binary loads build.devUrl (localhost:5173) instead of embedded
// assets and renders a blank window standalone. Assets embed at COMPILE
// time, so `npx vite build` must come before the cargo build.
const binary = binaryOverride ?? path.join(repoRoot, "src-tauri", "target", "release", "gridpath");
if (!fs.existsSync(binary)) {
  console.error(
    "release binary not found — build it:\n" +
      "  npx vite build && (cd src-tauri && cargo build --release --features custom-protocol)\n" +
      "(without --features custom-protocol the binary loads the vite dev URL → blank window)",
  );
  process.exit(2);
}
if (!binaryOverride) {
  const debugPath = path.join(repoRoot, "src-tauri", "target", "debug", "gridpath");
  if (fs.existsSync(debugPath) && fs.statSync(debugPath).mtimeMs - fs.statSync(binary).mtimeMs > 5 * 60_000) {
    console.warn(
      "WARNING: target/debug is newer than target/release — the release binary may be stale. " +
        "Rebuild: npx vite build && (cd src-tauri && cargo build --release)",
    );
  }
}
if (!fs.existsSync(path.join(repoRoot, "dist", "index.html"))) {
  console.error("dist/ missing — build the frontend first: npx vite build");
  process.exit(2);
}

// --- stage the output file --------------------------------------------------
// The app opens THIS file and saves in place; grading reads it afterwards.
const outPath = path.join(runDir, "output.xlsx");
let originalCopy = null;
if (task.start_file) {
  const src = path.join(evalDir, task.start_file);
  if (!fs.existsSync(src)) {
    console.error(`start file missing: ${src} — drop the corpus file into eval/corpus/ first`);
    process.exit(2);
  }
  fs.copyFileSync(src, outPath);
  originalCopy = path.join(runDir, "original.xlsx");
  fs.copyFileSync(src, originalCopy);
} else {
  // Build task: minimal blank workbook, generated (not committed) so the
  // repo carries no binary fixture. ExcelJS comes from the root package.
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet("Sheet1");
  await wb.xlsx.writeFile(outPath);
}

// --- launch -----------------------------------------------------------------
console.log(`task: ${taskId} | binary: ${path.relative(repoRoot, binary)} | run: ${runDir}`);
console.log(`timeout: ${timeoutMin} min — the app window will open; don't interact with it.`);
const logPath = path.join(runDir, "app.log");
const log = fs.openSync(logPath, "w");
const t0 = Date.now();
const child = spawn(binary, [], {
  env: {
    ...process.env,
    GRIDPATH_EVAL_PROMPT: task.prompt,
    GRIDPATH_EVAL_START_FILE: outPath,
    GRIDPATH_EVAL_OUT_DIR: runDir,
  },
  stdio: ["ignore", log, log],
});

const exitCode = await new Promise((resolve) => {
  const timer = setTimeout(() => {
    console.error(`\nTIMEOUT after ${timeoutMin} min — killing the app.`);
    child.kill("SIGKILL");
    resolve(124);
  }, timeoutMin * 60_000);
  child.on("exit", (code) => {
    clearTimeout(timer);
    resolve(code ?? 1);
  });
});
fs.closeSync(log);

const secs = Math.round((Date.now() - t0) / 1000);
const metaPath = path.join(runDir, "meta.json");
const meta = fs.existsSync(metaPath) ? JSON.parse(fs.readFileSync(metaPath, "utf8")) : null;
console.log(`\nexit=${exitCode} wall=${secs}s`);
if (meta) {
  console.log(
    `model=${meta.model} effort=${meta.effort} batches=${meta.batches} ` +
      `tokens in/out=${meta.input_tokens}/${meta.output_tokens} saved=${meta.saved}` +
      (meta.error ? `\nerror: ${meta.error}` : ""),
  );
} else {
  console.error(`no meta.json — the driver never finished. Tail of ${logPath}:`);
  try {
    const tail = fs.readFileSync(logPath, "utf8").split("\n").slice(-15).join("\n");
    console.error(tail);
  } catch {}
}

if (exitCode !== 0 || !fs.existsSync(outPath)) {
  console.error(`\nRUN FAILED — see ${runDir}`);
  process.exit(exitCode || 1);
}

// --- grade ------------------------------------------------------------------
const gradeArgs = [path.join(evalDir, "grade.mjs"), taskId, outPath, "--recalc"];
if (originalCopy) gradeArgs.push("--original", originalCopy);
const grade = spawnSync("node", gradeArgs, { stdio: "inherit" });
process.exit(grade.status ?? 0);
