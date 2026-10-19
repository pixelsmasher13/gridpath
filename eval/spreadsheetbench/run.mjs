#!/usr/bin/env node
/**
 * SpreadsheetBench lane for GridPath's self-driving eval mode.
 *
 *   node eval/spreadsheetbench/run.mjs [options]
 *
 *   --run <name>          run directory name under runs/ (default: timestamp)
 *   --dataset <dir>       dataset root with dataset.json + spreadsheet/
 *                         (default: spreadsheetbench_verified_400 next to this file)
 *   --ids a,b,c           only these task ids
 *   --limit n / --offset n  slice of the (ordered) dataset
 *   --shard i/n           take every n-th task starting at i (0-based)
 *   --batch n             tasks per app launch (default 25)
 *   --task-timeout m      soft per-task budget in minutes (default 10); the
 *                         driver stops the turn, the wrapper kills the app
 *                         3 min after that if it still hasn't progressed
 *   --prompt-suffix "…"   text appended to every instruction (default: none;
 *                         the instruction is submitted verbatim)
 *   --binary <path>       app binary (default: src-tauri/target/release/gridpath)
 *   --retry-failed        clear meta.json of failed tasks so they re-run
 *   --no-grade            skip grade.py at the end
 *   --dry-run             stage + print the plan, don't launch
 *
 * Mechanics: for every selected task the wrapper stages the benchmark input
 * as runs/<run>/<id>/output.xlsx (the app saves IN PLACE), then launches the
 * app with GRIDPATH_EVAL_MANIFEST pointing at a JSON list of up to --batch
 * tasks. The app opens each file, submits the instruction, auto-accepts the
 * batches, saves, writes <id>/meta.json, closes the tab and moves on;
 * startup cost is paid once per batch. The wrapper watches meta.json
 * arrivals as a heartbeat, kills a stuck app, records the in-flight task as
 * failed, and relaunches with what's left. Re-running the same --run name
 * resumes: tasks with a meta.json are skipped.
 *
 * Prereqs (same as eval/run-gridpath.mjs):
 *   npx vite build && (cd src-tauri && cargo build --release --features custom-protocol)
 *   GridPath must NOT be running (single-instance lock); the app DB must have
 *   credentials + the model/effort under test; launch the release binary once
 *   by hand to get past login/onboarding; keep the screen unlocked.
 */

import { spawn, spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, "..", "..");

/** macOS lock screen suspends the webview: the driver can't progress and
 *  every deadline fires as a false failure. Detect it and pause instead. */
function isScreenLocked() {
  if (process.platform !== "darwin") return false;
  const r = spawnSync("ioreg", ["-n", "Root", "-d1", "-a"], { encoding: "utf8" });
  if (r.status !== 0) return false;
  const i = r.stdout.indexOf("CGSSessionScreenIsLocked");
  return i >= 0 && r.stdout.slice(i, i + 80).includes("<true/>");
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function waitForUnlock() {
  if (!isScreenLocked()) return;
  console.log(`  screen is locked — waiting (${new Date().toLocaleTimeString()})`);
  while (isScreenLocked()) await sleep(15_000);
  console.log(`  screen unlocked (${new Date().toLocaleTimeString()}) — continuing in 20s`);
  await sleep(20_000);
}

// --- args -------------------------------------------------------------------
const argv = process.argv.slice(2);
const opt = (name, dflt) => {
  const i = argv.indexOf(`--${name}`);
  return i >= 0 && i + 1 < argv.length ? argv[i + 1] : dflt;
};
const flag = (name) => argv.includes(`--${name}`);

const runName = opt("run", new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19));
const datasetDir = path.resolve(opt("dataset", path.join(here, "spreadsheetbench_verified_400")));
const onlyIds = opt("ids", null)?.split(",").map((s) => s.trim()).filter(Boolean) ?? null;
const limit = Number(opt("limit", 0)) || 0;
const offset = Number(opt("offset", 0)) || 0;
const shard = opt("shard", null);
const batchSize = Math.max(1, Number(opt("batch", 25)) || 25);
const taskTimeoutMin = Number(opt("task-timeout", 10)) || 10;
const promptSuffix = opt("prompt-suffix", "");
const binary = path.resolve(opt("binary", path.join(repoRoot, "src-tauri", "target", "release", "gridpath")));
const retryFailed = flag("retry-failed");
const noGrade = flag("no-grade");
const dryRun = flag("dry-run");

// --- dataset ----------------------------------------------------------------
const datasetJson = path.join(datasetDir, "dataset.json");
if (!fs.existsSync(datasetJson)) {
  console.error(`dataset.json not found under ${datasetDir}\n` +
    "download + extract a SpreadsheetBench tarball there (see README.md)");
  process.exit(2);
}
let dataset = JSON.parse(fs.readFileSync(datasetJson, "utf8")).map((d) => ({ ...d, id: String(d.id) }));

if (onlyIds) {
  const want = new Set(onlyIds);
  dataset = dataset.filter((d) => want.has(d.id));
  const missing = onlyIds.filter((id) => !dataset.some((d) => d.id === id));
  if (missing.length) console.warn(`ids not in dataset: ${missing.join(", ")}`);
}
if (shard) {
  const [i, n] = shard.split("/").map(Number);
  if (!(n > 0) || !(i >= 0) || i >= n) {
    console.error("--shard expects i/n with 0 <= i < n");
    process.exit(2);
  }
  dataset = dataset.filter((_, idx) => idx % n === i);
}
if (offset) dataset = dataset.slice(offset);
if (limit) dataset = dataset.slice(0, limit);
if (dataset.length === 0) {
  console.error("no tasks selected");
  process.exit(2);
}

/** Benchmark input for test case k. Verified 400 ships `1_<id>_init.xlsx`
 *  (5 tasks: `initial.xlsx`); the 912 set ships `<k>_<id>_input.xlsx`. */
function findInput(d, k = 1) {
  const dir = path.join(datasetDir, d.spreadsheet_path ?? `spreadsheet/${d.id}`);
  const candidates = [`${k}_${d.id}_init.xlsx`, `${k}_${d.id}_input.xlsx`, "initial.xlsx", "input.xlsx"];
  for (const c of candidates) {
    const p = path.join(dir, c);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

// --- stage run dir ------------------------------------------------------------
const runDir = path.join(here, "runs", runName);
fs.mkdirSync(runDir, { recursive: true });
const manifestsDir = path.join(runDir, "_manifests");
const logsDir = path.join(runDir, "_logs");
fs.mkdirSync(manifestsDir, { recursive: true });
fs.mkdirSync(logsDir, { recursive: true });

const runInfoPath = path.join(runDir, "run.json");
const runInfo = fs.existsSync(runInfoPath) ? JSON.parse(fs.readFileSync(runInfoPath, "utf8")) : {};
Object.assign(runInfo, {
  harness: "gridpath",
  dataset: datasetDir,
  prompt_suffix: promptSuffix,
  task_timeout_min: taskTimeoutMin,
  batch: batchSize,
  binary,
  started_at: runInfo.started_at ?? new Date().toISOString(),
  invocations: [...(runInfo.invocations ?? []), { at: new Date().toISOString(), argv }],
});
fs.writeFileSync(runInfoPath, JSON.stringify(runInfo, null, 2));

const metaPathOf = (id) => path.join(runDir, id, "meta.json");
const readMeta = (id) => {
  try {
    return JSON.parse(fs.readFileSync(metaPathOf(id), "utf8"));
  } catch {
    return null;
  }
};
const isFailed = (m) => !m || !m.saved || !!m.error;

let pending = [];
let skippedDone = 0;
let skippedMissing = 0;
for (const d of dataset) {
  const taskDir = path.join(runDir, d.id);
  const meta = readMeta(d.id);
  if (meta && !(retryFailed && isFailed(meta))) {
    skippedDone++;
    continue;
  }
  const input = findInput(d);
  if (!input) {
    console.warn(`no input file for ${d.id} — skipped`);
    skippedMissing++;
    continue;
  }
  fs.mkdirSync(taskDir, { recursive: true });
  if (meta) fs.rmSync(metaPathOf(d.id), { force: true });
  fs.copyFileSync(input, path.join(taskDir, "output.xlsx"));
  fs.writeFileSync(
    path.join(taskDir, "task.json"),
    JSON.stringify({ ...d, input_file: input }, null, 2),
  );
  pending.push({
    id: d.id,
    prompt: d.instruction + (promptSuffix ? `\n\n${promptSuffix}` : ""),
    start_file: path.join(taskDir, "output.xlsx"),
    out_dir: taskDir,
  });
}

console.log(
  `run: ${path.relative(repoRoot, runDir)} | dataset: ${path.basename(datasetDir)} | ` +
    `selected ${dataset.length}, pending ${pending.length}, already done ${skippedDone}` +
    (skippedMissing ? `, missing input ${skippedMissing}` : ""),
);
if (dryRun) {
  for (const t of pending) console.log(`  ${t.id}`);
  process.exit(0);
}
if (pending.length === 0) {
  console.log("nothing to run");
} else {
  if (!fs.existsSync(binary)) {
    console.error(
      `binary not found: ${binary}\n` +
        "  npx vite build && (cd src-tauri && cargo build --release --features custom-protocol)",
    );
    process.exit(2);
  }
  if (!fs.existsSync(path.join(repoRoot, "dist", "index.html"))) {
    console.error("dist/ missing — build the frontend first: npx vite build");
    process.exit(2);
  }
  console.log(
    `batch ${batchSize}/launch, task budget ${taskTimeoutMin} min — the app window will open; ` +
      "don't interact with it, don't lock the screen.",
  );
}

// --- batch loop -------------------------------------------------------------
const t0 = Date.now();
// Per-invocation prefix so a resume never overwrites an earlier launch's
// manifest/log (batch numbering restarts at 1 each time).
const launchStamp = new Date().toISOString().replace(/[:.]/g, "-").slice(5, 19);
let batchNo = 0;
let zeroProgressBatches = 0;
let completed = 0;
let failed = 0;

while (pending.length > 0) {
  batchNo++;
  const batch = pending.slice(0, batchSize);
  const manifest = path.join(manifestsDir, `${launchStamp}-batch-${String(batchNo).padStart(3, "0")}.json`);
  fs.writeFileSync(manifest, JSON.stringify(batch, null, 2));
  const logPath = path.join(logsDir, `${launchStamp}-batch-${String(batchNo).padStart(3, "0")}.log`);
  const log = fs.openSync(logPath, "w");
  const tb = Date.now();
  console.log(`\n[batch ${batchNo}] ${batch.length} task(s): ${batch.map((t) => t.id).join(", ")}`);

  await waitForUnlock();
  const child = spawn(binary, [], {
    env: {
      ...process.env,
      GRIDPATH_EVAL_MANIFEST: manifest,
      GRIDPATH_EVAL_TASK_TIMEOUT_MS: String(Math.round(taskTimeoutMin * 60_000)),
    },
    stdio: ["ignore", log, log],
  });

  // Heartbeat watchdog: meta.json arrivals mark progress. The driver's own
  // soft budget fires at taskTimeoutMin; give it 3 more minutes (open +
  // save + cancel grace) before declaring the app stuck.
  const stallMs = (taskTimeoutMin + 3) * 60_000;
  const doneCount = () => batch.filter((t) => fs.existsSync(metaPathOf(t.id))).length;
  let lockKill = false;
  const exitCode = await new Promise((resolve) => {
    let lastProgress = Date.now();
    let lastCount = 0;
    const poll = setInterval(() => {
      if (isScreenLocked()) {
        console.error("  SCREEN LOCKED — killing the app; in-flight work is requeued, not failed");
        lockKill = true;
        clearInterval(poll);
        child.kill("SIGKILL");
        return;
      }
      const n = doneCount();
      if (n !== lastCount) {
        for (const t of batch.slice(lastCount, n)) {
          const m = readMeta(t.id);
          const secs = m ? Math.round(m.duration_ms / 1000) : "?";
          console.log(
            `  ${t.id}: ${isFailed(m) ? "FAIL" : "ok  "} ${secs}s` +
              (m ? ` batches=${m.batches} tok=${m.input_tokens}/${m.output_tokens}` : "") +
              (m?.timed_out ? " (timed out)" : "") +
              (m?.error ? ` — ${String(m.error).slice(0, 120)}` : ""),
          );
        }
        lastCount = n;
        lastProgress = Date.now();
      } else if (Date.now() - lastProgress > stallMs) {
        console.error(`  STALLED: no task finished in ${Math.round(stallMs / 60_000)} min — killing the app`);
        clearInterval(poll);
        child.kill("SIGKILL");
      }
    }, 5_000);
    child.on("exit", (code) => {
      clearInterval(poll);
      resolve(code ?? 1);
    });
  });
  fs.closeSync(log);

  // Settle: the driver writes meta.json before exit; give the FS a beat.
  await new Promise((r) => setTimeout(r, 500));
  if (lockKill) {
    // Anything the driver gave up on while the screen was locked ("open
    // timed out") is a false failure — drop those metas so they re-run.
    for (const t of batch) {
      const m = readMeta(t.id);
      if (m && /open timed out|webview stalled/.test(String(m.error ?? ""))) {
        fs.rmSync(metaPathOf(t.id), { force: true });
        fs.copyFileSync(JSON.parse(fs.readFileSync(path.join(t.out_dir, "task.json"), "utf8")).input_file, t.start_file);
        console.log(`  ${t.id}: requeued (stalled during lock)`);
      }
    }
  }
  const finished = batch.filter((t) => fs.existsSync(metaPathOf(t.id)));
  const unfinished = batch.filter((t) => !fs.existsSync(metaPathOf(t.id)));
  // The first unfinished task is the one the app died on — record it so a
  // crash loop can't replay it forever. The rest go back in the queue.
  // Not on a lock kill: that task did nothing wrong.
  if (unfinished.length > 0 && !lockKill) {
    const victim = unfinished[0];
    const meta = {
      harness: "gridpath",
      id: victim.id,
      prompt: victim.prompt,
      model: "unknown",
      effort: "unknown",
      duration_ms: Date.now() - tb,
      batches: 0,
      accepted_batches: 0,
      input_tokens: 0,
      output_tokens: 0,
      cache_read_tokens: 0,
      cache_creation_tokens: 0,
      saved: false,
      status: "crashed",
      timed_out: exitCode === null,
      error: `app exited before finishing (exit ${exitCode}); see ${path.relative(repoRoot, logPath)}`,
    };
    fs.writeFileSync(metaPathOf(victim.id), JSON.stringify(meta, null, 2));
    console.error(`  ${victim.id}: FAIL — ${meta.error}`);
  }
  const batchDone = finished.length + (unfinished.length && !lockKill ? 1 : 0);
  const batchFailed = batch
    .filter((t) => fs.existsSync(metaPathOf(t.id)))
    .filter((t) => isFailed(readMeta(t.id))).length;
  completed += batchDone;
  failed += batchFailed;
  // "no pending sender" = the webview delivered a tool result the loop wasn't
  // waiting for. Mostly benign (skip-readback tools such as set_format never
  // register a sender, yet the webview still replies) but a genuine late
  // delivery after a loop timeout logs the same line — surface the count.
  let toolTimeouts = 0;
  try {
    toolTimeouts = (fs.readFileSync(logPath, "utf8").match(/no pending sender/g) ?? []).length;
  } catch {}
  console.log(
    `[batch ${batchNo}] exit=${exitCode} ${Math.round((Date.now() - tb) / 1000)}s — ` +
      `${batchDone - batchFailed} ok, ${batchFailed} failed` +
      (toolTimeouts ? `, ${toolTimeouts} discarded tool result(s)` : ""),
  );

  if (finished.length === 0 && !lockKill) {
    zeroProgressBatches++;
    if (zeroProgressBatches >= 2) {
      console.error(
        "\nTwo launches in a row finished no task — the app is probably not set up " +
          "(login/onboarding, credentials, stale binary). Check the batch logs. Aborting.",
      );
      break;
    }
  } else {
    zeroProgressBatches = 0;
  }
  pending = pending.filter((t) => !fs.existsSync(metaPathOf(t.id)));
}

const secs = Math.round((Date.now() - t0) / 1000);
console.log(`\ndone: ${completed} task(s) in ${Math.floor(secs / 60)}m${secs % 60}s, ${failed} failed to complete`);

// --- grade ------------------------------------------------------------------
if (!noGrade) {
  const res = spawnSync("python3", [path.join(here, "grade.py"), runDir, "--dataset", datasetDir], {
    stdio: "inherit",
  });
  process.exit(res.status ?? 0);
}
