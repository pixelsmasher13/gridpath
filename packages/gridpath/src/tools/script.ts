/**
 * `run_script` for Node: the same `scriptCore` the desktop app ships into a
 * Blob worker, here in a `worker_threads` worker with a hard timeout. The
 * in-process fallback in core/scriptRunner.ts has no kill switch, so it is
 * never used here.
 */
import { Worker } from "node:worker_threads";
import {
  scriptCore,
  buildScriptModelFromSnapshot,
  scriptOpsToMutations,
  SCRIPT_MAX_LOGS,
  SCRIPT_MAX_TOUCHED,
  SCRIPT_TIMEOUT_MS,
  type ScriptCoreResult,
  type ScriptExecResult,
} from "../core/scriptRunner";

const WORKER_SOURCE =
  `"use strict";\n` +
  `const { parentPort } = require("node:worker_threads");\n` +
  // No network from inside a script. (`process` stays: the worker runtime needs it.)
  `for (const k of ["fetch", "XMLHttpRequest", "WebSocket", "EventSource"]) { try { globalThis[k] = undefined; } catch {} }\n` +
  `const __core = ${scriptCore.toString()};\n` +
  `parentPort.once("message", (p) => {\n` +
  `  let r;\n` +
  `  try { r = __core(p); } catch (err) { r = { ok: false, error: String((err && err.message) || err), logs: [] }; }\n` +
  `  parentPort.postMessage(r);\n` +
  `});\n`;

export async function runScript(
  snapshot: unknown,
  code: string,
  opts?: { timeoutMs?: number; maxTouched?: number },
): Promise<ScriptExecResult> {
  const timeoutMs = opts?.timeoutMs ?? SCRIPT_TIMEOUT_MS;
  let model;
  try {
    model = buildScriptModelFromSnapshot(snapshot);
  } catch (e: unknown) {
    return { ok: false, error: e instanceof Error ? e.message : String(e), logs: [] };
  }
  const payload = { code, model, limits: { maxTouched: opts?.maxTouched ?? SCRIPT_MAX_TOUCHED, maxLogs: SCRIPT_MAX_LOGS } };
  const worker = new Worker(WORKER_SOURCE, { eval: true });
  const core = await new Promise<ScriptCoreResult>((resolve) => {
    const timer = setTimeout(() => {
      void worker.terminate();
      resolve({ ok: false, error: `script timed out after ${timeoutMs}ms (infinite loop?) — simplify it or split the work`, logs: [] });
    }, timeoutMs);
    worker.once("message", (r: ScriptCoreResult) => {
      clearTimeout(timer);
      resolve(r);
    });
    worker.once("error", (e: Error) => {
      clearTimeout(timer);
      resolve({ ok: false, error: `script worker error: ${e.message || "unknown"}`, logs: [] });
    });
    worker.postMessage(payload);
  }).finally(() => void worker.terminate());
  if (!core.ok) return { ok: false, error: core.error, logs: core.logs };
  return { ok: true, mutations: scriptOpsToMutations(core.ops), writes: core.ops.length, logs: core.logs };
}
