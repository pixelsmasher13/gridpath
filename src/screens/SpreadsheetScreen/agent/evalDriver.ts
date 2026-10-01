/**
 * Frontend half of the self-driving eval mode (see src-tauri/src/engine/
 * eval_mode.rs, eval/run-gridpath.mjs and eval/spreadsheetbench/run.mjs).
 * When the app was launched by an eval wrapper, `getEvalConfig()` returns a
 * task queue; SpreadsheetScreen then drives the REAL product path for each
 * task in turn — open, prompt, auto-accept, save in place, close — reporting
 * per-task metrics through `evalTaskDone` (writes meta.json) and finally
 * `evalFinish`, which exits the process.
 */
import { invoke } from "@tauri-apps/api/core";

export type EvalTask = {
  id: string;
  prompt: string;
  /** Pre-created output.xlsx the driver opens and saves IN PLACE. */
  start_file: string;
  out_dir: string;
};

export type EvalConfig = {
  tasks: EvalTask[];
  /** Soft per-task budget; the driver stops the agent turn when it elapses. */
  task_timeout_ms: number;
};

export type EvalMeta = {
  harness: "gridpath";
  id: string;
  prompt: string;
  model: string;
  effort: string;
  duration_ms: number;
  batches: number;
  accepted_batches: number;
  input_tokens: number;
  output_tokens: number;
  cache_read_tokens: number;
  cache_creation_tokens: number;
  saved: boolean;
  /** Why the surgical save bailed and what the gate would have done (eval
   *  mode always takes the export); null when the in-place patch saved. */
  save_note: string | null;
  status: string;
  timed_out: boolean;
  error: string | null;
};

export async function getEvalConfig(): Promise<EvalConfig | null> {
  try {
    return (await invoke<EvalConfig | null>("eval_config")) ?? null;
  } catch {
    // Command missing (old binary) or IPC failure — never block a normal
    // launch on eval plumbing.
    return null;
  }
}

export async function evalTaskDone(outDir: string, meta: EvalMeta): Promise<void> {
  try {
    await invoke("eval_task_done", { outDir, metaJson: JSON.stringify(meta, null, 2) });
  } catch (e) {
    console.error("[eval] eval_task_done failed:", e);
  }
}

export async function evalFinish(success: boolean): Promise<void> {
  try {
    await invoke("eval_finish", { success });
  } catch (e) {
    console.error("[eval] eval_finish failed:", e);
  }
}
