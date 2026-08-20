/**
 * Frontend half of the self-driving eval mode (see src-tauri/src/engine/
 * eval_mode.rs and eval/run-gridpath.mjs). When the app was launched by the
 * eval wrapper, `getEvalConfig()` returns the task; SpreadsheetScreen then
 * drives the REAL product path — open, prompt, auto-accept, save in place —
 * and reports metrics through `evalFinish`, which writes meta.json and
 * exits the process.
 */
import { invoke } from "@tauri-apps/api/core";

export type EvalConfig = {
  prompt: string;
  /** Pre-created output.xlsx the driver opens and saves IN PLACE. */
  start_file: string;
  out_dir: string;
};

export type EvalMeta = {
  harness: "gridpath";
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
  status: string;
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

export async function evalFinish(meta: EvalMeta, success: boolean): Promise<void> {
  try {
    await invoke("eval_finish", { metaJson: JSON.stringify(meta, null, 2), success });
  } catch (e) {
    console.error("[eval] eval_finish failed:", e);
  }
}
