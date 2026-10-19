//! Eval driver plumbing — the Rust half of `gridpath` self-driving eval mode.
//!
//! Two launch shapes, both env-driven so they can't collide with normal
//! launches (no vars, no eval mode):
//!
//! * **Single task** (`eval/run-gridpath.mjs`): `GRIDPATH_EVAL_PROMPT`,
//!   `GRIDPATH_EVAL_START_FILE`, `GRIDPATH_EVAL_OUT_DIR`. One task, then exit.
//! * **Manifest** (`eval/spreadsheetbench/run.mjs`): `GRIDPATH_EVAL_MANIFEST`
//!   points at a JSON array of tasks. The webview runs them back to back in
//!   one process — open, prompt, auto-accept, save in place, close tab, next
//!   — paying app startup once per batch instead of once per task.
//!
//! The webview polls `eval_config` on startup, calls `eval_task_done` after
//! each task (writes that task's meta.json) and `eval_finish` when the queue
//! is drained (exits the process).

use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Clone)]
pub struct EvalTask {
    /// Stable task id (benchmark id or eval task name); echoed into meta.json.
    #[serde(default)]
    pub id: String,
    /// Task prompt, submitted verbatim as the first (only) user message.
    pub prompt: String,
    /// Absolute path of the workbook to open. The wrapper pre-creates this
    /// as `<out_dir>/output.xlsx` (blank template for build tasks, input
    /// copy for edit/benchmark tasks) so a normal IN-PLACE save lands
    /// exactly where the grader looks — no save dialog to automate.
    pub start_file: String,
    /// Per-task directory; `eval_task_done` writes meta.json here.
    pub out_dir: String,
}

#[derive(Serialize)]
pub struct EvalConfig {
    pub tasks: Vec<EvalTask>,
    /// Soft per-task budget: after this the driver stops the agent turn and
    /// proceeds with whatever was produced. `GRIDPATH_EVAL_TASK_TIMEOUT_MS`,
    /// default 10 minutes.
    pub task_timeout_ms: u64,
}

/// Returns the eval config when the process was launched by an eval
/// wrapper, else null. The webview calls this once on mount.
#[tauri::command]
pub fn eval_config() -> Option<EvalConfig> {
    let task_timeout_ms = std::env::var("GRIDPATH_EVAL_TASK_TIMEOUT_MS")
        .ok()
        .and_then(|s| s.parse::<u64>().ok())
        .unwrap_or(10 * 60 * 1000);

    if let Ok(manifest) = std::env::var("GRIDPATH_EVAL_MANIFEST") {
        if !manifest.is_empty() {
            let raw = match std::fs::read_to_string(&manifest) {
                Ok(s) => s,
                Err(e) => {
                    eprintln!("[eval] cannot read manifest {}: {}", manifest, e);
                    return None;
                }
            };
            let tasks: Vec<EvalTask> = match serde_json::from_str(&raw) {
                Ok(t) => t,
                Err(e) => {
                    eprintln!("[eval] manifest {} is not a task array: {}", manifest, e);
                    return None;
                }
            };
            if tasks.is_empty() {
                return None;
            }
            return Some(EvalConfig { tasks, task_timeout_ms });
        }
    }

    let prompt = std::env::var("GRIDPATH_EVAL_PROMPT").ok()?;
    let start_file = std::env::var("GRIDPATH_EVAL_START_FILE").ok()?;
    let out_dir = std::env::var("GRIDPATH_EVAL_OUT_DIR").ok()?;
    if prompt.is_empty() || start_file.is_empty() || out_dir.is_empty() {
        return None;
    }
    Some(EvalConfig {
        tasks: vec![EvalTask { id: "task".to_string(), prompt, start_file, out_dir }],
        task_timeout_ms,
    })
}

/// Persist one task's metadata into its out_dir. Called after every task in
/// the queue (including the last one, before `eval_finish`).
#[tauri::command]
pub fn eval_task_done(out_dir: String, meta_json: String) -> Result<(), String> {
    let path = std::path::Path::new(&out_dir).join("meta.json");
    std::fs::write(&path, &meta_json).map_err(|e| format!("write {}: {}", path.display(), e))
}

/// Terminate the process once the queue is drained. Exit code 0 = every task
/// completed and saved; 1 = at least one task failed (its meta.json records
/// why). The short delay lets the webview's IPC response and console output
/// flush before the hard exit.
#[tauri::command]
pub fn eval_finish(success: bool) {
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_millis(400));
        std::process::exit(if success { 0 } else { 1 });
    });
}
