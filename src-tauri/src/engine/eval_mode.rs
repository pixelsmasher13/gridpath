//! Eval driver plumbing — the Rust half of `gridpath` self-driving eval mode.
//!
//! The eval wrapper (`eval/run-gridpath.mjs`) launches the app binary with
//! three env vars set; the webview polls `eval_config` on startup and, when
//! they're present, drives the REAL product path end to end: open the
//! pre-created output file, force auto-accept, submit the task prompt, save
//! in place, then call `eval_finish` with run metrics. Keeping the config in
//! env (not CLI args) sidesteps Tauri arg plumbing and can't collide with
//! normal launches — no vars, no eval mode.

use serde::Serialize;

#[derive(Serialize)]
pub struct EvalConfig {
    /// Task prompt, submitted verbatim as the first (only) user message.
    pub prompt: String,
    /// Absolute path of the workbook to open. The wrapper pre-creates this
    /// as `<run_dir>/output.xlsx` (blank template for build tasks, corpus
    /// copy for edit tasks) so a normal IN-PLACE save lands exactly where
    /// the grader looks — no save dialog to automate.
    pub start_file: String,
    /// Run directory; `eval_finish` writes meta.json here.
    pub out_dir: String,
}

/// Returns the eval config when the process was launched by the eval
/// wrapper, else null. The webview calls this once on mount.
#[tauri::command]
pub fn eval_config() -> Option<EvalConfig> {
    let prompt = std::env::var("GRIDPATH_EVAL_PROMPT").ok()?;
    let start_file = std::env::var("GRIDPATH_EVAL_START_FILE").ok()?;
    let out_dir = std::env::var("GRIDPATH_EVAL_OUT_DIR").ok()?;
    if prompt.is_empty() || start_file.is_empty() || out_dir.is_empty() {
        return None;
    }
    Some(EvalConfig { prompt, start_file, out_dir })
}

/// Persist run metadata and terminate the process. Exit code 0 = the run
/// completed and saved; 1 = agent error / save failure (meta.json records
/// which). The short delay lets the webview's IPC response and console
/// output flush before the hard exit.
#[tauri::command]
pub fn eval_finish(meta_json: String, success: bool) {
    if let Ok(out_dir) = std::env::var("GRIDPATH_EVAL_OUT_DIR") {
        let path = std::path::Path::new(&out_dir).join("meta.json");
        if let Err(e) = std::fs::write(&path, &meta_json) {
            eprintln!("[eval] failed to write meta.json: {}", e);
        }
    }
    std::thread::spawn(move || {
        std::thread::sleep(std::time::Duration::from_millis(400));
        std::process::exit(if success { 0 } else { 1 });
    });
}
