use std::fs::{File, OpenOptions};
use std::io::{BufRead, BufReader, Write};
use std::path::{Path, PathBuf};

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use log::info;
use tauri::{AppHandle, Manager};

/// Anchor relative workbook identifiers (`untitled-<id>`) inside the OS app-data
/// directory so they don't leak into the current working directory. Absolute paths
/// — real user files like `/Users/x/foo.xlsx` — pass through unchanged.
fn resolve_workbook_path(app_handle: &AppHandle, workbook_path: &str) -> Result<PathBuf, String> {
    let p = Path::new(workbook_path);
    if p.is_absolute() {
        return Ok(p.to_path_buf());
    }

    let untitled_dir = app_handle
        .path()
        .app_data_dir()
        .map_err(|e| format!("app_data_dir failed: {}", e))?
        .join("untitled_sessions");

    std::fs::create_dir_all(&untitled_dir)
        .map_err(|e| format!("create untitled_sessions dir failed: {}", e))?;

    Ok(untitled_dir.join(workbook_path))
}

fn change_log_path(resolved_workbook_path: &Path) -> PathBuf {
    let mut s = resolved_workbook_path.as_os_str().to_owned();
    s.push(".changes.jsonl");
    PathBuf::from(s)
}

#[tauri::command]
pub async fn read_workbook_file(app_handle: AppHandle, path: String) -> Result<String, String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    info!("read_workbook_file: {}", resolved.display());
    let bytes = std::fs::read(&resolved).map_err(|e| format!("read failed: {}", e))?;
    Ok(BASE64.encode(bytes))
}

#[tauri::command]
pub async fn write_workbook_file(app_handle: AppHandle, path: String, bytes_b64: String) -> Result<(), String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    info!("write_workbook_file: {} ({} b64 chars)", resolved.display(), bytes_b64.len());
    let bytes = BASE64
        .decode(bytes_b64.as_bytes())
        .map_err(|e| format!("base64 decode failed: {}", e))?;

    if let Some(parent) = resolved.parent() {
        if !parent.as_os_str().is_empty() {
            std::fs::create_dir_all(parent)
                .map_err(|e| format!("create parent dir failed: {}", e))?;
        }
    }

    let mut tmp = resolved.as_os_str().to_owned();
    tmp.push(".tmp");
    let tmp = PathBuf::from(tmp);
    {
        let mut f = File::create(&tmp).map_err(|e| format!("create tmp failed: {}", e))?;
        f.write_all(&bytes).map_err(|e| format!("write tmp failed: {}", e))?;
        f.sync_all().map_err(|e| format!("sync tmp failed: {}", e))?;
    }
    std::fs::rename(&tmp, &resolved).map_err(|e| format!("rename failed: {}", e))?;
    Ok(())
}

#[tauri::command]
pub async fn append_change_batch(app_handle: AppHandle, workbook_path: String, batch_json: String) -> Result<(), String> {
    let resolved = resolve_workbook_path(&app_handle, &workbook_path)?;
    let log = change_log_path(&resolved);
    info!("append_change_batch: {}", log.display());

    if batch_json.contains('\n') {
        return Err("batch_json must be a single JSON line with no newlines".into());
    }

    let mut f = OpenOptions::new()
        .create(true)
        .append(true)
        .open(&log)
        .map_err(|e| format!("open log failed: {}", e))?;
    f.write_all(batch_json.as_bytes())
        .map_err(|e| format!("write log failed: {}", e))?;
    f.write_all(b"\n")
        .map_err(|e| format!("write newline failed: {}", e))?;
    Ok(())
}

#[tauri::command]
pub async fn read_change_log(app_handle: AppHandle, workbook_path: String) -> Result<Vec<String>, String> {
    let resolved = resolve_workbook_path(&app_handle, &workbook_path)?;
    let log = change_log_path(&resolved);
    if !log.exists() {
        return Ok(Vec::new());
    }
    let f = File::open(&log).map_err(|e| format!("open log failed: {}", e))?;
    let reader = BufReader::new(f);
    let mut out = Vec::new();
    for line in reader.lines() {
        let line = line.map_err(|e| format!("read line failed: {}", e))?;
        if !line.trim().is_empty() {
            out.push(line);
        }
    }
    Ok(out)
}
