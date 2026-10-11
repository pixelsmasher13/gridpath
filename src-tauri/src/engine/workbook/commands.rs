use std::collections::HashMap;
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use log::{info, warn};
use once_cell::sync::Lazy;
use tauri::{AppHandle, Manager};
use xxhash_rust::xxh64::xxh64;

/// Hash of the bytes each open workbook was last read from / written to.
/// `save_workbook_patched` compares against the file's current bytes to
/// detect out-of-band modification before patching on top of it.
static WORKBOOK_HASHES: Lazy<Mutex<HashMap<PathBuf, u64>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// The bytes each workbook's frontend SESSION BASELINE corresponds to —
/// what the file held when it was opened, or what the last full-bytes
/// write put there. Surgical saves always patch on top of these, not the
/// current file: the frontend's patch is a full REPLAY of every change
/// since its baseline (change batches are not consumed by surgical saves),
/// so applying it to a file that already contains an earlier surgical save
/// would run structural ops (row inserts, sheet deletes…) twice.
static WORKBOOK_BASES: Lazy<Mutex<HashMap<PathBuf, Vec<u8>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

fn remember_hash(path: &Path, bytes: &[u8]) {
    WORKBOOK_HASHES
        .lock()
        .unwrap()
        .insert(path.to_path_buf(), xxh64(bytes, 0));
}

fn remember_base(path: &Path, bytes: &[u8]) {
    WORKBOOK_BASES
        .lock()
        .unwrap()
        .insert(path.to_path_buf(), bytes.to_vec());
}

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

#[tauri::command]
pub async fn read_workbook_file(app_handle: AppHandle, path: String) -> Result<String, String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    info!("read_workbook_file: {}", resolved.display());
    let bytes = std::fs::read(&resolved).map_err(|e| format!("read failed: {}", e))?;
    remember_hash(&resolved, &bytes);
    // Opening starts a fresh frontend session whose baseline is these bytes.
    remember_base(&resolved, &bytes);
    Ok(BASE64.encode(bytes))
}

/// Surgical save: apply a cell/format patch to the workbook AS IT IS ON DISK,
/// preserving every zip entry the patch doesn't touch. Fails with
/// "BASE_CHANGED:" when the file was modified since we last read/wrote it —
/// the frontend surfaces that instead of silently clobbering.
#[tauri::command]
pub async fn save_workbook_patched(
    app_handle: AppHandle,
    path: String,
    patch_json: String,
) -> Result<(), String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    info!(
        "save_workbook_patched: {} ({} patch chars)",
        resolved.display(),
        patch_json.len()
    );
    patch_into(&resolved, &resolved, &patch_json)
}

/// Save As, surgically: clone the SOURCE package and apply the patch to the
/// copy, writing the result to TARGET. The source file is untouched; the
/// copy keeps every part (external links, charts, media…) byte-identical.
#[tauri::command]
pub async fn save_workbook_patched_as(
    app_handle: AppHandle,
    source_path: String,
    target_path: String,
    patch_json: String,
) -> Result<(), String> {
    let source = resolve_workbook_path(&app_handle, &source_path)?;
    let target = resolve_workbook_path(&app_handle, &target_path)?;
    info!(
        "save_workbook_patched_as: {} -> {} ({} patch chars)",
        source.display(),
        target.display(),
        patch_json.len()
    );
    patch_into(&source, &target, &patch_json)
}

fn patch_into(source: &Path, target: &Path, patch_json: &str) -> Result<(), String> {
    let disk = std::fs::read(source).map_err(|e| format!("read failed: {}", e))?;
    let expected = WORKBOOK_HASHES.lock().unwrap().get(source).copied();
    match expected {
        Some(h) if h != xxh64(&disk, 0) => {
            warn!("surgical save: base hash mismatch for {}", source.display());
            return Err(format!(
                "BASE_CHANGED: {} was modified outside GridPath since it was opened",
                source.display()
            ));
        }
        // No recorded hash (e.g. app restarted): nothing to compare against;
        // proceed — the patch still applies to the current on-disk bytes.
        _ => {}
    }

    // Patch on top of the frontend's session-baseline bytes, not the
    // current file. The patch replays every change since that baseline
    // (surgical saves don't consume change batches), so applying it to a
    // file that already contains an earlier surgical save would run
    // structural ops — row inserts, sheet deletes — a second time.
    // No recorded base (app restarted between open and save): the disk
    // bytes ARE the baseline, since the frontend session restarted too.
    let base = WORKBOOK_BASES
        .lock()
        .unwrap()
        .get(source)
        .cloned()
        .unwrap_or_else(|| disk.clone());

    let out = crate::engine::workbook::xlsx_patch::apply_patch_json(&base, patch_json)
        .map_err(|e| format!("PATCH_FAILED: {}", e))?;

    write_atomic(target, &out)?;
    remember_hash(target, &out);
    // Follow-up saves may land on `target` (Save As renames the tab onto
    // it) and will replay the same session history — so the target's
    // replay base is the same load-time bytes.
    remember_base(target, &base);
    Ok(())
}

fn write_atomic(resolved: &Path, bytes: &[u8]) -> Result<(), String> {
    if let Some(parent) = resolved.parent() {
        if !parent.as_os_str().is_empty() {
            std::fs::create_dir_all(parent).map_err(|e| format!("create parent dir failed: {}", e))?;
        }
    }
    let mut tmp = resolved.as_os_str().to_owned();
    tmp.push(".tmp");
    let tmp = PathBuf::from(tmp);
    {
        let mut f = File::create(&tmp).map_err(|e| format!("create tmp failed: {}", e))?;
        f.write_all(bytes).map_err(|e| format!("write tmp failed: {}", e))?;
        f.sync_all().map_err(|e| format!("sync tmp failed: {}", e))?;
    }
    std::fs::rename(&tmp, resolved).map_err(|e| format!("rename failed: {}", e))?;
    Ok(())
}

#[tauri::command]
pub async fn write_workbook_file(app_handle: AppHandle, path: String, bytes_b64: String) -> Result<(), String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    info!("write_workbook_file: {} ({} b64 chars)", resolved.display(), bytes_b64.len());
    let bytes = BASE64
        .decode(bytes_b64.as_bytes())
        .map_err(|e| format!("base64 decode failed: {}", e))?;

    // Full-bytes .xlsx writes come from the fallback exporter (ExcelJS) —
    // the path MOST likely to emit something Excel would repair. Same
    // invariant as the surgical path: validate before a byte hits disk.
    // Non-xlsx writes (.gpsnap session snapshots) pass through.
    let is_xlsx = resolved
        .extension()
        .and_then(|e| e.to_str())
        .map_or(false, |e| e.eq_ignore_ascii_case("xlsx"));
    if is_xlsx {
        if let Err(e) = crate::engine::workbook::xlsx_patch::validate::validate_package(&bytes) {
            warn!("write_workbook_file: export failed validation: {}", e);
            return Err(format!("EXPORT_INVALID: {}", e));
        }
    }

    write_atomic(&resolved, &bytes)?;
    remember_hash(&resolved, &bytes);
    // A full-bytes write (full-export save) resets the frontend's session
    // baseline to exactly these bytes — subsequent surgical patches replay
    // on top of them.
    remember_base(&resolved, &bytes);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::engine::workbook::xlsx_patch::tests::fixture;

    fn tmp_path(name: &str) -> PathBuf {
        let mut p = std::env::temp_dir();
        p.push(format!("gridpath-cmd-test-{}-{}", std::process::id(), name));
        p
    }

    fn sheet1(bytes: &[u8]) -> String {
        use std::io::Read;
        let mut ar = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut f = ar.by_name("xl/worksheets/sheet1.xml").unwrap();
        let mut s = String::new();
        f.read_to_string(&mut s).unwrap();
        s
    }

    /// The frontend's patch REPLAYS the whole session on every save. Two
    /// consecutive saves that both carry the same insertRows op must not
    /// shift the sheet twice — the second save patches the remembered
    /// session base, not the output of the first save.
    #[test]
    fn resave_replays_against_session_base_not_prior_output() {
        let path = tmp_path("replay.xlsx");
        let base = fixture();
        std::fs::write(&path, &base).unwrap();
        // Simulate the open (read_workbook_file's bookkeeping).
        remember_hash(&path, &base);
        remember_base(&path, &base);

        // Save #1: insert one row before row 2 (0-indexed 1).
        let p1 = r#"{"version":1,"rowColOps":[{"op":"insertRows","sheet":"Alpha","before":1,"count":1}]}"#;
        patch_into(&path, &path, p1).unwrap();
        let after1 = std::fs::read(&path).unwrap();
        let s1 = sheet1(&after1);
        assert!(s1.contains(r#"<c r="A3"><v>5</v></c>"#), "row 2 shifted to 3 once: {s1}");

        // Save #2 replays the SAME op plus a new cell in live coordinates.
        let p2 = r#"{"version":1,"rowColOps":[{"op":"insertRows","sheet":"Alpha","before":1,"count":1}],"sheets":[{"name":"Alpha","cells":[{"r":1,"c":0,"v":{"t":"n","n":42}}]}]}"#;
        patch_into(&path, &path, p2).unwrap();
        let after2 = std::fs::read(&path).unwrap();
        let s2 = sheet1(&after2);
        // Still shifted exactly once — NOT r="A4".
        assert!(s2.contains(r#"<c r="A3"><v>5</v></c>"#), "double shift detected: {s2}");
        assert!(s2.contains(r#"<c r="A2"><v>42</v></c>"#), "new cell missing: {s2}");

        // A full-bytes write (full-export save) moves the session base to
        // the written bytes; the next patch applies on top of them.
        remember_hash(&path, &after2);
        remember_base(&path, &after2);
        let p3 = r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":0,"v":{"t":"n","n":7}}]}]}"#;
        patch_into(&path, &path, p3).unwrap();
        let s3 = sheet1(&std::fs::read(&path).unwrap());
        assert!(s3.contains(r#"<c r="A1"><v>7</v></c>"#), "cell patch missing: {s3}");
        assert!(s3.contains(r#"<c r="A2"><v>42</v></c>"#), "prior save's cell lost: {s3}");

        std::fs::remove_file(&path).ok();
    }

    /// Out-of-band modification still fails closed even with a session base.
    #[test]
    fn base_changed_detected_before_replay() {
        let path = tmp_path("basechanged.xlsx");
        let base = fixture();
        std::fs::write(&path, &base).unwrap();
        remember_hash(&path, &base);
        remember_base(&path, &base);

        // Someone else rewrites the file.
        let mut foreign = base.clone();
        foreign.push(0);
        std::fs::write(&path, &foreign).unwrap();

        let p = r#"{"version":1,"sheets":[{"name":"Alpha","cells":[{"r":0,"c":0,"v":{"t":"n","n":1}}]}]}"#;
        let err = patch_into(&path, &path, p).unwrap_err();
        assert!(err.starts_with("BASE_CHANGED:"), "unexpected error: {err}");

        std::fs::remove_file(&path).ok();
    }
}


// ---------------------------------------------------------------------------
// Unsaved-work journal (see src/screens/SpreadsheetScreen/journal.ts).
//
// Small JSON files under <app_data>/journals, one per open file-backed
// workbook, holding the edits made since the last save so a crash (the app
// running out of memory on a 2M-cell model, a kernel panic…) does not lose
// them. Written atomically (temp file + rename) so a crash mid-write leaves
// the previous journal intact.
// ---------------------------------------------------------------------------

fn journal_dir(app_handle: &AppHandle) -> Result<PathBuf, String> {
    let dir = app_handle
        .path()
        .app_data_dir()
        .map_err(|e| format!("app_data_dir failed: {}", e))?
        .join("journals");
    std::fs::create_dir_all(&dir).map_err(|e| format!("create journals dir failed: {}", e))?;
    Ok(dir)
}

fn journal_path(app_handle: &AppHandle, name: &str) -> Result<PathBuf, String> {
    if name.is_empty()
        || name.len() > 128
        || !name
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '-' || c == '_' || c == '.')
        || name.starts_with('.')
    {
        return Err(format!("invalid journal name: {name:?}"));
    }
    Ok(journal_dir(app_handle)?.join(name))
}

#[tauri::command]
pub async fn journal_write(app_handle: AppHandle, name: String, json: String) -> Result<(), String> {
    let path = journal_path(&app_handle, &name)?;
    let tmp = path.with_extension("json.tmp");
    std::fs::write(&tmp, json.as_bytes()).map_err(|e| format!("journal write failed: {}", e))?;
    std::fs::rename(&tmp, &path).map_err(|e| format!("journal rename failed: {}", e))?;
    Ok(())
}

#[tauri::command]
pub async fn journal_read(app_handle: AppHandle, name: String) -> Result<Option<String>, String> {
    let path = journal_path(&app_handle, &name)?;
    match std::fs::read_to_string(&path) {
        Ok(s) => Ok(Some(s)),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(None),
        Err(e) => Err(format!("journal read failed: {}", e)),
    }
}

#[tauri::command]
pub async fn journal_delete(app_handle: AppHandle, name: String) -> Result<(), String> {
    let path = journal_path(&app_handle, &name)?;
    match std::fs::remove_file(&path) {
        Ok(()) => Ok(()),
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(e) => Err(format!("journal delete failed: {}", e)),
    }
}

#[derive(serde::Serialize)]
pub struct FileStamp {
    pub size: u64,
    pub mtime_ms: u64,
}

/// Size and modification time of a workbook file: the journal's fingerprint
/// of the file it belongs to, so a journal is never replayed onto a file
/// that changed underneath it (saved from another app, replaced…).
#[tauri::command]
pub async fn workbook_file_stamp(app_handle: AppHandle, path: String) -> Result<FileStamp, String> {
    let resolved = resolve_workbook_path(&app_handle, &path)?;
    let meta = std::fs::metadata(&resolved).map_err(|e| format!("stat failed: {}", e))?;
    let mtime_ms = meta
        .modified()
        .ok()
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0);
    Ok(FileStamp { size: meta.len(), mtime_ms })
}


// ---------------------------------------------------------------------------
// Memory report. The calc engine lives in this process; the grid (Univer)
// and everything JavaScript keeps around it live in WebKit's content
// process, which macOS terminates under pressure — that, not the engine, is
// what dies on a 2M-cell workbook. The content process is not our child
// (WebKit spawns it over XPC), so we ask the WKWebView for its pid.
// ---------------------------------------------------------------------------

#[derive(serde::Serialize)]
pub struct MemoryReport {
    /// Resident size of this (Rust / engine) process, bytes.
    pub app_rss: u64,
    pub webview_pid: Option<u32>,
    /// Resident size of the WebKit content process hosting the grid, bytes.
    pub webview_rss: Option<u64>,
}

#[cfg(target_os = "macos")]
async fn webview_process_id(app_handle: &AppHandle) -> Option<u32> {
    use tauri::Manager;
    let window = app_handle.get_webview_window("main")?;
    let (tx, rx) = std::sync::mpsc::channel::<Option<u32>>();
    let dispatched = window
        .with_webview(move |pw| {
            // WKWebView's `_webProcessIdentifier` (what Safari's own
            // diagnostics use). Private, stable since 10.13.
            let obj: *mut objc2::runtime::AnyObject = pw.inner().cast();
            let pid: i32 = unsafe { objc2::msg_send![obj, _webProcessIdentifier] };
            let _ = tx.send(if pid > 0 { Some(pid as u32) } else { None });
        })
        .is_ok();
    if !dispatched {
        return None;
    }
    tauri::async_runtime::spawn_blocking(move || rx.recv_timeout(std::time::Duration::from_secs(2)).ok().flatten())
        .await
        .ok()
        .flatten()
}

#[cfg(not(target_os = "macos"))]
async fn webview_process_id(_app_handle: &AppHandle) -> Option<u32> {
    None
}

#[tauri::command]
pub async fn memory_report(app_handle: AppHandle) -> Result<MemoryReport, String> {
    use sysinfo::{Pid, ProcessesToUpdate, System};
    let webview_pid = webview_process_id(&app_handle).await;
    let app_pid = std::process::id();
    let mut pids = vec![Pid::from_u32(app_pid)];
    if let Some(p) = webview_pid {
        pids.push(Pid::from_u32(p));
    }
    // macOS: `ps` is the reliable reader here. sysinfo's per-process memory
    // from a fresh System reported 25–55 MB for a process holding a 2.3M-cell
    // model, interleaved with correct GB readings, so the monitor's series
    // was noise. ps costs ~5 ms every 10 s.
    #[cfg(target_os = "macos")]
    {
        let args: Vec<String> = pids.iter().map(|p| p.as_u32().to_string()).collect();
        if let Ok(out) = std::process::Command::new("ps").arg("-o").arg("pid=,rss=").arg("-p").arg(args.join(",")).output() {
            let mut by_pid = std::collections::HashMap::new();
            for line in String::from_utf8_lossy(&out.stdout).lines() {
                let mut it = line.split_whitespace();
                if let (Some(pid), Some(kb)) = (it.next().and_then(|s| s.parse::<u32>().ok()), it.next().and_then(|s| s.parse::<u64>().ok())) {
                    by_pid.insert(pid, kb * 1024);
                }
            }
            if by_pid.contains_key(&app_pid) {
                return Ok(MemoryReport {
                    app_rss: by_pid[&app_pid],
                    webview_pid,
                    webview_rss: webview_pid.and_then(|p| by_pid.get(&p).copied()),
                });
            }
        }
    }
    let mut sys = System::new();
    sys.refresh_processes(ProcessesToUpdate::Some(&pids), true);
    let rss = |pid: u32| sys.process(Pid::from_u32(pid)).map(|p| p.memory());
    Ok(MemoryReport {
        app_rss: rss(app_pid).unwrap_or(0),
        webview_pid,
        webview_rss: webview_pid.and_then(rss),
    })
}
