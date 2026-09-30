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
