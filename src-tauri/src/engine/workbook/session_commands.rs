use log::info;
use tauri::AppHandle;

use crate::configuration::state::ServiceAccess;
use crate::entity::spreadsheet_session::{SpreadsheetSession, SpreadsheetSessionMessage};
use crate::repository::spreadsheet_session_repository as repo;

fn map_err<E: std::fmt::Display>(e: E) -> String {
    format!("DB error: {}", e)
}

#[tauri::command]
pub async fn spreadsheet_session_upsert(
    app_handle: AppHandle,
    id: String,
    name: String,
    workbook_path: String,
) -> Result<(), String> {
    info!("session_upsert: id={} name={:?} path={}", id, name, workbook_path);
    app_handle
        .db(|db| repo::upsert_session(db, &id, &name, &workbook_path))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_rename(
    app_handle: AppHandle,
    id: String,
    name: String,
) -> Result<(), String> {
    app_handle
        .db(|db| repo::rename_session(db, &id, &name))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_archive(
    app_handle: AppHandle,
    id: String,
) -> Result<(), String> {
    app_handle
        .db(|db| repo::archive_session(db, &id))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_delete(
    app_handle: AppHandle,
    id: String,
) -> Result<(), String> {
    app_handle
        .db(|db| repo::delete_session(db, &id))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_list(
    app_handle: AppHandle,
    limit: Option<i64>,
) -> Result<Vec<SpreadsheetSession>, String> {
    let lim = limit.unwrap_or(50);
    app_handle
        .db(|db| repo::list_sessions(db, lim))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_get_messages(
    app_handle: AppHandle,
    session_id: String,
) -> Result<Vec<SpreadsheetSessionMessage>, String> {
    app_handle
        .db(|db| repo::get_messages(db, &session_id))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_append_message(
    app_handle: AppHandle,
    session_id: String,
    role: String,
    payload: String,
) -> Result<i64, String> {
    app_handle
        .db(|db| repo::append_message(db, &session_id, &role, &payload))
        .map_err(map_err)
}

#[tauri::command]
pub async fn spreadsheet_session_add_tokens(
    app_handle: AppHandle,
    session_id: String,
    input: i64,
    output: i64,
    cache_read: i64,
    cache_creation: i64,
) -> Result<(), String> {
    app_handle
        .db(|db| repo::add_session_tokens(db, &session_id, input, output, cache_read, cache_creation))
        .map_err(map_err)
}
