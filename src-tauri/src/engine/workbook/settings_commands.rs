use log::info;
use tauri::AppHandle;

use crate::configuration::state::ServiceAccess;
use crate::entity::setting::Setting;
use crate::repository::settings_repository::{get_setting, insert_or_update_setting};

/// Read a single setting value by key. Returns "" if missing.
/// Used by the spreadsheet workspace settings modal to read just the keys
/// it needs (api_key_claude, api_key_claude_oauth, api_choice) without
/// pulling the entire heavy Settings struct.
#[tauri::command]
pub async fn ssws_get_setting(app_handle: AppHandle, key: String) -> Result<String, String> {
    app_handle
        .db(|db| get_setting(db, &key))
        .map(|s| s.setting_value)
        .map_err(|e| format!("get_setting failed: {}", e))
}

#[tauri::command]
pub async fn ssws_set_setting(
    app_handle: AppHandle,
    key: String,
    value: String,
) -> Result<(), String> {
    info!("ssws_set_setting: key={} (value redacted)", key);
    app_handle
        .db(|db| {
            insert_or_update_setting(
                db,
                Setting { setting_key: key.clone(), setting_value: value.clone() },
            )
        })
        .map_err(|e| format!("insert_or_update_setting failed: {}", e))
}

/// Return the effective model for a provider (override → default fallback).
/// "claude" | "openai-codex" — same labels used elsewhere in the codebase.
#[tauri::command]
pub async fn ssws_get_model(provider: String) -> Result<String, String> {
    Ok(crate::engine::provider_config::get_model(&provider))
}

/// Set the model for a provider, updating BOTH the in-memory provider_config
/// override (so the agent uses the new model on its next turn — no restart
/// needed) AND the persisted settings row (so the choice survives restart).
///
/// `provider` is "claude" or "openai-codex". `model` is the API model id
/// (e.g. "claude-sonnet-4-6-20250929" or "gpt-5.5"). Empty string clears the
/// override and falls back to the built-in default.
#[tauri::command]
pub async fn ssws_set_model(
    app_handle: AppHandle,
    provider: String,
    model: String,
) -> Result<(), String> {
    info!("ssws_set_model: provider={} model={}", provider, model);
    let key = match provider.as_str() {
        "claude" => "model_claude",
        "openai-codex" => "model_openai_codex",
        other => return Err(format!("unknown provider '{}'", other)),
    };
    // In-memory override — agent picks it up on the next turn.
    crate::engine::provider_config::set_model(&provider, &model);
    // Persisted row — survives restart.
    app_handle
        .db(|db| {
            insert_or_update_setting(
                db,
                Setting { setting_key: key.to_string(), setting_value: model.clone() },
            )
        })
        .map_err(|e| format!("insert_or_update_setting failed: {}", e))
}
