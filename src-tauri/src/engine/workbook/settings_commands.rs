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

/// Return the effective effort level for a provider.
/// "claude" → output_config.effort ("low"|"medium"|"high"|"xhigh"|"max");
/// "openai-codex" → reasoning.effort ("minimal"|"low"|"medium"|"high"|"xhigh").
#[tauri::command]
pub async fn ssws_get_effort(provider: String) -> Result<String, String> {
    match provider.as_str() {
        "claude" => Ok(crate::engine::provider_config::get_claude_effort()),
        "openai-codex" => Ok(crate::engine::provider_config::get_openai_codex_reasoning_effort()),
        other => Err(format!("unknown provider '{}'", other)),
    }
}

/// Set the effort for a provider — in-memory (next turn picks it up) AND
/// persisted. Empty string clears back to the provider default. Same
/// dual-write pattern as ssws_set_model.
#[tauri::command]
pub async fn ssws_set_effort(
    app_handle: AppHandle,
    provider: String,
    effort: String,
) -> Result<(), String> {
    info!("ssws_set_effort: provider={} effort={}", provider, effort);
    let key = match provider.as_str() {
        "claude" => {
            crate::engine::provider_config::set_claude_effort(&effort);
            "claude_effort"
        }
        "openai-codex" => {
            crate::engine::provider_config::set_openai_codex_reasoning_effort(&effort);
            "openai_codex_reasoning_effort"
        }
        other => return Err(format!("unknown provider '{}'", other)),
    };
    app_handle
        .db(|db| {
            insert_or_update_setting(
                db,
                Setting { setting_key: key.to_string(), setting_value: effort.clone() },
            )
        })
        .map_err(|e| format!("insert_or_update_setting failed: {}", e))
}

/// Prove a credential works before onboarding / Settings marks setup done.
///
/// `kind`:
///   - `"claude"` / `"claude-subscription"` — `credential` must be the API key
///     or OAuth token; we fire a tiny Haiku Messages call.
///   - `"openai-codex"` — ignores `credential`; refreshes the stored ChatGPT
///     OAuth session (proves sign-in is still live).
///
/// Returns `Ok(())` on success. Errors are short and UI-safe (no full HTTP
/// bodies dumped into alerts).
#[tauri::command]
pub async fn validate_llm_credential(
    app_handle: AppHandle,
    kind: String,
    credential: Option<String>,
) -> Result<(), String> {
    match kind.as_str() {
        "claude" | "claude-subscription" => {
            let cred = credential
                .map(|s| s.chars().filter(|c| !c.is_whitespace()).collect::<String>())
                .unwrap_or_default();
            if cred.is_empty() {
                return Err("Paste a Claude credential first.".to_string());
            }
            if kind == "claude-subscription" && !cred.starts_with("sk-ant-oat01-") {
                return Err(
                    "That doesn't look like a Claude subscription token (expected sk-ant-oat01-…). \
                     Run `claude setup-token` and paste the full token."
                        .to_string(),
                );
            }
            if kind == "claude" && cred.starts_with("sk-ant-oat01-") {
                return Err(
                    "That looks like a subscription OAuth token. Switch to the Claude Pro / Max tab, \
                     or paste an API key starting with sk-ant-api03-."
                        .to_string(),
                );
            }
            info!("validate_llm_credential: probing Claude ({})", kind);
            match crate::engine::llm_providers::claude::complete_claude_brief(
                &cred,
                "claude-haiku-4-5",
                "",
                "Reply with the single word OK.",
                8,
            )
            .await
            {
                Ok(_) => {
                    info!("validate_llm_credential: Claude OK");
                    Ok(())
                }
                Err(e) => {
                    info!("validate_llm_credential: Claude failed: {}", e);
                    Err(friendly_claude_auth_error(&e))
                }
            }
        }
        "openai-codex" => {
            info!("validate_llm_credential: probing ChatGPT Codex session");
            match crate::auth::openai_codex_oauth::get_active_credentials(&app_handle).await {
                Ok(_) => {
                    info!("validate_llm_credential: Codex OK");
                    Ok(())
                }
                Err(e) => {
                    info!("validate_llm_credential: Codex failed: {}", e);
                    Err(e)
                }
            }
        }
        other => Err(format!("Unknown credential kind '{}'", other)),
    }
}

fn friendly_claude_auth_error(raw: &str) -> String {
    let lower = raw.to_lowercase();
    if lower.contains("401") || lower.contains("invalid") || lower.contains("authentication") {
        return "Claude rejected this credential (invalid or expired). \
                Re-run `claude setup-token` / check the API key and try again."
            .to_string();
    }
    if lower.contains("403") || lower.contains("permission") {
        return "Claude rejected this credential (permission denied). \
                Check that the account has API / subscription access."
            .to_string();
    }
    if lower.contains("429") || lower.contains("rate") {
        return "Claude rate-limited the test call. Wait a moment and try again.".to_string();
    }
    if lower.contains("timeout") || lower.contains("timed out") {
        return "Couldn't reach Anthropic (timeout). Check your network and try again.".to_string();
    }
    // Cap length so a giant HTML/JSON body never fills the alert.
    let trimmed = if raw.chars().count() > 180 {
        format!("{}…", raw.chars().take(180).collect::<String>())
    } else {
        raw.to_string()
    };
    format!("Couldn't verify Claude credential: {}", trimmed)
}
