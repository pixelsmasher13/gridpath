// OpenAI Codex (ChatGPT subscription) OAuth flow.
//
// Mirrors the protocol used by `@earendil-works/pi-ai`'s `loginOpenAICodex` / openclaw:
//   * Authorization Code + PKCE (S256)
//   * client_id is the public ChatGPT/Codex CLI client id
//   * redirect_uri MUST be exactly http://localhost:1455/auth/callback (registered server-side)
//   * Token endpoint returns access_token, refresh_token, expires_in
//   * The access_token is a JWT containing `https://api.openai.com/auth.chatgpt_account_id`,
//     which must be sent as the `chatgpt-account-id` header on every API call.
//
// We use `tauri-plugin-oauth` to host the local callback server on port 1455. The plugin
// returns the full callback URL via a closure, which we forward through a oneshot channel.

use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine as _;
use log::{error, info, warn};
use rusqlite::Connection;
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::time::Duration;
use tauri::{AppHandle, Emitter};
use tauri_plugin_oauth::{cancel as cancel_oauth_server, start_with_config, OauthConfig};
use tokio::sync::oneshot;
use url::Url;

use crate::configuration::state::ServiceAccess;
use crate::entity::setting::Setting;
use crate::repository::settings_repository::{get_setting, insert_or_update_setting};

const CLIENT_ID: &str = "app_EMoamEEZ73f0CkXaXp7hrann";
const AUTHORIZE_URL: &str = "https://auth.openai.com/oauth/authorize";
const TOKEN_URL: &str = "https://auth.openai.com/oauth/token";
const REDIRECT_URI: &str = "http://localhost:1455/auth/callback";
const SCOPE: &str = "openid profile email offline_access";
const ORIGINATOR: &str = "gridpath";
const CALLBACK_PORT: u16 = 1455;
const JWT_CLAIM_PATH: &str = "https://api.openai.com/auth";

// Settings keys for storing tokens in the existing SQLite settings table.
pub const KEY_ACCESS: &str = "openai_codex_access";
pub const KEY_REFRESH: &str = "openai_codex_refresh";
pub const KEY_EXPIRES: &str = "openai_codex_expires_ms";
pub const KEY_ACCOUNT_ID: &str = "openai_codex_account_id";

/// Result of a successful OAuth login or refresh.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CodexCredentials {
    pub access: String,
    pub refresh: String,
    /// Absolute ms-since-epoch when the access token expires.
    pub expires_ms: i64,
    pub account_id: String,
}

/// Public-facing status returned to the frontend.
#[derive(Debug, Clone, Serialize)]
pub struct CodexAuthStatus {
    pub logged_in: bool,
    pub account_id: Option<String>,
    pub expires_ms: Option<i64>,
}

#[derive(Deserialize)]
struct TokenResponse {
    access_token: Option<String>,
    refresh_token: Option<String>,
    expires_in: Option<i64>,
    error: Option<String>,
    error_description: Option<String>,
}

#[derive(Deserialize)]
struct JwtAuthClaim {
    chatgpt_account_id: Option<String>,
}

fn now_ms() -> i64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis() as i64)
        .unwrap_or(0)
}

/// Generate (verifier, challenge) PKCE pair using S256.
fn generate_pkce() -> (String, String) {
    let mut bytes = [0u8; 64];
    getrandom_bytes(&mut bytes);
    let verifier = URL_SAFE_NO_PAD.encode(bytes);

    let mut hasher = Sha256::new();
    hasher.update(verifier.as_bytes());
    let challenge = URL_SAFE_NO_PAD.encode(hasher.finalize());

    (verifier, challenge)
}

fn random_state() -> String {
    let mut bytes = [0u8; 16];
    getrandom_bytes(&mut bytes);
    bytes.iter().map(|b| format!("{:02x}", b)).collect()
}

fn getrandom_bytes(buf: &mut [u8]) {
    use rand::RngCore;
    rand::thread_rng().fill_bytes(buf);
}

/// Decode the JWT payload (no signature verification needed; we trust the source we just got it from).
fn decode_account_id(access_token: &str) -> Option<String> {
    let mut parts = access_token.split('.');
    let _header = parts.next()?;
    let payload = parts.next()?;
    let decoded = URL_SAFE_NO_PAD.decode(payload).ok()?;
    let value: serde_json::Value = serde_json::from_slice(&decoded).ok()?;
    let claim = value.get(JWT_CLAIM_PATH)?;
    let parsed: JwtAuthClaim = serde_json::from_value(claim.clone()).ok()?;
    parsed.chatgpt_account_id.filter(|s| !s.is_empty())
}

fn build_authorize_url(challenge: &str, state: &str) -> String {
    let params = [
        ("response_type", "code"),
        ("client_id", CLIENT_ID),
        ("redirect_uri", REDIRECT_URI),
        ("scope", SCOPE),
        ("code_challenge", challenge),
        ("code_challenge_method", "S256"),
        ("state", state),
        ("id_token_add_organizations", "true"),
        ("codex_cli_simplified_flow", "true"),
        ("originator", ORIGINATOR),
    ];
    let qs = params
        .iter()
        .map(|(k, v)| format!("{}={}", k, urlencoding::encode(v)))
        .collect::<Vec<_>>()
        .join("&");
    format!("{}?{}", AUTHORIZE_URL, qs)
}

/// Open the user's default browser to a URL via `tauri-plugin-opener`, which
/// wraps `ShellExecuteW` (Windows), `open` (macOS), and `xdg-open` (Linux).
///
/// We previously shelled out manually, but `explorer.exe <https-url>` on Windows
/// makes File Explorer try to treat the URL as a folder path instead of handing
/// it to the default browser. The opener plugin delegates to the OS protocol
/// handler on every platform, which is the only call that reliably hits the
/// user's default browser on Windows.
fn open_browser(url: &str) -> Result<(), String> {
    tauri_plugin_opener::open_url(url, None::<&str>)
        .map_err(|e| format!("Failed to open browser for {}: {}", url, e))
}

/// Run the full OAuth login flow: spin up the local callback server on port 1455,
/// open the browser, wait for the redirect, exchange the code, and return creds.
pub async fn login(app_handle: AppHandle) -> Result<CodexCredentials, String> {
    let (verifier, challenge) = generate_pkce();
    let state = random_state();
    let authorize_url = build_authorize_url(&challenge, &state);

    let success_html = String::from(
        "<!doctype html><html><head><title>GridPath \u{2192} ChatGPT</title>\
         <style>body{font-family:system-ui,-apple-system,sans-serif;background:#0b0d12;\
         color:#e6e8ee;display:flex;align-items:center;justify-content:center;height:100vh;\
         margin:0}main{text-align:center;max-width:480px;padding:32px}\
         h1{font-size:22px;margin-bottom:12px}p{color:#9aa3b2;line-height:1.5}\
         </style></head><body><main><h1>Signed in to ChatGPT</h1>\
         <p>You can close this window and return to GridPath.</p></main></body></html>",
    );

    let (tx, rx) = oneshot::channel::<String>();
    let tx_cell: std::sync::Mutex<Option<oneshot::Sender<String>>> = std::sync::Mutex::new(Some(tx));

    let port = start_with_config(
        OauthConfig {
            // Pin to port 1455 because the OpenAI client is registered with that exact redirect URI.
            ports: Some(vec![CALLBACK_PORT]),
            response: Some(success_html.into()),
        },
        move |received_url| {
            if let Ok(mut guard) = tx_cell.lock() {
                if let Some(sender) = guard.take() {
                    let _ = sender.send(received_url);
                }
            }
        },
    )
    .map_err(|e| {
        format!(
            "Failed to start OAuth callback server on port {}: {}. \
             Is another login already in progress, or is port {} in use?",
            CALLBACK_PORT, e, CALLBACK_PORT
        )
    })?;

    if port != CALLBACK_PORT {
        // Should never happen because we passed exactly one port, but guard regardless.
        let _ = cancel_oauth_server(port);
        return Err(format!(
            "OAuth server bound to unexpected port {} (need {})",
            port, CALLBACK_PORT
        ));
    }

    // Inform the UI that the browser is about to open (best-effort).
    let _ = app_handle.emit("openai-codex-oauth://opening-browser", &authorize_url);

    if let Err(e) = open_browser(&authorize_url) {
        warn!(
            "Could not auto-open browser for OpenAI Codex OAuth ({}). Falling back to manual.",
            e
        );
        let _ = app_handle.emit("openai-codex-oauth://manual-url", &authorize_url);
    }

    // Wait up to 5 minutes for the user to complete the flow.
    let received = match tokio::time::timeout(Duration::from_secs(300), rx).await {
        Ok(Ok(url)) => url,
        Ok(Err(_)) => {
            let _ = cancel_oauth_server(port);
            return Err("OAuth callback channel was closed before a redirect arrived".into());
        }
        Err(_) => {
            let _ = cancel_oauth_server(port);
            return Err("Timed out waiting for ChatGPT sign-in (5 minutes)".into());
        }
    };

    // Best-effort shutdown; the plugin shuts itself down once the closure returns,
    // but cancel() is idempotent and safe to call.
    let _ = cancel_oauth_server(port);

    let parsed = Url::parse(&received)
        .map_err(|e| format!("Could not parse OAuth callback URL '{}': {}", received, e))?;

    let mut code: Option<String> = None;
    let mut state_back: Option<String> = None;
    let mut err_param: Option<String> = None;
    for (k, v) in parsed.query_pairs() {
        match k.as_ref() {
            "code" => code = Some(v.into_owned()),
            "state" => state_back = Some(v.into_owned()),
            "error" | "error_description" => err_param = Some(v.into_owned()),
            _ => {}
        }
    }

    if let Some(e) = err_param {
        return Err(format!("OpenAI returned an OAuth error: {}", e));
    }
    let code = code.ok_or_else(|| "OAuth callback missing `code` query parameter".to_string())?;
    let state_back =
        state_back.ok_or_else(|| "OAuth callback missing `state` query parameter".to_string())?;
    if state_back != state {
        return Err("OAuth state mismatch (possible CSRF; ignoring callback)".into());
    }

    let creds = exchange_code(&code, &verifier).await?;
    persist(&app_handle, &creds)?;
    info!(
        "OpenAI Codex OAuth: signed in (account {}, expires in {} min)",
        &creds.account_id,
        ((creds.expires_ms - now_ms()) / 60_000).max(0)
    );
    Ok(creds)
}

async fn exchange_code(code: &str, verifier: &str) -> Result<CodexCredentials, String> {
    let client = reqwest::Client::builder()
        .timeout(Duration::from_secs(30))
        .build()
        .map_err(|e| format!("Failed to build HTTP client: {}", e))?;

    let form = [
        ("grant_type", "authorization_code"),
        ("client_id", CLIENT_ID),
        ("code", code),
        ("code_verifier", verifier),
        ("redirect_uri", REDIRECT_URI),
    ];

    let resp = client
        .post(TOKEN_URL)
        .form(&form)
        .send()
        .await
        .map_err(|e| format!("Token exchange request failed: {}", e))?;

    let status = resp.status();
    let body = resp
        .text()
        .await
        .map_err(|e| format!("Could not read token response body: {}", e))?;

    if !status.is_success() {
        return Err(format!(
            "OpenAI token exchange failed ({}): {}",
            status.as_u16(),
            body
        ));
    }

    let parsed: TokenResponse = serde_json::from_str(&body)
        .map_err(|e| format!("Token response was not valid JSON: {} (body: {})", e, body))?;

    if let Some(err) = parsed.error {
        return Err(format!(
            "OpenAI token exchange returned error: {}{}",
            err,
            parsed
                .error_description
                .map(|d| format!(" — {}", d))
                .unwrap_or_default()
        ));
    }

    let access = parsed
        .access_token
        .ok_or("Token response missing access_token")?;
    let refresh = parsed
        .refresh_token
        .ok_or("Token response missing refresh_token")?;
    let expires_in = parsed
        .expires_in
        .ok_or("Token response missing expires_in")?;

    let account_id = decode_account_id(&access).ok_or_else(|| {
        "Could not extract chatgpt_account_id from access token (malformed JWT?)".to_string()
    })?;

    Ok(CodexCredentials {
        access,
        refresh,
        expires_ms: now_ms() + expires_in.saturating_mul(1000),
        account_id,
    })
}

async fn refresh_with(refresh_token: &str) -> Result<CodexCredentials, String> {
    let client = reqwest::Client::builder()
        .timeout(Duration::from_secs(30))
        .build()
        .map_err(|e| format!("Failed to build HTTP client: {}", e))?;

    let form = [
        ("grant_type", "refresh_token"),
        ("refresh_token", refresh_token),
        ("client_id", CLIENT_ID),
    ];

    let resp = client
        .post(TOKEN_URL)
        .form(&form)
        .send()
        .await
        .map_err(|e| format!("Token refresh request failed: {}", e))?;

    let status = resp.status();
    let body = resp
        .text()
        .await
        .map_err(|e| format!("Could not read token refresh body: {}", e))?;

    if !status.is_success() {
        return Err(format!(
            "OpenAI token refresh failed ({}): {}",
            status.as_u16(),
            body
        ));
    }

    let parsed: TokenResponse = serde_json::from_str(&body)
        .map_err(|e| format!("Refresh response was not valid JSON: {} ({})", e, body))?;
    let access = parsed
        .access_token
        .ok_or("Refresh response missing access_token")?;
    let refresh = parsed
        .refresh_token
        .ok_or("Refresh response missing refresh_token")?;
    let expires_in = parsed
        .expires_in
        .ok_or("Refresh response missing expires_in")?;
    let account_id = decode_account_id(&access)
        .ok_or_else(|| "Refreshed token had no chatgpt_account_id claim".to_string())?;
    Ok(CodexCredentials {
        access,
        refresh,
        expires_ms: now_ms() + expires_in.saturating_mul(1000),
        account_id,
    })
}

fn read(db: &Connection, key: &str) -> String {
    get_setting(db, key)
        .map(|s| s.setting_value)
        .unwrap_or_default()
}

fn write(db: &Connection, key: &str, value: &str) -> Result<(), String> {
    insert_or_update_setting(
        db,
        Setting {
            setting_key: key.to_string(),
            setting_value: value.to_string(),
        },
    )
    .map_err(|e| format!("Failed to persist {}: {}", key, e))
}

/// Read currently-stored credentials, if any.
pub fn load(app_handle: &AppHandle) -> Option<CodexCredentials> {
    let (access, refresh, expires_str, account_id) = app_handle.db(|db| {
        (
            read(db, KEY_ACCESS),
            read(db, KEY_REFRESH),
            read(db, KEY_EXPIRES),
            read(db, KEY_ACCOUNT_ID),
        )
    });
    if access.is_empty() || refresh.is_empty() || account_id.is_empty() {
        return None;
    }
    let expires_ms = expires_str.parse::<i64>().unwrap_or(0);
    Some(CodexCredentials {
        access,
        refresh,
        expires_ms,
        account_id,
    })
}

fn persist(app_handle: &AppHandle, creds: &CodexCredentials) -> Result<(), String> {
    app_handle.db(|db| -> Result<(), String> {
        write(db, KEY_ACCESS, &creds.access)?;
        write(db, KEY_REFRESH, &creds.refresh)?;
        write(db, KEY_EXPIRES, &creds.expires_ms.to_string())?;
        write(db, KEY_ACCOUNT_ID, &creds.account_id)?;
        Ok(())
    })
}

/// Wipe stored credentials.
pub fn clear(app_handle: &AppHandle) -> Result<(), String> {
    app_handle.db(|db| -> Result<(), String> {
        write(db, KEY_ACCESS, "")?;
        write(db, KEY_REFRESH, "")?;
        write(db, KEY_EXPIRES, "")?;
        write(db, KEY_ACCOUNT_ID, "")?;
        Ok(())
    })
}

/// Status snapshot for the UI.
pub fn status(app_handle: &AppHandle) -> CodexAuthStatus {
    match load(app_handle) {
        Some(c) => CodexAuthStatus {
            logged_in: true,
            account_id: Some(c.account_id),
            expires_ms: Some(c.expires_ms),
        },
        None => CodexAuthStatus {
            logged_in: false,
            account_id: None,
            expires_ms: None,
        },
    }
}

/// Get an access token usable right now, refreshing if it's within 60s of expiry.
/// Returns `(access_token, account_id)`.
pub async fn get_active_credentials(app_handle: &AppHandle) -> Result<(String, String), String> {
    let creds = load(app_handle).ok_or_else(|| {
        "Not signed in to ChatGPT. Open Settings \u{2192} ChatGPT (Subscription) and sign in."
            .to_string()
    })?;
    let needs_refresh = creds.expires_ms - now_ms() < 60_000;
    let active = if needs_refresh {
        info!("OpenAI Codex token near expiry; refreshing");
        let refreshed = refresh_with(&creds.refresh).await.map_err(|e| {
            error!("OpenAI Codex refresh failed: {}", e);
            format!("ChatGPT session expired and refresh failed: {}. Please sign in again.", e)
        })?;
        persist(app_handle, &refreshed)?;
        refreshed
    } else {
        creds
    };
    Ok((active.access, active.account_id))
}

// --- Tauri commands ---

#[tauri::command]
pub async fn openai_codex_login(app_handle: AppHandle) -> Result<CodexAuthStatus, String> {
    login(app_handle.clone()).await?;
    Ok(status(&app_handle))
}

#[tauri::command]
pub fn openai_codex_logout(app_handle: AppHandle) -> Result<CodexAuthStatus, String> {
    clear(&app_handle)?;
    Ok(status(&app_handle))
}

#[tauri::command]
pub fn openai_codex_status(app_handle: AppHandle) -> CodexAuthStatus {
    status(&app_handle)
}

