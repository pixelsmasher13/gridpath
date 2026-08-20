use lazy_static::lazy_static;
use std::collections::HashMap;
use std::sync::Mutex;

pub const DEFAULT_CLAUDE_MODEL: &str = "claude-opus-5";
pub const DEFAULT_OPENAI_MODEL: &str = "gpt-5";
/// Default model for the ChatGPT-subscription (Codex Responses) backend.
/// Only models accessible to a ChatGPT Plus/Pro plan via the Codex CLI work here.
/// ChatGPT subscription accounts CANNOT use `gpt-5-codex`, `gpt-5`, or the bare
/// `gpt-5.6` Sol alias — the Codex /responses endpoint rejects them with HTTP 400
/// on normal Plus/Pro. `gpt-5.6-terra` (and `gpt-5.6-luna`) are the tiers that
/// work on those plans.
pub const DEFAULT_OPENAI_CODEX_MODEL: &str = "gpt-5.6-terra";
pub const DEFAULT_GROK_MODEL: &str = "grok-3";
pub const DEFAULT_GEMINI_MODEL: &str = "gemini-2.5-flash";

lazy_static! {
    static ref MODEL_OVERRIDES: Mutex<HashMap<String, String>> = Mutex::new(HashMap::new());
    /// Reasoning effort for the OpenAI Codex (ChatGPT subscription) backend.
    /// Valid values: "minimal" | "low" | "medium" | "high" | "xhigh".
    /// Empty string means "let the backend pick its default" (effectively medium).
    static ref OPENAI_CODEX_REASONING_EFFORT: Mutex<String> = Mutex::new(String::new());
    /// Effort for the Claude backend (`output_config.effort` on the Messages
    /// API). Valid values: "low" | "medium" | "high" | "xhigh" | "max".
    /// Empty string means the default. Read per-request in claude.rs; models
    /// that don't support the parameter ignore this (see thinking_gen there).
    static ref CLAUDE_EFFORT: Mutex<String> = Mutex::new(String::new());
}

/// Default Claude effort. `medium`, deliberately below the API's own `high`
/// default: on adaptive models `high` produced 5–10 minute thinks on complex
/// builds (billed as output against the user's subscription quota), and on
/// the 4.6 family effort maps to a hard thinking budget where `medium` =
/// 4,000 tokens — the exact pre-adaptive behavior. Deep-verification turns
/// escalate per session via the StatusBar picker.
pub const DEFAULT_CLAUDE_EFFORT: &str = "medium";

pub fn set_claude_effort(effort: &str) {
    let mut slot = CLAUDE_EFFORT.lock().unwrap();
    *slot = effort.to_string();
}

/// Returns the configured Claude effort, or the default if unset.
pub fn get_claude_effort() -> String {
    let slot = CLAUDE_EFFORT.lock().unwrap();
    if slot.is_empty() {
        DEFAULT_CLAUDE_EFFORT.to_string()
    } else {
        slot.clone()
    }
}

/// Default reasoning effort for the OpenAI Codex backend. `medium` matches the
/// OpenAI default for GPT-5.x and balances latency vs. agent decision quality.
pub const DEFAULT_OPENAI_CODEX_REASONING_EFFORT: &str = "medium";

pub fn set_openai_codex_reasoning_effort(effort: &str) {
    let mut slot = OPENAI_CODEX_REASONING_EFFORT.lock().unwrap();
    *slot = effort.to_string();
}

/// Returns the configured reasoning effort, or the default if unset.
/// `gpt-5.5` clamps `minimal` -> `low`; we pass through whatever the user picks
/// and let the backend handle the mapping.
pub fn get_openai_codex_reasoning_effort() -> String {
    let slot = OPENAI_CODEX_REASONING_EFFORT.lock().unwrap();
    if slot.is_empty() {
        DEFAULT_OPENAI_CODEX_REASONING_EFFORT.to_string()
    } else {
        slot.clone()
    }
}

/// Store a model override for a provider. Empty string clears the override.
pub fn set_model(provider: &str, model: &str) {
    let mut map = MODEL_OVERRIDES.lock().unwrap();
    if model.is_empty() {
        map.remove(provider);
    } else {
        map.insert(provider.to_string(), model.to_string());
    }
}

/// Returns the active model for a provider: override if set, otherwise the hardcoded default.
pub fn get_model(provider: &str) -> String {
    let overrides = MODEL_OVERRIDES.lock().unwrap();
    if let Some(m) = overrides.get(provider) {
        if !m.is_empty() {
            return m.clone();
        }
    }
    match provider {
        "openai" => DEFAULT_OPENAI_MODEL.to_string(),
        "openai-codex" => DEFAULT_OPENAI_CODEX_MODEL.to_string(),
        "grok"   => DEFAULT_GROK_MODEL.to_string(),
        "gemini" => DEFAULT_GEMINI_MODEL.to_string(),
        _        => DEFAULT_CLAUDE_MODEL.to_string(),
    }
}

