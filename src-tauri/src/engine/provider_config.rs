use lazy_static::lazy_static;
use std::collections::HashMap;
use std::sync::Mutex;

pub const DEFAULT_CLAUDE_MODEL: &str = "claude-sonnet-4-6";
pub const DEFAULT_OPENAI_MODEL: &str = "gpt-5";
/// Default model for the ChatGPT-subscription (Codex Responses) backend.
/// Only models accessible to a ChatGPT Plus/Pro plan via the Codex CLI work here
/// (e.g. `gpt-5-codex`, `gpt-5`). Setting an arbitrary `gpt-4o-mini` will 404.
// ChatGPT subscription accounts CANNOT use `gpt-5-codex` (or `gpt-5`); the OpenAI
// Codex /responses endpoint rejects them with HTTP 400 "model is not supported when
// using Codex with a ChatGPT account." `gpt-5.5` is the top-tier subscription model
// and confirmed working end-to-end.
pub const DEFAULT_OPENAI_CODEX_MODEL: &str = "gpt-5.5";
pub const DEFAULT_GROK_MODEL: &str = "grok-3";
pub const DEFAULT_GEMINI_MODEL: &str = "gemini-2.5-flash";

lazy_static! {
    static ref MODEL_OVERRIDES: Mutex<HashMap<String, String>> = Mutex::new(HashMap::new());
    /// Reasoning effort for the OpenAI Codex (ChatGPT subscription) backend.
    /// Valid values: "minimal" | "low" | "medium" | "high" | "xhigh".
    /// Empty string means "let the backend pick its default" (effectively medium).
    static ref OPENAI_CODEX_REASONING_EFFORT: Mutex<String> = Mutex::new(String::new());
}

/// Default reasoning effort for the OpenAI Codex backend. `medium` matches the
/// OpenAI default for `gpt-5.5` and balances latency vs. agent decision quality.
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

