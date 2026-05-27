import { invoke } from "@tauri-apps/api/core";

/**
 * Read/write individual settings rows by key. Used by the spreadsheet
 * workspace settings modal. The Rust side is a thin pass-through over
 * the existing settings_repository — same DB row used by the legacy
 * automation flow, so credentials are shared (no duplicate config UI).
 */
export async function getSettingValue(key: string): Promise<string> {
  try {
    return await invoke<string>("ssws_get_setting", { key });
  } catch (e) {
    console.warn("[settings] get failed:", key, e);
    return "";
  }
}

export async function setSettingValue(key: string, value: string): Promise<void> {
  await invoke("ssws_set_setting", { key, value });
}

export const SETTING_KEYS = {
  apiKeyClaude: "api_key_claude",
  apiKeyClaudeOauth: "api_key_claude_oauth",
  apiChoice: "api_choice",
  /** "1" = auto-accept agent batches on done; "0" = require manual review. */
  autoApply: "ssws_auto_apply",
} as const;

export type Provider = "claude" | "openai-codex";

/**
 * Read the effective model for a provider — override if set, otherwise the
 * built-in default. Wraps provider_config::get_model() on the Rust side.
 */
export async function getModel(provider: Provider): Promise<string> {
  try {
    return await invoke<string>("ssws_get_model", { provider });
  } catch (e) {
    console.warn("[settings] get model failed:", provider, e);
    return "";
  }
}

/**
 * Update the model for a provider. Empty string clears the override and
 * falls back to the built-in default. Takes effect on the agent's next turn —
 * no restart needed (Rust updates both in-memory override + DB row).
 */
export async function setModel(provider: Provider, model: string): Promise<void> {
  await invoke("ssws_set_model", { provider, model });
}

/**
 * Common model choices we surface in the picker. Free-form text input is
 * also allowed so new models work the day they ship without a release.
 */
export const MODEL_PRESETS: Record<Provider, string[]> = {
  claude: [
    "claude-sonnet-4-6",
    "claude-sonnet-4-5-20250929",
    "claude-opus-4-7",
    "claude-haiku-4-5-20251001",
  ],
  "openai-codex": [
    "gpt-5.5",
    "gpt-5",
    "gpt-5-codex",
  ],
};
