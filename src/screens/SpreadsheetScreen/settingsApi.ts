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
 * Read the effective effort level for a provider (configured value, or the
 * provider default when unset). Wraps provider_config on the Rust side.
 */
export async function getEffort(provider: Provider): Promise<string> {
  try {
    return await invoke<string>("ssws_get_effort", { provider });
  } catch (e) {
    console.warn("[settings] get effort failed:", provider, e);
    return "";
  }
}

/**
 * Update the effort for a provider. Takes effect on the agent's next turn.
 * Empty string clears back to the provider default.
 */
export async function setEffort(provider: Provider, effort: string): Promise<void> {
  await invoke("ssws_set_effort", { provider, effort });
}

export type EffortPreset = {
  /** Raw provider value ("low", "high", "xhigh", …). */
  id: string;
  label: string;
  hint: string;
};

/**
 * The three-way effort toggle, mapped to each provider's native scale.
 * Claude: `output_config.effort` on the newest models, mapped to a hard
 *   `budget_tokens` ceiling on the 4.6 family (see thinking_and_output_config
 *   in claude.rs). App default is "medium" — deliberately below the API's own
 *   "high", which produced multi-minute thinks on complex builds.
 * Codex:  `reasoning.effort` (default "medium" — the OpenAI default).
 * The middle option is always the APP default so an untouched setting
 * changes nothing.
 */
export const EFFORT_PRESETS: Record<Provider, EffortPreset[]> = {
  claude: [
    { id: "low", label: "Fast", hint: "Quick edits, formatting, small fixes" },
    { id: "medium", label: "Standard", hint: "Balanced — the app default" },
    { id: "xhigh", label: "Thorough", hint: "Full model builds, multi-sheet work — thinks much longer" },
  ],
  "openai-codex": [
    { id: "low", label: "Fast", hint: "Quick edits, formatting, small fixes" },
    { id: "medium", label: "Standard", hint: "Balanced — the model's default" },
    { id: "high", label: "Thorough", hint: "Full model builds, multi-sheet work" },
  ],
};

/** One entry in the status-bar / Settings popular-model shortlist. */
export type ModelPreset = {
  /** Exact provider API model id. */
  id: string;
  /** Short label shown in the picker. */
  label: string;
  /**
   * Marks the app's recommended default for this provider. Should match
   * the Rust `DEFAULT_*_MODEL` so "Recommended" and blank Settings agree.
   */
  recommended?: boolean;
};

/**
 * Popular models surfaced in the status-bar picker and Settings datalist.
 * Free-form text in Settings still accepts any provider-supported id.
 *
 * Keep this list short — it's a one-click shortlist, not a catalog.
 */
export const MODEL_PRESETS: Record<Provider, ModelPreset[]> = {
  claude: [
    { id: "claude-opus-5", label: "Opus 5", recommended: true },
    { id: "claude-sonnet-4-6", label: "Sonnet 4.6" },
    { id: "claude-sonnet-5", label: "Sonnet 5" },
    { id: "claude-fable-5", label: "Fable 5" },
    { id: "claude-haiku-4-5", label: "Haiku 4.5" },
  ],
  "openai-codex": [
    // Bare `gpt-5.6` (Sol) is rejected on normal ChatGPT Plus/Pro via Codex —
    // Terra/Luna are the tiers that work on those plans.
    { id: "gpt-5.6-terra", label: "GPT-5.6 Terra", recommended: true },
    { id: "gpt-5.6-luna", label: "GPT-5.6 Luna" },
    { id: "gpt-5.5", label: "GPT-5.5" },
  ],
};

/** Recommended / default model id for a provider (first recommended preset). */
export function defaultModelId(provider: Provider): string {
  const hit = MODEL_PRESETS[provider].find((p) => p.recommended);
  return hit?.id ?? MODEL_PRESETS[provider][0]?.id ?? "";
}

/**
 * Compress a model string into a status-bar-friendly label.
 *   claude-sonnet-4-6       → "Sonnet 4.6"
 *   claude-opus-5           → "Opus 5"
 *   claude-fable-5          → "Fable 5"
 *   gpt-5.6                 → "GPT-5.6"
 *   gpt-5.6-terra           → "GPT-5.6 Terra"
 */
export function humanizeModel(raw: string, apiChoice?: string): string {
  if (!raw) {
    if (apiChoice === "openai-codex") return "ChatGPT";
    if (apiChoice === "claude-subscription") return "Claude Pro";
    if (apiChoice === "claude") return "Claude";
    return "";
  }
  // Prefer the shortlist label when we know the id.
  for (const presets of Object.values(MODEL_PRESETS)) {
    const hit = presets.find((p) => p.id === raw);
    if (hit) return hit.label;
  }
  const claude = raw.match(/^claude-(sonnet|opus|haiku|fable)-(\d+(?:[-.]\d+)?)/i);
  if (claude) {
    const tier = claude[1][0].toUpperCase() + claude[1].slice(1);
    const ver = claude[2].replace(/-/g, ".");
    return `${tier} ${ver}`;
  }
  const gpt = raw.match(/^gpt-([\d.]+)(?:-(.+))?$/i);
  if (gpt) {
    const ver = gpt[1];
    const suffix = gpt[2] ? ` ${gpt[2][0].toUpperCase() + gpt[2].slice(1)}` : "";
    return `GPT-${ver}${suffix}`;
  }
  return raw;
}
