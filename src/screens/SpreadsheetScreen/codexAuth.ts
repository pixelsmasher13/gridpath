import { invoke } from "@tauri-apps/api/core";

/**
 * Thin TS wrappers over the existing Codex (ChatGPT subscription) OAuth
 * Tauri commands. Auth happens via the system browser — Rust spins up a
 * localhost callback, the user signs into ChatGPT, the access token gets
 * stored in the same auth keychain heelix_notes already uses.
 */

export type CodexStatus = {
  logged_in: boolean;
  account_id?: string | null;
  expires_ms?: number | null;
};

export async function codexLogin(): Promise<void> {
  await invoke("openai_codex_login");
}

export async function codexLogout(): Promise<void> {
  await invoke("openai_codex_logout");
}

export async function codexStatus(): Promise<CodexStatus> {
  return invoke<CodexStatus>("openai_codex_status");
}
