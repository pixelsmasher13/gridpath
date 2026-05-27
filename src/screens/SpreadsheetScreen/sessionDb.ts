import { invoke } from "@tauri-apps/api/core";
import type { ChangeBatch } from "./types";

export type SessionRow = {
  id: string;
  name: string;
  workbook_path: string;
  created_at: string;
  updated_at: string;
  last_opened_at: string;
  archived: number;
  /** Lifetime token totals persisted across app restarts (added 2026-05-25). */
  total_input_tokens?: number;
  total_output_tokens?: number;
  total_cache_read_tokens?: number;
  total_cache_creation_tokens?: number;
};

export type MessageRow = {
  id: number;
  session_id: string;
  role: "user" | "agent_text" | "agent_batch";
  payload: string; // stringified JSON; shape depends on role
  created_at: string;
};

export type UserPayload = { prompt: string };
export type AgentTextPayload = { text: string };
export type AgentBatchPayload = { batch: ChangeBatch };

export async function upsertSession(id: string, name: string, workbookPath: string): Promise<void> {
  await invoke("spreadsheet_session_upsert", { id, name, workbookPath });
}

export async function renameSession(id: string, name: string): Promise<void> {
  await invoke("spreadsheet_session_rename", { id, name });
}

export async function archiveSession(id: string): Promise<void> {
  await invoke("spreadsheet_session_archive", { id });
}

export async function deleteSession(id: string): Promise<void> {
  await invoke("spreadsheet_session_delete", { id });
}

export async function listSessions(limit = 50): Promise<SessionRow[]> {
  return invoke<SessionRow[]>("spreadsheet_session_list", { limit });
}

export async function getMessages(sessionId: string): Promise<MessageRow[]> {
  return invoke<MessageRow[]>("spreadsheet_session_get_messages", { sessionId });
}

export async function appendMessage(
  sessionId: string,
  role: "user" | "agent_text" | "agent_batch",
  payload: UserPayload | AgentTextPayload | AgentBatchPayload,
): Promise<number> {
  return invoke<number>("spreadsheet_session_append_message", {
    sessionId,
    role,
    payload: JSON.stringify(payload),
  });
}

/**
 * Bump the per-session lifetime token counters. Called from the agent's
 * `done` handler so cumulative usage survives app restarts (in-memory
 * tab state alone resets when the tab closes).
 */
export async function addSessionTokens(
  sessionId: string,
  input: number,
  output: number,
  cacheRead: number,
  cacheCreation: number,
): Promise<void> {
  await invoke("spreadsheet_session_add_tokens", {
    sessionId,
    input,
    output,
    cacheRead,
    cacheCreation,
  });
}
