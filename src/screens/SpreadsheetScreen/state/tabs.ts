import { v4 as uuidv4 } from "uuid";
import type { ChangeBatch } from "../types";

export type StatusPhase = "idle" | "thinking" | "writing" | "done" | "error";

export type WorkbookTab = {
  id: string;
  path: string;
  filename: string;
  /**
   * Human-readable session name. Empty until the first prompt is submitted —
   * at that point we auto-generate one from the prompt (heelix_notes-style).
   * Users can rename by double-clicking the tab.
   */
  name: string;
  dirty: boolean;
  /** Epoch ms of the last successful save, or null if never saved this session. */
  lastSavedAt: number | null;
  batches: ChangeBatch[];
  agentRunning: boolean;
  statusPhase: StatusPhase;
  statusMessage: string;
  /** Live streaming prose from the agent for the current turn. Cleared when a new turn starts. */
  streamingText: string;
  inputTokens?: number;
  outputTokens?: number;
  /**
   * Anthropic prompt-cache stats accumulated across this session's turns.
   * `cacheReadTokens` is the cheap part — system prompt + tools schema
   * served from cache. `cacheCreationTokens` is the (one-time per ~5min
   * TTL) refresh cost. Both 0 on the Codex provider.
   */
  cacheReadTokens?: number;
  cacheCreationTokens?: number;
};

export type WorkspaceState = {
  tabs: WorkbookTab[];
  activeTabId: string | null;
};

export const initialWorkspace: WorkspaceState = {
  tabs: [],
  activeTabId: null,
};

export function newTab(path: string): WorkbookTab {
  return {
    id: uuidv4(),
    path,
    filename: path.split("/").pop() ?? path,
    name: "",
    dirty: false,
    lastSavedAt: null,
    batches: [],
    agentRunning: false,
    statusPhase: "idle",
    statusMessage: "",
    streamingText: "",
  };
}

/**
 * Heuristic: derive a short session name from the user's first prompt.
 * Trim, take leading words up to ~5–6 or 40 chars, strip trailing punct,
 * Title-Case the first letter. Cheap and predictable — no extra LLM call.
 */
export function sessionNameFromPrompt(prompt: string): string {
  const trimmed = prompt.trim().replace(/\s+/g, " ");
  if (!trimmed) return "Untitled session";
  const words = trimmed.split(" ").slice(0, 7);
  let name = words.join(" ");
  if (name.length > 42) name = name.slice(0, 42).replace(/\s+\S*$/, "") + "…";
  name = name.replace(/[.,;:!?]+$/, "");
  return name.charAt(0).toUpperCase() + name.slice(1);
}

export type TabAction =
  | { type: "open"; tab: WorkbookTab }
  | { type: "close"; tabId: string }
  | { type: "activate"; tabId: string }
  | { type: "mark_dirty"; tabId: string; dirty: boolean }
  | { type: "mark_saved"; tabId: string; at: number }
  | { type: "set_name"; tabId: string; name: string }
  | { type: "rename"; tabId: string; path: string }
  | { type: "set_status"; tabId: string; phase: StatusPhase; message: string }
  | { type: "set_tokens"; tabId: string; input: number; output: number; cacheRead: number; cacheCreation: number }
  | { type: "set_agent_running"; tabId: string; running: boolean }
  | { type: "stream_text_append"; tabId: string; delta: string }
  | { type: "stream_text_clear"; tabId: string }
  | { type: "batches_replace"; tabId: string; batches: ChangeBatch[] }
  | { type: "batch_add"; tabId: string; batch: ChangeBatch }
  | { type: "batch_append_mutation"; tabId: string; batchId: string; mutation: ChangeBatch["mutations"][number] }
  | { type: "batch_set_justification"; tabId: string; batchId: string; justification: string }
  | { type: "batch_set_agent_text"; tabId: string; batchId: string; agentText: string }
  | { type: "batch_add_fetched_urls"; tabId: string; batchId: string; urls: string[] }
  | { type: "batch_finalize"; tabId: string; batchId: string }
  | { type: "batch_accept"; tabId: string; batchId: string }
  | { type: "batch_reject"; tabId: string; batchId: string };

export function reduceWorkspace(state: WorkspaceState, action: TabAction): WorkspaceState {
  switch (action.type) {
    case "open":
      return {
        tabs: [...state.tabs, action.tab],
        activeTabId: action.tab.id,
      };
    case "close": {
      const idx = state.tabs.findIndex((t) => t.id === action.tabId);
      if (idx < 0) return state;
      const remaining = state.tabs.filter((t) => t.id !== action.tabId);
      let activeTabId = state.activeTabId;
      if (state.activeTabId === action.tabId) {
        const fallback = remaining[idx] ?? remaining[idx - 1] ?? remaining[0] ?? null;
        activeTabId = fallback?.id ?? null;
      }
      return { tabs: remaining, activeTabId };
    }
    case "activate":
      return { ...state, activeTabId: action.tabId };
    case "mark_dirty":
      return mapTab(state, action.tabId, (t) => ({ ...t, dirty: action.dirty }));
    case "mark_saved":
      return mapTab(state, action.tabId, (t) => ({ ...t, dirty: false, lastSavedAt: action.at }));
    case "set_name":
      return mapTab(state, action.tabId, (t) => ({ ...t, name: action.name }));
    case "rename":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        path: action.path,
        filename: action.path.split("/").pop() ?? action.path,
      }));
    case "set_status":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        statusPhase: action.phase,
        statusMessage: action.message,
      }));
    case "set_tokens":
      // ACCUMULATE rather than overwrite. The agent_event 'done' fires
      // once per batch with that batch's token counts; the workbook-level
      // Usage tab should show the sum across every batch in the session,
      // not just the last one. Previously a tiny follow-up batch (e.g.
      // rename_sheet, out=154) would replace a big build's totals
      // (out=5184) and make it look like the workbook only used 154
      // output tokens despite minutes of agent work.
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        inputTokens: (t.inputTokens ?? 0) + action.input,
        outputTokens: (t.outputTokens ?? 0) + action.output,
        cacheReadTokens: (t.cacheReadTokens ?? 0) + action.cacheRead,
        cacheCreationTokens: (t.cacheCreationTokens ?? 0) + action.cacheCreation,
      }));
    case "set_agent_running":
      return mapTab(state, action.tabId, (t) => ({ ...t, agentRunning: action.running }));
    case "stream_text_append":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        streamingText: t.streamingText + action.delta,
      }));
    case "stream_text_clear":
      return mapTab(state, action.tabId, (t) => ({ ...t, streamingText: "" }));
    case "batches_replace":
      return mapTab(state, action.tabId, (t) => ({ ...t, batches: action.batches }));
    case "batch_add":
      return mapTab(state, action.tabId, (t) => ({ ...t, batches: [...t.batches, action.batch] }));
    case "batch_append_mutation":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId ? { ...b, mutations: [...b.mutations, action.mutation] } : b,
        ),
      }));
    case "batch_set_justification":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId ? { ...b, justification: action.justification } : b,
        ),
      }));
    case "batch_set_agent_text":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId ? { ...b, agent_text: action.agentText } : b,
        ),
      }));
    case "batch_add_fetched_urls":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId
            ? { ...b, fetched_urls: [...(b.fetched_urls ?? []), ...action.urls] }
            : b,
        ),
      }));
    case "batch_finalize":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId && b.status === "streaming" ? { ...b, status: "pending" } : b,
        ),
      }));
    case "batch_accept":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId ? { ...b, status: "accepted" } : b,
        ),
      }));
    case "batch_reject":
      return mapTab(state, action.tabId, (t) => ({
        ...t,
        batches: t.batches.map((b) =>
          b.id === action.batchId ? { ...b, status: "rejected" } : b,
        ),
      }));
  }
}

function mapTab(
  state: WorkspaceState,
  tabId: string,
  f: (t: WorkbookTab) => WorkbookTab,
): WorkspaceState {
  return { ...state, tabs: state.tabs.map((t) => (t.id === tabId ? f(t) : t)) };
}

export function findTab(state: WorkspaceState, tabId: string | null): WorkbookTab | null {
  if (!tabId) return null;
  return state.tabs.find((t) => t.id === tabId) ?? null;
}

export function findTabByBatch(state: WorkspaceState, batchId: string): WorkbookTab | null {
  return state.tabs.find((t) => t.batches.some((b) => b.id === batchId)) ?? null;
}

export function findTabByPath(state: WorkspaceState, path: string): WorkbookTab | null {
  return state.tabs.find((t) => t.path === path) ?? null;
}
