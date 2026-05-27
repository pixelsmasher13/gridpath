import { invoke } from "@tauri-apps/api/core";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";

export type AgentEvent =
  | { kind: "started"; tab_id: string; batch_id: string }
  | { kind: "text_delta"; tab_id: string; batch_id: string; delta: string }
  | {
      kind: "tool_call";
      tab_id: string;
      batch_id: string;
      tool_use_id: string;
      name: string;
      input: any;
    }
  | {
      kind: "done";
      tab_id: string;
      batch_id: string;
      stop_reason: string;
      input_tokens: number;
      output_tokens: number;
      /** Anthropic prompt-cache stats — always 0 for the Codex provider. */
      cache_read_tokens: number;
      cache_creation_tokens: number;
    }
  | { kind: "error"; tab_id: string; batch_id: string; message: string };

export type WorkbookContext = {
  path: string;
  sheets: Array<{
    name: string;
    row_count: number;
    column_count: number;
    cells_preview: string;
  }>;
  /**
   * Optional "User focus" block — built from live selection + @-mentions
   * at submit time. When present, Rust injects it into the user message
   * above the prompt so the agent knows the user is directing edits at
   * specific cells.
   */
  focus?: string;
};

export async function startAgentTurn(args: {
  tabId: string;
  batchId: string;
  prompt: string;
  workbookContext: WorkbookContext;
  priorBatchesContext?: string;
}): Promise<void> {
  await invoke("spreadsheet_agent_turn", {
    tabId: args.tabId,
    batchId: args.batchId,
    prompt: args.prompt,
    workbookContext: args.workbookContext,
    priorBatchesContext: args.priorBatchesContext ?? "",
  });
}

export async function stopAgentTurn(batchId: string): Promise<void> {
  await invoke("spreadsheet_agent_stop", { batchId });
}

/**
 * Report the result of a tool_use back to the Rust agent loop so it can
 * compose the next turn's tool_result block with evaluated cell values.
 * `content` is a JSON string the agent will see verbatim as the
 * tool_result content (Claude parses it).
 */
export async function reportToolResult(toolUseId: string, content: string): Promise<void> {
  await invoke("spreadsheet_tool_result", { toolUseId, content });
}

let listenerPromise: Promise<UnlistenFn> | null = null;
const subscribers = new Set<(ev: AgentEvent) => void>();

/**
 * Ensures a single Tauri listener is registered for the "spreadsheet:event"
 * channel and fans out every event to all subscribers. We keep one listener
 * for the whole lifetime of the SpreadsheetScreen — per-tab filtering
 * happens in the subscriber. This avoids the listener-leak class of bugs
 * that hits when a tab unmounts mid-stream.
 */
export function subscribeAgentEvents(handler: (ev: AgentEvent) => void): () => void {
  subscribers.add(handler);
  if (!listenerPromise) {
    listenerPromise = listen<AgentEvent>("spreadsheet:event", (event) => {
      for (const sub of subscribers) {
        try {
          sub(event.payload);
        } catch (e) {
          console.error("agent event handler threw:", e);
        }
      }
    });
  }
  return () => {
    subscribers.delete(handler);
  };
}
