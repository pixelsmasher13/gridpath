import { invoke } from "@tauri-apps/api/core";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";

export type AgentEvent =
  | { kind: "started"; tab_id: string; batch_id: string }
  | { kind: "text_delta"; tab_id: string; batch_id: string; delta: string }
  | { kind: "reasoning"; tab_id: string; batch_id: string; delta: string }
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

export type SheetContext = {
  name: string;
  row_count: number;
  column_count: number;
  /** Used data extent in A1 ("A1:Q48"); null for an empty sheet. */
  used_range?: string | null;
  cells_preview: string;
};

/**
 * A read-only reference workbook attached to the session (another xlsx the
 * user wants the agent to compare against). `label` is the filename the
 * agent addresses it by in `read_reference` calls.
 */
export type ReferenceWorkbookContext = {
  path: string;
  label: string;
  sheets: SheetContext[];
};

export type WorkbookContext = {
  path: string;
  sheets: SheetContext[];
  /**
   * Optional "User focus" block — built from live selection + @-mentions
   * at submit time. When present, Rust injects it into the user message
   * above the prompt so the agent knows the user is directing edits at
   * specific cells.
   */
  focus?: string;
  /**
   * Changed-cells block for the base+delta capture: when the workbook has
   * been edited since the memoized preview snapshot, `sheets` still carries
   * the byte-identical (cache-hit) base preview and this field lists the
   * cells that changed since. Rust injects it into the uncached turn tail.
   */
  delta?: string;
  /**
   * One-line live-vs-file-saved divergence warning (see agent/calcHealth.ts).
   * Computed once per loaded workbook and cached, so its bytes are stable —
   * Rust renders it inside the CACHED context block.
   */
  calc_health?: string;
  /**
   * Read-only reference workbooks attached to this session. Each ships a
   * compact preview; the agent pulls detail on demand via read_reference.
   */
  references?: ReferenceWorkbookContext[];
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
 * Tell the Rust loop that this tool_call has reached the head of the per-tab
 * queue and is STARTING now — not merely that its event arrived.
 *
 * Tool calls execute serially here while Rust awaits a whole turn's tools
 * concurrently, so a tool can sit queued for far longer than it takes to run.
 * Rust bills that queue time to a separate, much larger budget; without this
 * ping it billed queue time to the tool's own execution budget, gave up on
 * anything past the first tool or two of a batch, and then discarded the real
 * results as late deliveries for the rest of the run.
 *
 * Fire-and-forget: an unmatched id is a no-op on the Rust side, and a failure
 * to ping must never take down the tool it precedes.
 */
export async function reportToolStarted(toolUseId: string): Promise<void> {
  try {
    await invoke("spreadsheet_tool_started", { toolUseId });
  } catch (e) {
    console.warn("[agent] reportToolStarted failed:", e);
  }
}

/**
 * Report the result of a tool_use back to the Rust agent loop so it can
 * compose the next turn's tool_result block with evaluated cell values.
 * `content` is a JSON string the agent will see verbatim as the
 * tool_result content (Claude parses it).
 */
export async function reportToolResult(toolUseId: string, content: string): Promise<void> {
  // Diagnostic tap: localStorage.setItem("gridpath.tapToolResults", "1")
  // logs the EXACT payload the model will see for every tool result —
  // byte-truthful, so "the agent claims it got an empty result" can be
  // checked against what was actually sent.
  try {
    if (localStorage.getItem("gridpath.tapToolResults") === "1") {
      console.log(
        `[tool-tap] ${toolUseId} (${content.length}B):`,
        content.length > 2000 ? content.slice(0, 2000) + "…" : content,
      );
    }
  } catch {
    /* no localStorage */
  }
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
