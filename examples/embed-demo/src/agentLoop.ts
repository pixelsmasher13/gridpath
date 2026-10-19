/**
 * The browser half of the agent loop. The server (server/agent.ts) makes one
 * Claude call per request; this side keeps the conversation, runs each
 * tool call against the live grid with `sheet.ai.run()`, and posts the
 * results back, until the model stops asking for tools.
 */
import type { Spreadsheet } from "gridpath";

type Block = { type: string; [key: string]: any };
export type Message = { role: "user" | "assistant"; content: Block[] };

export type AgentEvent =
  | { kind: "text"; text: string }
  | { kind: "tool"; name: string; ok: boolean }
  | { kind: "error"; text: string };

const MAX_TURNS = 40;

export async function ask(sheet: Spreadsheet, history: Message[], prompt: string, onEvent: (e: AgentEvent) => void): Promise<Message[]> {
  // The first request carries the workbook overview.
  const intro: Block[] = history.length === 0 ? [{ type: "text", text: `The open workbook:\n${await sheet.ai.context()}` }] : [];
  const messages: Message[] = [...history, { role: "user", content: [...intro, { type: "text", text: prompt }] }];
  for (let turn = 0; turn < MAX_TURNS; turn++) {
    const res = await fetch("/api/agent", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ messages }) });
    const out = await res.json();
    if (!res.ok) {
      onEvent({ kind: "error", text: out?.error ?? res.statusText });
      // Nothing reached the model on the first call: keep the old history.
      return turn === 0 ? history : messages;
    }
    // The assistant turn goes back exactly as it came (thinking and
    // fallback blocks included): the API checks it on the next call.
    messages.push({ role: "assistant", content: out.content });
    for (const b of out.content as Block[]) if (b.type === "text" && b.text) onEvent({ kind: "text", text: b.text });
    const calls = (out.content as Block[]).filter((b) => b.type === "tool_use");
    if (out.stop_reason !== "tool_use" || calls.length === 0) {
      if (out.stop_reason === "refusal") onEvent({ kind: "error", text: "The model declined this request." });
      if (out.stop_reason === "max_tokens") onEvent({ kind: "error", text: "The reply was cut off (max_tokens)." });
      return messages;
    }
    // In order (writes depend on each other), all results in ONE message.
    const results: Block[] = [];
    for (const call of calls) {
      const result = (await sheet.ai.run(call.name, call.input)) as { ok?: boolean } | null;
      const failed = result?.ok === false;
      results.push({ type: "tool_result", tool_use_id: call.id, content: JSON.stringify(result), ...(failed ? { is_error: true } : {}) });
      onEvent({ kind: "tool", name: call.name, ok: !failed });
    }
    messages.push({ role: "user", content: results });
  }
  onEvent({ kind: "error", text: `Stopped after ${MAX_TURNS} model calls.` });
  return messages;
}
