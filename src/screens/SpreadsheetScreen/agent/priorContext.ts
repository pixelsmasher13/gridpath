import type { ChangeBatch } from "../types";

/**
 * Serialize prior turns in this session into a context block the agent
 * receives with every new prompt. Each turn is rendered as
 *
 *   User: "<prompt>"
 *   Assistant: "<agent_text or justification>" [accepted · 12 cells]
 *
 * Conversational turns (no edits) are INCLUDED — the agent's own
 * clarifying questions and the user's replies (e.g. "sure", "yes
 * confirming") need to be in context so the agent doesn't act amnesiac
 * across batches. Streaming/in-flight batches are still skipped — the
 * agent shouldn't see its own unfinished work.
 */
const MAX_BATCHES = 10;
const MAX_CHARS = 4000;
const MAX_TEXT_PER_TURN = 600;

export function buildPriorBatchesContext(batches: ChangeBatch[]): string {
  // Include any batch that's settled, regardless of whether it produced
  // edits or was accepted. "pending" with zero mutations is the typical
  // shape of a conversational turn (agent asked a clarifying question
  // and used no tools), and those MUST flow into context.
  const settled = batches.filter((b) => b.status !== "streaming");
  if (settled.length === 0) return "";
  const recent = settled.slice(-MAX_BATCHES);

  const blocks: string[] = [];
  for (const b of recent) {
    // Prefer the agent's actual prose for context — that's where the
    // back-and-forth lives. Fall back to justification (which is the
    // summary the agent attached to a `done` tool call after edits).
    const reply = (b.agent_text?.trim() || b.justification?.trim() || "").slice(0, MAX_TEXT_PER_TURN);
    const cells = b.mutations.length;
    let status: string;
    if (b.status === "accepted") status = `accepted · ${cells} cell${cells === 1 ? "" : "s"}`;
    else if (b.status === "rejected") status = `rejected · ${cells} cell${cells === 1 ? "" : "s"} rolled back`;
    else if (cells > 0) status = `pending review · ${cells} cell${cells === 1 ? "" : "s"}`;
    else status = `no edits`;

    const userLine = `User: ${truncate(b.prompt, 240)}`;
    const replyLine = reply ? `Assistant: ${reply}` : "Assistant: (no reply text recorded)";
    blocks.push(`${userLine}\n${replyLine} [${status}]`);
  }

  let joined = blocks.join("\n\n");
  if (joined.length > MAX_CHARS) {
    // Drop oldest blocks until we fit, rather than mid-sentence truncate.
    while (joined.length > MAX_CHARS && blocks.length > 1) {
      blocks.shift();
      joined = blocks.join("\n\n");
    }
    if (joined.length > MAX_CHARS) {
      joined = joined.slice(0, MAX_CHARS - 1) + "…";
    }
  }
  return joined;
}

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return s.slice(0, n - 1) + "…";
}
