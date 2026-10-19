import type { ChangeBatch } from "../types";

/**
 * Serialize prior turns in this session into a context block the agent
 * receives with every new prompt. Each turn is rendered as
 *
 *   User: "<prompt>"
 *   Assistant: "<agent_text or justification>" [accepted · 12 cells]
 *
 * Two tiers:
 *  - The MOST RECENT settled turn carries its reply near-verbatim. Follow-up
 *    prompts ("lets do that", "yes, the second option") refer to whatever the
 *    agent just said — usually a recommendation at the END of a long answer —
 *    so this tier is generous and any truncation preserves the tail.
 *  - Older turns are digested to a small budget each. Truncation there is
 *    middle-out (head + tail) for the same reason: agent_text accumulates
 *    mid-turn narration BEFORE the final answer, so a plain head-slice spends
 *    the whole budget on "Let me trace…" blips and drops the conclusion.
 *    (That head-slice was exactly how a proposed fix vanished from context
 *    and the agent re-derived it from scratch.)
 *
 * Replies prefer the agent-authored `turn_summary` (the structured "what I
 * did / key decisions / key figures" record it attaches to `done`) over a
 * truncated slice of its streaming prose — the summary was WRITTEN to
 * survive digestion, prose wasn't. The last turn keeps its prose
 * near-verbatim (short follow-ups refer to it) and carries the summary
 * alongside.
 *
 * User prompts: the last few turns keep the prompt near-verbatim — pasted
 * instructions and data in a recent prompt are usually the highest-value
 * tokens in the whole history ("continue" only means something if turn 1's
 * spec survived). Older prompts get a generous middle-out cap.
 *
 * Conversational turns (no edits) are INCLUDED — the agent's own
 * clarifying questions and the user's replies (e.g. "sure", "yes
 * confirming") need to be in context so the agent doesn't act amnesiac
 * across batches. Streaming/in-flight batches are still skipped — the
 * agent shouldn't see its own unfinished work.
 */
const MAX_BATCHES = 10;
const MAX_CHARS = 20000;
/** Older turns: per-turn reply digest budget, split head/tail. */
const DIGEST_HEAD = 200;
const DIGEST_TAIL = 380;
/** Most recent turn: near-verbatim reply budget, tail-heavy when over. */
const LAST_TURN_HEAD = 1000;
const LAST_TURN_TAIL = 2900;
/** How many of the most recent turns keep their user prompt near-verbatim. */
const VERBATIM_PROMPT_TURNS = 3;
/** Near-verbatim prompt cap for recent turns (middle-out when over). */
const PROMPT_RECENT_HEAD = 2500;
const PROMPT_RECENT_TAIL = 3500;
/** Older turns' prompt cap (middle-out). */
const PROMPT_HEAD = 800;
const PROMPT_TAIL = 1200;
/** Cap for the agent-authored turn_summary (already written to be compact). */
const SUMMARY_HEAD = 700;
const SUMMARY_TAIL = 1300;

export function buildPriorBatchesContext(batches: ChangeBatch[]): string {
  // Include any batch that's settled, regardless of whether it produced
  // edits or was accepted. "pending" with zero mutations is the typical
  // shape of a conversational turn (agent asked a clarifying question
  // and used no tools), and those MUST flow into context.
  const settled = batches.filter((b) => b.status !== "streaming");
  if (settled.length === 0) return "";
  const recent = settled.slice(-MAX_BATCHES);

  const blocks: string[] = [];
  recent.forEach((b, i) => {
    const isLast = i === recent.length - 1;
    const isRecent = i >= recent.length - VERBATIM_PROMPT_TURNS;
    const summary = b.turn_summary?.trim() || "";
    // Prefer the agent's actual prose for context — that's where the
    // back-and-forth lives. Fall back to justification (which is the
    // summary the agent attached to a `done` tool call after edits).
    const raw = b.agent_text?.trim() || b.justification?.trim() || "";
    const cells = b.mutations.length;
    let status: string;
    if (b.status === "accepted") status = `accepted · ${cells} cell${cells === 1 ? "" : "s"}`;
    else if (b.status === "rejected") status = `rejected · ${cells} cell${cells === 1 ? "" : "s"} rolled back`;
    else if (cells > 0) status = `pending review · ${cells} cell${cells === 1 ? "" : "s"}`;
    else status = `no edits`;

    const promptText = isRecent
      ? truncateMiddle(b.prompt, PROMPT_RECENT_HEAD, PROMPT_RECENT_TAIL)
      : truncateMiddle(b.prompt, PROMPT_HEAD, PROMPT_TAIL);
    const userLine = `User: ${promptText}`;

    let replyLine: string;
    if (isLast) {
      // Near-verbatim prose (short follow-ups refer to it), plus the
      // structured summary when the agent wrote one.
      const reply = truncateMiddle(raw, LAST_TURN_HEAD, LAST_TURN_TAIL);
      replyLine = reply ? `Assistant: ${reply}` : "Assistant: (no reply text recorded)";
      replyLine += ` [${status}]`;
      if (summary) {
        replyLine += `\nTurn summary: ${truncateMiddle(summary, SUMMARY_HEAD, SUMMARY_TAIL)}`;
      }
    } else if (summary) {
      // The agent wrote this specifically to survive digestion — it carries
      // decisions and sourced figures the prose slice would drop.
      replyLine = `Assistant (turn summary): ${truncateMiddle(summary, SUMMARY_HEAD, SUMMARY_TAIL)} [${status}]`;
    } else {
      const reply = truncateMiddle(raw, DIGEST_HEAD, DIGEST_TAIL);
      replyLine = reply ? `Assistant: ${reply}` : "Assistant: (no reply text recorded)";
      replyLine += ` [${status}]`;
    }
    blocks.push(`${userLine}\n${replyLine}`);
  });

  let joined = blocks.join("\n\n");
  if (joined.length > MAX_CHARS) {
    // Drop oldest blocks until we fit, rather than mid-sentence truncate.
    // The last (verbatim) block is never dropped: its own caps keep a
    // single block well under MAX_CHARS.
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

/**
 * Keep the start and the END of an over-budget reply. The end is where
 * recommendations and conclusions live, and it's what short follow-up
 * prompts refer to.
 */
function truncateMiddle(s: string, head: number, tail: number): string {
  if (s.length <= head + tail) return s;
  return `${s.slice(0, head)} […] ${s.slice(s.length - tail)}`;
}
