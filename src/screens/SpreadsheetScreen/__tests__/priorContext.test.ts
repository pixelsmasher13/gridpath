import { describe, expect, it } from "vitest";
import { buildPriorBatchesContext } from "../agent/priorContext";
import type { ChangeBatch } from "../types";

function batch(over: Partial<ChangeBatch>): ChangeBatch {
  return {
    id: "b1",
    prompt: "do something",
    justification: "",
    mutations: [],
    status: "accepted",
    created_at: "2026-07-21T00:00:00Z",
    ...over,
  } as ChangeBatch;
}

const cellMutation = { type: "set_cell" } as any;

describe("buildPriorBatchesContext", () => {
  it("returns empty string with no settled batches", () => {
    expect(buildPriorBatchesContext([])).toBe("");
    expect(buildPriorBatchesContext([batch({ status: "streaming" })])).toBe("");
  });

  it("keeps the last turn's closing recommendation verbatim", () => {
    // Shaped like the real failure: mid-turn narration first, then a long
    // diagnosis whose actionable fix is in the final sentence.
    const narration = "Let me trace it to the root. ".repeat(20); // ~580 chars
    const diagnosis = "The dependency trail runs through the cash block. ".repeat(40);
    const fix = "FIX: hardcode BU803 to 1 so the #VALUE! clears.";
    const b = batch({ agent_text: narration + diagnosis + fix });
    const out = buildPriorBatchesContext([b]);
    expect(out).toContain(fix);
  });

  it("preserves the tail of older turns too (middle-out digest)", () => {
    const old = batch({
      id: "old",
      prompt: "diagnose the error",
      agent_text: "Narration first. ".repeat(40) + "END-RECOMMENDATION: set X to 1.",
    });
    const last = batch({ id: "new", prompt: "ok", agent_text: "Done." });
    const out = buildPriorBatchesContext([old, last]);
    expect(out).toContain("END-RECOMMENDATION: set X to 1.");
    // The digest is still bounded: the middle of the old turn is elided.
    expect(out).toContain("[…]");
  });

  it("labels statuses including rollback of rejected batches", () => {
    const out = buildPriorBatchesContext([
      batch({ id: "a", agent_text: "did it", mutations: [cellMutation], status: "accepted" }),
      batch({ id: "r", agent_text: "tried it", mutations: [cellMutation, cellMutation], status: "rejected" }),
      batch({ id: "q", agent_text: "a question?", status: "pending" }),
    ]);
    expect(out).toContain("[accepted · 1 cell]");
    expect(out).toContain("[rejected · 2 cells rolled back]");
    expect(out).toContain("[no edits]");
  });

  it("falls back to justification when there is no agent text", () => {
    const out = buildPriorBatchesContext([batch({ justification: "Added the summary sheet." })]);
    expect(out).toContain("Assistant: Added the summary sheet.");
  });

  it("drops oldest turns first under the total cap, never the last turn", () => {
    const filler = (id: string) =>
      batch({ id, prompt: `PROMPT-${id} ` + "p".repeat(2500), agent_text: "x".repeat(700) + ` TAIL-${id}` });
    const many = Array.from({ length: 10 }, (_, i) => filler(`old${i}`));
    const last = batch({
      id: "last",
      prompt: "and now?",
      agent_text: "y".repeat(3000) + " FINAL-ANSWER-TAIL",
    });
    const out = buildPriorBatchesContext([...many, last]);
    expect(out.length).toBeLessThanOrEqual(20000);
    expect(out).toContain("FINAL-ANSWER-TAIL");
    // The very oldest turn should be the one sacrificed.
    expect(out).not.toContain("PROMPT-old0");
  });

  it("keeps recent user prompts near-verbatim and digests older ones", () => {
    // A detailed spec pasted several turns ago must survive if it's within
    // the last VERBATIM_PROMPT_TURNS; a same-length prompt further back is
    // middle-out truncated instead of head-sliced to 240 chars.
    const spec = "Build a 3-statement model. " + "d".repeat(1500) + " END-OF-SPEC-DETAILS";
    const batches = [
      batch({ id: "ancient", prompt: "a".repeat(3000) + " ANCIENT-PROMPT-TAIL", agent_text: "ok" }),
      batch({ id: "spec", prompt: spec, agent_text: "built it" }),
      batch({ id: "tweak", prompt: "make headers blue", agent_text: "done" }),
      batch({ id: "last", prompt: "continue", agent_text: "continuing" }),
    ];
    const out = buildPriorBatchesContext(batches);
    // Recent prompt (2nd-from-last window of 3) survives in full.
    expect(out).toContain("END-OF-SPEC-DETAILS");
    expect(out).toContain("d".repeat(1500));
    // The ancient prompt is truncated middle-out but keeps its tail.
    expect(out).toContain("ANCIENT-PROMPT-TAIL");
    expect(out).not.toContain("a".repeat(3000));
  });

  it("prefers turn_summary over digested prose for older turns", () => {
    const older = batch({
      id: "older",
      prompt: "build the model",
      agent_text: "Narration. ".repeat(200),
      turn_summary: "Built DCF on 'Model' rows 5-42; WACC 8.5% in B22; FY2024A revenue $97.69B from src1.",
      mutations: [cellMutation],
    });
    const last = batch({ id: "last", prompt: "ok", agent_text: "Done." });
    const out = buildPriorBatchesContext([older, last]);
    expect(out).toContain(
      "Assistant (turn summary): Built DCF on 'Model' rows 5-42; WACC 8.5% in B22; FY2024A revenue $97.69B from src1.",
    );
    // The prose is replaced by the summary for older turns.
    expect(out).not.toContain("Narration.");
  });

  it("keeps the last turn's prose and appends its turn_summary", () => {
    const last = batch({
      id: "last",
      prompt: "build it",
      agent_text: "Here is the full explanation of what I built.",
      turn_summary: "Built revenue build on 'Model' rows 3-20.",
    });
    const out = buildPriorBatchesContext([last]);
    expect(out).toContain("Assistant: Here is the full explanation of what I built.");
    expect(out).toContain("Turn summary: Built revenue build on 'Model' rows 3-20.");
  });
});
