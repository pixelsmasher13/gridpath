# `stage_data` — sheet-as-calculator, v1 design

*Drafted 2026-08-18, out of the parity benchmark + blind-judge findings.*

## Problem

On research-grounded builds the agent keeps fetched figures in prose
(`extracted_notes`) and its own reasoning, derives intermediate values by
mental arithmetic, and writes **results** into cells as hardcodes. Four
measured costs:

1. **Wall-clock & tokens** — minutes of thinking (billed as output) doing
   subtraction a calc engine does in microseconds. Parity builds ran
   55–120k output tokens; most of it reasoning, not writing.
2. **Auditability** — judges' #1 recurring flaw in GridPath builds:
   "back-solved plugs", "disguised hardcodes". A reader can't tell derived
   from sourced.
3. **Confabulation** — numbers held in prose memory have no provenance;
   we watched the agent narrate figures from memory as if fetched.
4. **Mega-turn deaths** — planning an entire layout mentally then
   serializing one giant write exhausted the 32k (then 48k) turn ceiling.

Prompt rules cannot fix this (measured: two rule iterations made judged
quality *worse*). The environment must make the analyst workflow — data in
cells first, derive by formula, verify by readback — the cheapest path.

## v1: one new tool

```jsonc
{
  "name": "stage_data",
  "description": "Stage REPORTED figures into a data sheet the moment you
    extract them from a source — then build the model by FORMULA REFERENCE
    to the staged cells (e.g. =Data!C12) instead of retyping or mentally
    deriving numbers. The harness picks the location and returns the exact
    address of every row and column, so referencing costs nothing. Staged
    blocks carry their source and render as the workbook's data foundation;
    presentation sheets that reference them are auditable by construction.",
  "input_schema": {
    "sheet":   { "type": "string",  "description": "Staging sheet name (default \"Data\"). Auto-created." },
    "title":   { "type": "string",  "description": "Block caption, e.g. \"AAPL 10-Q Q3 FY2026 — income statement (reported)\"" },
    "source":  { "type": "string",  "description": "Provenance: URL or filing name. Rendered under the title." },
    "units":   { "type": "string",  "description": "e.g. \"$M\", \"€m\". Rendered with the source." },
    "columns": { "type": "array",   "description": "Period/header labels, e.g. [\"FY2024A\",\"FY2025A\"]" },
    "rows":    { "type": "array",   "description": "[{label, values:[...]}] — REPORTED figures only, raw numbers" }
  }
}
```

### Behavior (webview executor)

1. Ensure `sheet` exists (free now — auto-create landed 2026-08-18).
2. Find placement: two rows below the sheet's used range (blocks stack).
3. Emit ordinary mutations through the standard batch pipeline
   (reviewable, undoable, saved like everything else):
   title row (bold) → source+units row (small, gray) → header row →
   data rows (labels in col A, values across).
4. **Return an address map** — the whole point:

```jsonc
{
  "ok": true, "sheet": "Data",
  "block": "A12:F19",
  "cols": { "FY2024A": "B", "FY2025A": "C" },
  "rows": { "Revenue — Products": 15, "Revenue — Services": 16 },
  "hint": "Reference these cells by formula (=Data!C15). Derive, don't retype."
}
```

The model never plans staging layout, never tracks addresses mentally, and
gets formula targets keyed by the labels it just wrote.

### Ecosystem touches (small, deliberate)

- `keep_pages.extracted_notes` description: add one line — figures staged
  via `stage_data` need not be duplicated in notes; **staged cells are the
  durable record** (moves the scratchpad from prose to cells, no mandate).
- Color convention alignment (existing rule 15): staged cells are the blue
  hardcoded inputs; presentation cells referencing them are green
  cross-sheet links. The convention already describes this world.
- **No new prompt rules.** The tool description carries the workflow. We
  measured rule-mandates reducing quality; affordances + descriptions only.

### Explicit non-goals for v1

No financial-statement semantics, no auto-derivation, no forced usage, no
fetch→stage fusion (that's v2 if adoption is good). v1 is a well-lit path,
not a wall.

## Implementation sketch

| piece | where | est. |
|---|---|---|
| Tool spec | `src-tauri/.../tools.rs` | ~40 lines |
| Interpret + validate → `{kind:"stage"}` | `agent/toolToMutation.ts` | ~50 lines |
| Pure layout helper: (usedRange, payload) → mutations + address map | new `agent/stageLayout.ts` | ~80 lines, unit-tested |
| Apply branch (placement read → helper → standard apply → report map) | `SpreadsheetScreen.tsx` | ~60 lines |
| Tests | `__tests__/stageLayout.test.ts` | placement, stacking, map correctness, label collisions |

## Measurement (before/after, same judge pipeline)

3× dh + 3× aapl at xhigh on the new build. Success criteria vs parity
baseline (judges 5.3 / 5.5; 55k / 106k tokens; plugs flagged in every run):

- ≥ half of fetched figures land via `stage_data`; presentation rows
  reference staged cells by formula.
- Judge "architecture"/"data_quality" dimensions move up; "back-solved
  plug" gap mentions drop.
- Output tokens and wall-clock down (less mental arithmetic to think
  through).
- No regression on edit tasks (tool is inert there).

## v2 candidates (only if v1 adoption is real)

- `keep_pages` emits an optional structured table → harness stages it
  automatically (fetch→stage fusion; zero extra turns).
- `done`-gate lint: rows labeled *derived/total/margin* that are all
  literals get rejected mechanically.
- Auto-chunk oversized `set_range` payloads harness-side.
