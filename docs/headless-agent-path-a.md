# Headless agent execution — Path A: Univer in a Web Worker

Status: **design / not started** (written 2026-08-11, for evaluation)

## Problem

Every agent run executes against a mounted `UniverGrid` — DOM, canvas, render
units, style caches. `parallel_sessions` (a8e4a9a) froze the render loop of
background tabs, so hidden runs no longer cost CPU/paint, but each concurrent
run still requires a full mounted instance. That caps practical parallelism at
"a few tabs" and ties agent runs to mounted React components.

Univer's data model, command pipeline, and formula engine do not need any of
that. Path A: host a **full headless Univer instance in a dedicated Web
Worker** per background session and drive it with the same commands the grid
handle issues today. One toolchain, no pixels. (Path B — a Rust-native
executor on IronCalc — is a separate, more radical option with a restricted
tool surface and a second toolchain to maintain; not covered here.)

## Non-starter to name up front

The existing `univer.worker.ts` is a **calc replica** slaved to the
main-thread instance via Univer's worker sync protocol. It cannot be promoted
to primary. The headless worker is a new, self-sufficient instance assembled
from base packages:

- `@univerjs/core`, `@univerjs/sheets`, `@univerjs/engine-formula`,
  `@univerjs/sheets-formula`, `@univerjs/sheets-numfmt`
- the **base** (non-UI) halves of filter / conditional-formatting /
  data-validation
- **no** `@univerjs/ui`, no `@univerjs/sheets-ui`, no `@univerjs/engine-render`

This is Univer's supported headless composition (it is how Univer runs on
Node); booting it inside a browser worker instead is the first assumption to
validate.

## Components

### 1. `headless/univerHeadless.worker.ts`

Boots the headless instance, loads a workbook via `xlsxBytesToWorkbook`
(ExcelJS + `exceljsPatches` are pure JS — worker-safe), and exposes an RPC
surface over `postMessage` covering the **agent-facing subset** of
`UniverGridHandle`:

`setCells`, `setRange` (bulk), `readRange`, sheet create/rename/delete,
row/col insert-delete, `defineName`, `copyRange`, `setFormat`,
`describeWorkbook`/`findRows` (via `agent/workbookIndex.ts` — pure snapshot
code), `getSnapshot`, `exportBytes`, `whenCalculated`.

Explicitly out of scope: UX methods (`jumpToCell`, selection, CF/DV panel
opening, filter toggles as UI). Headless runs have no user.

### 2. `agent/executor.ts` — the load-bearing refactor

Extract SpreadsheetScreen's inline tool-apply logic (`runToolTask` and
friends) into a `WorkbookExecutor` interface with two implementations:

- `GridExecutor` — wraps the mounted `UniverGridHandle` (current behavior,
  used whenever the tab is open in the UI);
- `HeadlessExecutor` — wraps the worker RPC session.

The Rust loop changes ~nothing: it already emits `tool_call` events keyed by
`tab_id`; the main thread routes each tab's events to whichever executor owns
the tab, and `spreadsheet_tool_result` flows back unchanged. This refactor is
independently valuable (the apply loop becomes testable), and it is where most
of the engineering lives — today's loop entangles execution with tab state,
status UI, and review-preview building, which must be decomposed into
execution (worker-safe) vs presentation (main-thread only).

### 3. Tauri-invoke proxy

Workers cannot call `invoke` (no `__TAURI_INTERNALS__`). Everything the
headless session needs from Rust relays through a thin main-thread message
proxy:

- IronCalc applier calls (`calc_engine_*`)
- `sessionDb` writes (batches, session log rows)
- save commands (`patch_into` / `remember_base`)

Traffic is low-rate (per batch, not per cell). Boring plumbing, but it touches
everything — spike it early.

### 4. Calc

Keep IronCalc as the engine (the correctness default): the headless worker
mirrors edits through the proxy exactly like `ironcalcShadow.ts` does from the
main thread, and `whenCalculated` becomes an IronCalc-settled check. Side
effect: with calc in Rust, the worker holds only the data model + styles — the
lightest footprint, plausibly 30–60% of a mounted grid (style-render caches
and canvas bitmaps are the mounted instance's big allocations).

### 5. Save

Model A helps here: save = replay the whole session against the load-time
base. The machinery is TS (`surgicalPatch.ts` — pure, NUL-byte map keys and
all) + Rust (`patch_into`, via proxy), and the ExcelJS full-export fallback
(`workbookToXlsxBytes`) also runs in a worker. So headless saves can be
**feature-parity**, not restricted. One UX change: "reduced fidelity"
outcomes are recorded in the session log instead of raising
`FidelitySaveModal` mid-run (there is no user mid-run).

### 6. Review, run_script, and the two-writers rule

- Batches append to `sessionDb` identically (through the proxy); the existing
  restore-and-review flow works untouched when the user later opens the
  session in a grid.
- `run_script`'s sandbox worker talks to the headless worker over a
  `MessageChannel` instead of the grid handle.
- A session with an active headless run must not be simultaneously opened in
  a grid (two writers, one workbook). v1: block the open until the run
  finishes ("agent working — try again shortly"). Live handoff
  (worker ships snapshot → grid adopts mid-run) is a v2 problem; do not
  solve it up front.

## Evaluation sequence (each spike kills the biggest remaining risk)

1. **Boot spike (~a day).** Headless Univer in a worker from base packages;
   load a real client xlsx; `setValues` + formula readback. Measure worker
   memory vs the same file mounted. If the plugin graph won't assemble
   without the UI packages, Path A dies cheaply right here.
2. **Proxy spike.** Worker→main→Tauri relay; IronCalc applier + a `sessionDb`
   write from the worker.
3. **End-to-end A/B via the eval harness** (`evalDriver.ts` / `eval_mode.rs`
   — gridpath-only, which is why this evaluation belongs here). Same prompt,
   same workbook, run through `GridExecutor` and `HeadlessExecutor`; diff the
   saved xlsx and the session logs. Before building the full tool surface,
   mine existing session logs to rank which tools real runs actually use —
   that decides whether formats/merges/scripts must be in v1 or can trail.

## Sizing and decision points

- Spikes: ~1 week. Production quality (executor extraction, save parity,
  lifecycle, block-open rule): 3–6 weeks.
- Highest-risk items, in order: preset-free boot in a worker (unknown until
  spike 1), invoke-from-worker plumbing everywhere, save-fidelity parity on
  style-heavy models (spike 3 is designed to expose it).
- Strategic check before starting: post-render-freeze, mounted-but-frozen
  instances already cost only parked memory. Path A pays off when you want
  N well past 3 concurrent runs, or runs that survive with no tab open at
  all. If neither is pulling, it can wait.

## Porting note (heelix)

Heelix (HeelixNotes `ProjectModel`) is a port of this screen with a diverged
save model (Model B — incremental baking, see the port-map notes there). If
Path A lands here, the `WorkbookExecutor` abstraction ports cleanly; the save
section does NOT port as-is (heelix's mirror is incremental and bakes batches
per save — its headless save path would build the unbaked mirror instead of
replaying the session). Evaluate here first; port after the A/B holds.
