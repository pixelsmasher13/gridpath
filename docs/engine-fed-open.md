# Engine-fed open (desktop)

**Status:** shipped 2026-10-06, default on. Kill switch:
`localStorage.setItem("gridpath.engineOpen", "0")` + reload.

## Why

Opening a workbook in the desktop app used to parse the file twice: the
IronCalc engine natively (for calculation) and ExcelJS in a web worker (to
build Univer's grid). On a 2.3M-cell model the ExcelJS side produced 285 MB of
cell objects and 2.27M style ops, needed 4.3 GB of heap in Node, and the
structured clone of that result across the worker boundary alone took 13 s —
inside a webview, with a worker heap cap, it thrashed for minutes under the
"Parsing workbook…" label.

## What happens now

`UniverGrid.loadBytes(bytes, onProgress, { tabId, path })`:

1. `ironcalcEngineOpenSnapshot` (ironcalcEngine.ts) runs `calc_engine_load`
   and then `calc_engine_univer_snapshot`, a Tauri command returning the
   engine's whole-workbook snapshot as raw JSON bytes (`tauri::ipc::Response`,
   so the 100+ MB payload is not JSON-encoded twice). The serialiser is
   `crates/engine-core` (`snapshot_json`), shared with the wasm engine; it
   resolves theme colours to `#RRGGBB` and streams into one string.
2. In parallel, the ExcelJS worker runs in `metaOnly` mode
   (`xlsxBytesToWorkbook(bytes, { metaOnly: true })`): no per-cell styles, no
   cellData, no styleOps — only what the engine does not carry: conditional
   formatting, data validation, notes, outlines, images, external-reference
   pins, gridline/background-picture flags.
3. `mapSnapshot(engineSnapshot, id, { trim: true })` (packages/gridpath
   workbook/snapshot.ts — the same mapper the MCP session uses) builds the
   Univer workbook data: cells with inline styles shared by reference per
   style index, borders, merges, widths/heights, freeze, hidden, tab colours.
   Styled-blank tails are trimmed like the ExcelJS import did.
4. The ExcelJS resources and flags are merged in, `createWorkbook` runs, and
   the rest of the load (feature baseline, images, outlines) is unchanged.
   The engine is already loaded, so the caller must not start it again:
   `loadBytes` returns `{ engineLoaded: true }` and SpreadsheetScreen skips
   `ironcalcEngineOnOpen` (which would otherwise push 878k formula values into
   the grid as mutations — they are already in the snapshot).

If the engine is off, the path is not a file, or the engine load fails, the
grid is built from the full ExcelJS import exactly as before.

## Fidelity deltas vs the ExcelJS import (measured on CRM.xlsx, 192k cells)

- Values: 0 mismatches. Formulas: 0 lost. Engine values are freshly computed,
  the ExcelJS path showed the file's cached values until the engine caught up.
- Implicit intersection: the engine prints `=@IF(...)` where the file holds
  `IF(...)`; `stripImplicitIntersection` restores the file text (`@` is not
  valid inside `<f>`). Redundant parentheses are also normalised by the
  engine (`MATCH(("FY"&J16),…)` → `MATCH("FY"&J16,…)`); semantically equal.
- Number formats are the file's raw strings (`#,##0_);(#,##0)`,
  `yyyy\-mm\-dd`); ExcelJS had unescaped them.
- Pattern fills (`patternType="lightGray"` with a foreground colour) render
  as a solid fill of that colour; ExcelJS dropped them.
- Merged ranges carry only the anchor cell's value; ExcelJS replicated it
  into every covered cell.
- Cells with an explicit default font (black Calibri 11) carry no style.

## Lazy sheets (2026-10-06; OPT-IN since 2026-10-07)

**Default off.** Enable with `localStorage.setItem("gridpath.lazySheets", "1")`
and reload; the first click on a big lazy tab costs seconds and the eager open
is being compared against it. When enabled: above 400,000 cells
(`LAZY_SHEETS_CELL_THRESHOLD`), the engine snapshot
carries full cells for the first visible sheet and for every sheet of
100,000 cells or fewer (`LAZY_SHEET_MIN_CELLS`: small tabs must switch
instantly, and they are not where memory goes). Each remaining sheet arrives
with its header, extent and a **skeleton** — the first 60 rows and the first 12
columns (`SnapshotOptions.skeleton_rows/cols`) — flagged `lazy`. The skeleton is
enough for the tab strip, the sheet's scroll extent, the agent's structural
index (`describe_workbook` / `find_rows` read labels from the leading columns
and header rows) and a visible preview; the grid never holds the body of a
sheet nobody has looked at.

`UniverGrid.ensureSheetLoaded(name)` fetches a lazy sheet's cells with
`calc_engine_sheet_cells`, lands them as one raw `set-range-values` mutation
(no undo entry, invisible to the engine mirror), records the sheet's save
baseline from that content, drops the lazy flag and flushes the engine values
that queued for the sheet while it was lazy (the engine applier refuses writes
to lazy sheets so they wait in the bridge's pending queue). It runs:

- when the user activates the sheet (`sheet.operation.set-worksheet-active`);
- before any agent write that names the sheet (one ensure step at the tool
  funnel in SpreadsheetScreen, via `sheetsReferencedByMutations`);
- before `read_range` on the sheet, before a redo of a rejected batch, and
  before an unsaved-work journal replay (`journalSheets`).

Known gaps while a sheet is lazy: `find_rows` samples and the agent's
per-prompt previews see only the skeleton; a user paste can only target the
active (hence loaded) sheet. Nothing can write to a lazy sheet, so its baseline
captured at load time is the file's content.

### Resync and lazy sheets

The engine mirror's resync rebuilds the engine from the grid ("grid as truth")
after a command it cannot mirror (sort, range move). With lazy sheets the grid
is NOT the truth for sheets it holds only a skeleton of, so the grid's resync
snapshot marks those `keep: true` and the engine (`resync_clear_contents_except`)
leaves their cells untouched while re-seeding the others.

### Why new-sheet formulas used to read as null on huge workbooks

A sheet the agent creates reaches the engine through a queued sheet-op; the
write that follows (`set_range` on the new sheet) used to check the engine's
sheet list synchronously, before that op ran, call the sheet unknown and fire a
resync. On a grid over 2M cells a resync is refused and the tab is marked
*degraded*: every later edit is dropped and new formulas never get values. The
sheet index is now resolved when the queued edit runs (`enqueueEdits`), and a
degraded tab raises an error toast instead of failing silently.

### Landing a lazy sheet

`ensureSheetLoaded` lands the fetched cells in chunks of 40,000
(`LAZY_INGEST_CHUNK`), top rows first, yielding to the event loop between
chunks so the tab paints immediately and fills in downward instead of freezing.
A second ingest path writes straight into the worksheet's cell matrix and
nudges Univer with one tiny mutation (`localStorage.setItem("gridpath.lazyIngest",
"matrix")`); it bypasses the command pipeline and its interceptors. Both log
`[univer] lazy sheet loaded: … via <mode> in Nms (fetch+parse, map, ingest in
k chunks, baseline)`, so one tab click on a big sheet shows where the time goes
and which path to keep.

The large benchmark model (v2) with the floor: 13 of 17 sheets eager, 4 lazy (its four
largest sheets); the grid holds
480k of 2.3M cells; 27 MB snapshot; the 828k-cell sheet fetches in
2.2 s and maps in 0.3 s before ingest.

### Engine-mirror fallbacks that were poisoning big workbooks (2026-10-07)

Two agent tool paths made the mirror fall back to a grid→engine resync, which
on a grid over 2M cells is refused and degrades the tab (every later value
frozen, new formulas null):

- a write right after `create_sheet` (sheet index checked before the queued
  sheet-op ran) — fixed by resolving the index when the queued edit runs;
- `clear_range` — Univer's `range.clearContent()` dispatches the clear
  command with `ranges: [...]`; the mirror read a singular `range`, found
  nothing and resynced. Fixed to decode both shapes.

A degraded tab now raises an error toast with the reason. `describe_workbook`
/ `find_rows` and the per-prompt context capture no longer hash every cell of
the snapshot per call to detect change (seconds on 2M cells, which showed up
as tool timeouts); the grid exposes a content version bumped on every sheet
mutation and they key their memos on it.

### run_script on huge workbooks (2026-10-07)

The sandbox used to clone the whole grid snapshot into its worker and refuse
above 1M cells, which turned every scan on the large benchmark model into dozens of
500-cell reads. In engine mode the read model now comes from the ENGINE, per
sheet: `executeSheetScript(..., { source })` fetches `calc_engine_sheet_cells`
bytes for each sheet the script names as a string literal (or all sheets if it
calls `sheets()` / uses a dynamic name), transfers the buffers into the worker
and parses them there (`buildScriptModelFromEngineSheets`). Nothing is copied
out of the grid, lazy sheets are irrelevant, and the only cap is 400 MB of
sheet JSON per run. A sheet the script reads without naming throws an
agent-readable error instead of returning blanks. The grid-snapshot path stays
for the Univer-engine mode. `read_range` results also carry
`calc_engine: "degraded"` plus the reason when the mirror has given up.

## Deferred first evaluate (2026-10-07)

Native stage timings on the 2.3M-cell / 922k-formula benchmark model at
the app's optimization level: xlsx parse 6.8 s, formula parse 0.1 s, first
full evaluate 6.4 s, snapshot to 141 MB JSON 4.0 s, JSON.parse ~1–2 s, then
the grid build. The evaluate no longer sits on that path: `calc_engine_load`
takes `defer_eval`, the engine-fed open passes it, the snapshot carries the
file's cached values (what Excel shows on open too), and the mirror queues
`calc_engine_evaluate_pending` right after the open. It runs the evaluate
once and returns only formula cells whose computed value differs from the
cache — normally a handful (constructs we evaluate differently, stale caches
from manual-calc files). Any edit, sheet op, structural op or resync that
arrives first settles the evaluate itself, so computed values are never
mixed with cached ones in one reply.

The grid also logs its own build stages (baseline map, fidelity scan,
createWorkbook, total) to GridPath.log as `[univer] open stages`, next to
the engine's `load … parse N ms` line, so an open can be read stage by stage
after the fact. Remaining cuts, in order: parallel sheet parsing in the
importer (the two 800k-cell sheets bound it, ~6.8 → ~3 s), a columnar
snapshot instead of 141 MB of JSON, and progressive per-sheet ingest so the
first sheet paints before the rest are built.
