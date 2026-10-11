# Engine-served index and `query_rows`

Two agent read paths moved from the webview into the calculation engine
(2026-10-07). Both are in `crates/engine-core` and exposed by the desktop
(`calc_engine_*` Tauri commands) and the headless package (wasm exports).

## Why

On the 2.3M-cell datacenter model the agent's reads were the slow part of
every transcript:

| Before | Measured |
|---|---|
| `describe_workbook` / `find_rows` built a structural index by walking the Univer snapshot on the main thread | 2.2 s per build, rebuilt after every agent write (the content-version memo only helps when nothing changed) |
| "which rows have AWS in column DB" | no tool; 15 chunked `read_range` calls, or a script that first ships the whole sheet to a sandbox |

| After | Measured (native, same file) |
|---|---|
| Rust index, all 15 sheets | 969 ms once, then memoized per sheet; a write dirties only the sheets it touched |
| `query_rows` filter on the 7,389-row supply sheet (79 matches) | 66 ms |
| `query_rows` group-by over the same sheet | 53 ms |

## The index (`engine-core/src/index.rs`)

A port of `packages/gridpath/src/core/workbookIndex.ts#buildWorkbookIndex`
with the same heuristics and the same JSON shape, so `describeWorkbookPayload`
and `findRowsInIndex` consume either source unchanged. The parity test
(`packages/gridpath/src/core/workbookIndex.parity.test.ts`) asserts exact
agreement on two real workbooks: labels, counts, bold flags, header row and
headers, sections, formula-run spans and formula text.

One deliberate difference in method, not in result: formula runs. The
TypeScript builder shifts each A1 formula one column and compares text; the
Rust builder compares IronCalc's shared R1C1 formula strings, which is the
same test without string manipulation.

Memo: `IndexCache` keeps one serialised entry per sheet, keyed by position
and name. The Tauri commands mark sheets dirty: edits dirty the written
sheets plus every sheet whose values changed in the incremental recalc;
structural ops, sheet ops and resyncs dirty everything.

Consumers:

- Desktop `describe_workbook` / `find_rows` ask the engine first
  (`ironcalcEngineWorkbookIndex`) and fall back to the snapshot walk when the
  engine is off or degraded. Samples for `find_rows` are still read from the
  live grid.
- Headless `WorkbookSession.workbookIndex()` uses the wasm export and falls
  back to the TypeScript builder when the engine binary predates it.
- The review modal's column headers still use the TypeScript builder over
  the grid snapshot (memoized per modal open).

## `query_rows` (`engine-core/src/query.rs`)

Filter, project, order and aggregate the rows of one sheet.

```json
{"sheet": "NA Data Center Supply",
 "where": [{"column": "DB", "op": "eq", "value": "AWS"},
           {"column": "Capacity (MW)", "op": "gte", "value": 100}],
 "columns": ["A", "DB", "Capacity (MW)", "Online"],
 "order_by": {"column": "Capacity (MW)", "desc": true},
 "limit": 50}
```

- Columns: header text (exact, case-insensitive; unique substring accepted)
  first, then column letters. The header row is detected with the index's
  rule, or given as `header_row`.
- Ops: `eq ne contains not_contains starts_with ends_with gt gte lt lte
  between in not_in empty not_empty`. Text compares trimmed and
  case-insensitively, numbers numerically, ISO date strings against
  date-formatted cells by serial.
- Reply: `{sheet, header_row, columns:[{column, header}], rows:[{row,
  cells:{A: …}}], matched, truncated, rows_scanned, row_range}`. Dates come
  back as ISO strings; `include_formulas` returns `{v, f}` for formula cells.
- `group_by` + `aggregate` (`sum count min max avg`) returns `groups`
  instead of rows, ordered by the first aggregate.
- Caps: 500 rows, 60 columns, 10,000 cells per reply, 12 predicates.

Errors are strings meant for the agent: an unknown column lists the header
texts with their letters; an unknown op lists the ops.

## Surfaces touched

- Tool schema: `src-tauri/.../tools.rs` (`query_rows`, rule 17a), regenerated
  into `packages/gridpath/src/tools/tools.json`.
- Parser: `toolToMutation.ts` (`kind: "query_rows"`).
- Desktop: `ironcalcEngine.ts` (`ironcalcEngineWorkbookIndex`,
  `ironcalcEngineQueryRows`), `SpreadsheetScreen.tsx` handlers,
  `calc_engine.rs` commands `calc_engine_workbook_index`,
  `calc_engine_query_rows`.
- Headless: `session.ts` (`workbookIndex`, `queryRows`), `handlers.ts`,
  MCP read-only set; wasm exports `workbookIndex`, `queryRows`.
- Timing harness: `cargo run --release --example index_timing -- file.xlsx
  [sheet] [column] [value]` in `crates/engine-wasm`.

Note for engine-core code: `std::time::Instant` traps on
wasm32-unknown-unknown. Time things in the callers.

## Follow-up found on the first live run (2026-10-07): sheet ops dumped every value

Diagnosed from the app log plus a native replay of the saved workbook
(`crates/engine-wasm/examples/{replay_timing,chunk_replay,probe_cells}.rs`).
The agent built a 9k-cell "Google Capacity" sheet; the grid then showed no
numbers for minutes, copy_range calls timed out and reads returned nulls.

The engine was not the slow part: create_sheet's full evaluate was 14 s and
the nine 1000-cell write chunks took 2.3 s in total, incremental. The stall
was `calc_engine_sheet_op` (and `calc_engine_structural`) returning
`formula_cell_deltas` — the value of EVERY formula cell in the workbook,
~880k on this model — which the webview then applied through one giant
`set-range-values` per sheet on the main thread. Minutes per create_sheet,
during which every later tool call queued behind it.

Fix: both commands now diff formula values before and after the operation
(`formula_value_map` / `changed_deltas`, remapping pre-op coordinates by the
same shift the grid applied) and return only cells whose value actually
changed — typically none for a sheet add, a handful for a row insert. The
full dump remains only for resync, which is the broken-state recovery path.
Both commands also log their evaluate time and delta count, as does
`calc_engine_edit_batch`, so the next stall is readable from GridPath.log.
