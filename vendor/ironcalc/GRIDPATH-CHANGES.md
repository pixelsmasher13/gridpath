# Vendored IronCalc fork (gridpath calc-engine spike)

Vendored from https://github.com/ironcalc/IronCalc at upstream commit
`38cfe07` ("update: missing args for Z.TEST", 2026-07-24, v0.7.1 line).
Only the `base` and `xlsx` crates are vendored (engine + xlsx I/O); the
webapp, docs, bindings and codegen directories are not needed.

License: MIT OR Apache-2.0 (files included).

## Gridpath changes on top of upstream

All changes are candidates for upstreaming (see upstream issue #849 for the
incremental-recalc ask).

### 1. Incremental recalculation — `Model::evaluate_edited()` (base/src/model.rs)

Upstream `evaluate()` clears every cached result and re-evaluates the whole
workbook on each call. `evaluate_edited(&[(sheet, row, col)])` instead:

- inverts the runtime-recorded dependency map (`support`, a forward map
  reader → cells/ranges read, despite its doc comment),
- walks the dirty closure from the edited cells (spill anchors expand to
  their written areas; spill positions hold no formula, so dirty anchors are
  re-evaluated first, in the order fixed by the last full evaluation),
- invalidates and lazily re-evaluates only the dirty cells,
- returns the changed cells so a UI can refresh exactly those.

Requires dependency recording in `fn_indirect`, `fn_offset` and defined-name
resolution (they build references programmatically, invisible to the
parser-level ReferenceKind/RangeKind recording). Runtime-recorded deps make
INDIRECT-heavy models safe: the cells that *select* an INDIRECT target are
themselves recorded dependencies.

Not handled (callers must fall back to full `evaluate()`):
- structural changes (insert/delete rows/columns, sheet renames,
  defined-name redefinition),
- volatile refresh (RAND/NOW/TODAY are not re-evaluated),
- spill areas that grow or shrink.

### 2. Range clamp in `prepare_array` (base/src/functions/binary_search.rs)

Whole-column references (`G:G`, the Canalyst `MATCH(x, INDIRECT(...))`
pattern) materialized 1,048,576 cells per lookup (~11 ms each). The span is
clamped to the sheet's used area; cells beyond it are empty and can never
match a non-empty target.

### 3. xlsx import speedups (xlsx/src/import/worksheets.rs)

- One formula `Parser` per sheet instead of one per formula cell (each
  construction cloned the sheet list, tables and all defined names — ~7 s on
  a 55k-formula file with 3.4k defined names).
- `HashMap` formula-string index instead of a linear scan per cell
  (quadratic).
- Shared-formula anchor met after a daughter now *replaces* the placeholder
  (`shared_formulas[i] = f`) instead of `Vec::insert(i, f)`, which shifted
  every later index and corrupted already-assigned cells.

### 4. Instrumentation

`IC_DEBUG=1` prints phase timings for import and incremental evaluation to
stderr.

### 5. Bench binary (xlsx/src/bin/bench.rs, gridpath-only)

`cargo run --release --bin bench -- file.xlsx [sheet row col value]`
- load / full evaluate ×2 / value diff against the file's cached values
- incremental edits ×3 + full-recalc comparison
- correctness check: fresh model + same edit + full evaluate, compared
  cell-by-cell against the incrementally updated model (expect `0 diffs`).

### 6. Resync-from-grid support (base/src/model.rs)

Engine-mode recovery path for operations the app's command bridge cannot
mirror (sort, range moves, unrecognized commands): the model is reseeded IN
PLACE from a grid content dump instead of freezing until reopen.

- `Model::resync_clear_contents()` — clears all cell content, formula
  storage and evaluation caches while keeping sheets, defined names, styles
  and shared strings (the grid does not know the workbook's defined names,
  so the model must never be rebuilt from scratch).
- `Model::seed_formula_with_cached_value()` — the resync equivalent of an
  imported formula cell: parses and stores the formula, sets the grid's
  current value as the cell's cached value, and registers wrong-reference
  formulas in `import_wrongref_frozen` exactly like `from_workbook` — so
  frozen add-in / external-ref data survives a resync.
- `Model::seed_cell_literal()` — typed literal setter (no input-string
  parsing: a grid string "123" or "TRUE" stays text).

### 7. INDIRECT resolves defined names (base/src/functions/lookup_and_reference/mod.rs)

Excel's `INDIRECT("MyName")` resolves defined names; upstream only parses
A1-style cell/range strings and returned #REF! otherwise. Analyst models
lean on this (`ROW(INDIRECT("SomeNamedRow"))` to locate rows, then
`INDIRECT("Model!"&ADDRESS(row,col,4))` to pull values) — on ASML.xlsx the
gap silently zeroed IFERROR-wrapped lookups and cascaded into ~7.6k wrong
cells (#DIV/0! where Excel shows numbers).

On reference-parse failure `fn_indirect` now falls back to the
`parsed_defined_names` table: bare names try calling-sheet scope then
workbook scope; sheet-qualified names (`"Sheet2!Local"`, quoted sheets
handled) try that sheet's scope then workbook scope. Cell/range names
return references (support-recorded for incremental recalc); lambdas and
malformed names still return #REF!. Tests in
base/src/test/lookup_and_reference/test_fn_indirect.rs. Also added
xlsx/examples/debug_diverr.rs: loads a workbook, recalcs, and prints every
formula cell whose computed value is an error where the file's cached value
was not, grouped by error origin — the tool that isolated this bug.

## Measured on project-159.xlsx (55k formulas, 4.4k INDIRECT, M-series Mac)

| metric | upstream | this fork |
|---|---|---|
| load | ~11 s | ~3.3 s |
| full evaluate | ~13.3 s | ~3.1 s |
| single-edit recalc | ~13 s (full) | **~70 ms** (292-cell closure) |
| incremental vs full recalc diff | — | 0 cells |
| upstream test suites | — | base 2213 pass (1 pre-existing failure: bitcode message drift), xlsx 41 pass |

Univer OSS on the same file/machine: initial calc 3.7 s, per-edit ~1.3 s.

## wasm32 build support (2026-09-29)

Both crates now compile for `wasm32-unknown-unknown` (headless Node / browser
use of the engine via wasm-bindgen):

- `xlsx/Cargo.toml`: `zip` 0.6 with `default-features = false, features = ["deflate"]`.
  The default bzip2/zstd features pull in C libraries that don't build for wasm;
  xlsx packages are deflate-only anyway.
- The `IC_DEBUG=1` phase timers in `xlsx/src/import/mod.rs` and `base/src/model.rs`
  use `debug_timer::DebugTimer` instead of `std::time::Instant` (which aborts on
  wasm32). The timer is inert (reports 0 ms) on wasm32, unchanged natively.

Known pre-existing test failure, unrelated: `base` `test_to_from_bytes::errors`
asserts bitcode's old "invalid packing" message; the crate now says "bitcode error".

## Excel 1900 date-system fidelity: serial 0 and the phantom 29 Feb 1900 (2026-10-06)

Upstream rejects date serial 0 (`MINIMUM_DATE_SERIAL_NUMBER = 1`) and follows
Google Sheets for January–February 1900 (`DATE(1900,1,1)` = 2). Excel treats
serial 0 as "0 January 1900" (what a blank or zero date cell yields) and counts a
phantom 29 February 1900 as serial 60. On a real analyst model (datacenter
supply/demand, 878k formulas) ~880 cells did `YEAR`/`MONTH`/`EOMONTH` on blank
date cells; the resulting `#NUM!` cascaded through SUMIFS into 38k cells.

- `base/src/constants.rs`: `MINIMUM_DATE_SERIAL_NUMBER = 0`.
- `base/src/formatter/dates.rs`: `convert_to_serial_number` matches Excel before
  1900-03-01 (`DATE(1900,1,1)` = 1, `DATE(1900,2,28)` = 59, `DATE(1900,2,29)` = 60,
  `DATE(1900,1,0)` = 0). `from_excel_date` accepts 0 and keeps its uniform,
  weekday-preserving mapping so WEEKDAY/WORKDAY/NETWORKDAYS/EDATE still match
  Excel. New `excel_display_date` gives the formatter Excel's year/month/day for
  serials 1..=59.
- `base/src/functions/date_and_time.rs`: DAY/MONTH/YEAR and EOMONTH go through
  `excel_serial_to_ymd` (already used by DAYS360), so `YEAR(0)` = 1900,
  `MONTH(0)` = 1, `DAY(0)` = 0, `DAY(60)` = 29, `EOMONTH(0,0)` = 31,
  `EOMONTH(59,0)` = 60; months before January 1900 are `#NUM!`.
- Tests updated to Excel's values. base: 2226 pass.

Known remaining deviation: serial 0 and 60 have no calendar date, so they
format as 12/31/1899 and 3/1/1900 where Excel shows 1/0/1900 and 2/29/1900.

## `UserModel::get_model_mut` (2026-10-06)

`base/src/user_model/common.rs`: mutable access to the inner `Model`, so the
wasm engine (`crates/engine-wasm`) can pause the UserModel's automatic
evaluation and drive `Model::evaluate_edited` itself. Before this every
`set_user_input` through the UserModel ran a full recalculation (5+ minutes on
an 878k-formula workbook) and the headless session ran another one after the
batch; now a batch of cell edits recalculates its dependency closure once.

## SUMIFS criteria index — `Model::ifs_cache` (2026-10-06)

`base/src/functions/statistical/if_ifs.rs` `run_ifs` (SUMIFS, AVERAGEIFS and
the other `apply_ifs` consumers; COUNTIFS keeps its own walk for now): the
per-pass cache `ifs_cache` maps (sheet, top row, left column, height, width,
criterion) → the offsets of the criteria range that satisfy the criterion
(`IfsMatch`: ascending row-major `matched` list plus a bitset). The first
formula testing a (range, criterion) pair walks the range once; every later
one reuses it and visits only matched sum cells, intersecting multiple
criteria through the bitsets. Same lifecycle as `range_cache` (cleared at pass
start, spill boundary and in `evaluate_edited`). Shapes over 4M cells and
range/array criteria take the plain walk. Summation order and the
first-matched-error semantics are unchanged.

Why: a real supply-demand model has 204k SUMIFS over 7,167-row type / company /
region columns — 4.3 billion cell visits per recalc. Measured natively on it:
full evaluate 300 s → 51 s; parse stays 6 s. base: 2227 tests pass
(`test_sumifs_shared_criteria_index_matches_plain_walk` added).

## Per-pass `Worksheet::dimension()` memo — `Model::cached_dimension` (2026-10-06)

`Worksheet::dimension()` scans every cell of the sheet (the upstream FIXME says
as much). Whole-column references — `SUMIFS('Supply'!T:T, …)`, `SUM(A:A)`,
`XLOOKUP(.., A:A, ..)` — resolve their extent through it once per formula; on an
800k-cell sheet a `sample` profile put 98% of a 51 s recalculation inside it.
`Model::dimension_cache` memoises it per recalculation pass (same lifecycle as
`range_cache`; `WorksheetDimension` is now `Copy`), and the 39 formula-side call
sites of the `Ok(s) => s.dimension().max_row` pattern go through
`Model::cached_dimension(sheet)`. Non-evaluation callers (actions, clipboard)
still call the worksheet directly.

Measured natively on the supply-demand model: full evaluate 51 s → 10.3 s
(300 s before the SUMIFS index). base: 2227 tests pass.

## `Model::get_formatted_cell_value_using` (2026-10-06)

`base/src/model.rs`: `get_formatted_cell_value` with the number format passed
in. `get_formatted_cell_value` builds a full `Style` (font, fill, border
clones) per call just to read `num_fmt`; the wasm engine's snapshot serialiser
memoises `num_fmt` per style index and calls this instead. Part of the
streaming-snapshot rewrite in `crates/engine-wasm` (2.3M-cell snapshot: 147 s →
5.7 s in wasm, 10.6 s → 3.3 s natively).

## Array idioms: ROW/COLUMN arrays, `--(cond)`, `-(range)`, TEXTJOIN over arrays (2026-10-06)

Found by reading an agent's reasoning on a datacenter supply model: it had
concluded "dynamic arrays aren't supported here" after `SMALL(IF(rng<>"",
ROW(rng)), 2)` returned the first row and `TEXTJOIN(",", TRUE, IF(...))` gave
`#N/IMPL!`. Dynamic arrays (UNIQUE/FILTER/SORT/SEQUENCE/LET/LAMBDA) were fine;
four narrower things were not:

- `ROW(range)` / `COLUMN(range)` returned only the top-left number. Excel
  returns an array, which every ranking idiom relies on
  (`LARGE(IF(rng<>"", key*1e5+ROW(rng)), k)` then `MOD(…,1e5)` to decode the
  row — with a scalar ROW every key decoded to the same row; the agent's
  "Largest Data Centers" sheet was built on exactly that). Now an array for
  multi-cell references, a scalar when every element would be equal (one-row
  span for ROW, one-column span for COLUMN), whole-column/row references
  clamped to the sheet's used extent. `lookup_and_reference/mod.rs`.
- The parser folded consecutive signs, so `--(A1:A6<>"")` parsed as
  `A1:A6<>""`: the boolean→number coercion vanished and every
  `SUMPRODUCT(--(cond), vals)` was 0. Every unary minus is now kept.
  `parser/mod.rs parse_power`.
- Unary minus over an array or multi-cell range now broadcasts (`-(B1:B6)`
  in SUM/SUMPRODUCT was `#N/IMPL!`; `--(bool array)` gave 0). `model.rs`.
- Stringify printed `-(a<>b)` as `-a<>b` (which re-parses as `(-a)<>b`):
  comparisons and concatenations under a unary minus are parenthesised;
  `--x` round-trips as written. `parser/stringify.rs`.
- TEXTJOIN accepts array arguments (`IF(rng="x", vals, "")`), honouring
  ignore_empty. `text/common.rs`.

Not implemented (still `#NAME?`): AGGREGATE. base: 2228 tests pass
(`test_fn_row_column_arrays` added).

## Cached inverted dependency index for `evaluate_edited` (2026-10-07)

Every incremental pass rebuilt the dependency → readers map from the whole
`support` table before walking the dirty closure: 415–458 ms on the
922k-formula datacenter model, paid for a ONE-cell edit whose closure and
re-evaluation took 0–30 ms. An agent run that wrote 2,000 cells one command
at a time (copy_range "all") spent ten minutes in that rebuild.

`model.rs`: `ReaderIndex` (point readers, narrow column ranges, row bands,
wide rects — the same bucketing as before) lives on the model as
`reader_index: Option<ReaderIndex>`. It is `None` after a full `evaluate()`
(both clear sites), built on the first `evaluate_edited`, and patched on
every pass: each dirty cell's old edges leave with its `support` entry and
the edges its re-evaluation records are added back. A dirty cell's
evaluation can only recurse into other DIRTY cells (everything else is
cached in `cells`), so re-indexing `dirty` covers every cell that recorded
support. `new_empty.rs`: field init. IC_DEBUG prints `reader index: … ms
(rebuilt|cached)`.

## `Styles::get_style` made public (2026-10-08)

`base/src/styles.rs`: the viewport grid's layout serialiser enumerates the
workbook style table by index; the accessor was `pub(crate)`.

## Full evaluate: spill-ordering pre-pass skips scalar dynamic formulas (2026-10-08)

`Model::evaluate` runs a pre-pass over dynamic-array anchors that restarts the
whole evaluation (clearing every evaluated cell) whenever an anchor's spill
area turns out to be a dependency of an anchor evaluated earlier — up to n²
restarts. GridPath saves every agent-written formula as a dynamic-array
formula, so a model with 25k such anchors (datacenter v4) never finished its
first evaluate (7+ hours of CPU, `computed` cells climbing to millions).
A dynamic formula whose result is a single cell writes only its anchor, which
any earlier reader already evaluated on demand, so the ordering check (and
the O(i) support scan per anchor) now applies only to multi-cell spills.
v4: never → 7.2 s.

## Import: worksheet parts parsed in parallel (2026-10-09)

`load_sheets` reads every worksheet part sequentially (one zip reader),
then parses them on one thread per core (`std::thread::scope`; sequential
on wasm32). `load_sheet` became `parse_sheet_xml`, pure over its inputs:
the shared-string table is read-only during the parse (`StringTable`
interns new strings per sheet from a base id), conditional-format dxfs are
appended per sheet from `dxf_base`, and `load_sheets` merges the additions
in sheet order, shifting each sheet's new string/dxf ids by the sizes of
the sheets merged before it. The sequential and parallel parses produce a
byte-identical snapshot (datacenter v4, 142 MB JSON, same SHA-1).
`IC_SEQUENTIAL_PARSE=1` forces the single-threaded path. v4 parse on a
loaded 10-core machine: 2.8–4.1 s → 1.1–1.5 s. The linear shared-string
scan for `str`/inline cells is also gone (hash index); it was not the
hotspot here but was O(cells × strings).

## `reset_parsed_structures` is public (2026-10-09)

A defined name added after load must re-parse formulas that already
referenced it (they parsed as #NAME?). The desktop `calc_engine_define_name`
and the WASM `newDefinedName` call it after `new_defined_name`.

## *IFS criteria: lowercase once per call (2026-10-10)

`build_criteria` lowercased the criterion string inside the comparison
closure — once per compared cell, i.e. 7,000 times per SUMIFS over a
7,000-row range. Hoisted to one lowercase per criterion. Profile of the
datacenter v4 first evaluate: `to_lowercase` 3.9% + the closure 4.8% of
samples; evaluate 8.3–8.6 s → 7.6–7.9 s.

## Number formatting rounds half away from zero on the decimal value (2026-10-10)

`to_precision_str` rounded with `{:.*e}`, which resolves an exact binary tie
to even: 940.5 under `#,##0` displayed as 940, and a sum stored as
940.49999999999994 displayed as 940 where Excel shows 941. It now rounds
the 17-digit decimal expansion by hand: to Excel's 15 significant digits
first (so 2.675, stored as 2.67499999999999982, is 2.675), then to the
requested precision, half away from zero. Found by the byte-for-byte diff
of the agent context between the Univer grid (which used a JS formatter
with Excel's rounding) and the viewport grid (engine display strings).


## Outline groups imported (2026-10-10)

Excel's row/column groups (`<row outlineLevel>`, `<col outlineLevel>`,
`<sheetPr><outlinePr summaryBelow/summaryRight>`) were on the importer's
ignore list. `Worksheet` now carries `row_outline` / `col_outline`
(1-based index → level) and the two summary flags; `insert_rows`,
`delete_rows`, `insert_columns` and `delete_columns` shift them with the
rows/columns. Nothing in the evaluator reads them; the viewport grid draws
the group gutters and +/- buttons from the layout. Export does not write
them (GridPath saves through the patcher, which leaves the attributes in
place).

## Saved view: frozenSplit panes, topLeftCell, selected tab (2026-10-10)

`get_sheet_view` only treated `<pane state="frozen">` as frozen; Excel
also shows `frozenSplit` (a split that was then frozen) as frozen panes,
and research models commonly save that way (TOST: xSplit=2 ySplit=5
frozenSplit). Both states now freeze. The scrolling pane's first cell
(`pane topLeftCell`, else `sheetView topLeftCell`) now lands in
`WorksheetView.top_row/left_column` instead of the constant 1, so a
viewer can open the sheet where Excel left it. The selected tab was
already parsed (`tabSelected`); the layout now exposes it.

## AutoFilter imported (2026-10-10)

`<autoFilter ref>` with its `<filterColumn>` children now lands in
`Worksheet.auto_filter` (`AutoFilter { range, columns }`): value lists
(`<filters blank><filter val>`) and custom comparisons (`<customFilters
and><customFilter operator val>`) are modelled; top10, dynamic, colour,
icon and date-group criteria flag the column `other` so a viewer can show
it as filtered without re-evaluating it. Table (`<table>` part) filters
are not read. Nothing in the evaluator uses it.
