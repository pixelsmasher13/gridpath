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
