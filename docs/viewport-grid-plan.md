# Viewport grid: what is done and what remains

Status as of 2026-10-08, branch `viewport-grid`. This is the plan for taking
the viewport grid from tonight's prototype to the only grid in GridPath, and
for the engine work that goes with it.

## 1. Why

Univer keeps a JavaScript object per cell. On the datacenter model (17 MB,
2.37M cells, 926k formulas) that is 2.5–2.8 GB of webview memory and minutes
in "Building grid"; the machine swaps and the app feels dead. Excel opens the
same file comfortably because it keeps the workbook in a native core and
paints only the visible window.

The viewport grid is that architecture: IronCalc (our Rust fork) is the only
copy of the workbook; the webview holds an LRU of 64×32 tiles near the
viewport, fetched from the engine as the user scrolls, drawn on one canvas.

Measured tonight on the same model:

| | Univer | viewport grid |
|---|---|---|
| webview memory | 2.5–2.8 GB | 86–237 MB (JS heap flat at 23 MB in the harness) |
| grid visible after | minutes | ~5 s (file parse) |
| page-scroll frame, biggest sheet | — | 16.7 ms avg, 21 ms worst |

## 2. What exists today

Opt-in: `localStorage.setItem("gridpath.viewportGrid","1")`, reopen the file.
Only xlsx tabs use it; untitled tabs keep Univer.

**Engine** (`crates/engine-core/src/window.rs`, Tauri commands
`calc_engine_workbook_layout` / `calc_engine_sheet_window`, WASM
`workbookLayout` / `sheetWindow` / `openDeferred` / `evaluatePending`):
layout (geometry, merges, frozen panes, style table, ~190 KB) and windowed
cell reads with display strings. Deferred first evaluate; the grid lands its
first tiles before the evaluate starts.

**Grid** (`src/screens/SpreadsheetScreen/components/ViewportGrid.tsx`,
`viewportModel.ts`): canvas renderer matching Univer's conventions (24 px
rows, 88 px columns, Arial 11 pt, bottom-aligned, pixel-snapped), two-pass
painting, Excel overflow rules, frozen panes (text overflows across the split
at scroll 0), merges, fills, borders, font colour/size/bold/italic/underline/
strike, selection, headers, sheet tabs, formula bar.

Editing: typing, formula bar, Enter/Tab/Escape/F2/Delete, double-click,
copy (TSV), paste, undo/redo (cell writes), header-drag resize, header click
selects row/column.

Agent surface (UniverGridHandle): setCell/setCells, clearRange, set/get
format (overlay, display only), column width / row height, hide/show rows
and columns, create/delete/rename sheet, insert/delete rows and columns
(engine), getCell/getRangeData from tiles, prefetchRange (read_range and
copy_range use it), whenCalculated, jumpToCell, getActiveSelection, save via
the shared surgical patch builder (edit log + sizes merged into the mirror).
Opening context for the agent comes from the engine index
(`contextFromIndex`).

**Engine fixes found through this work** (both on `viewport-grid`, both
belong on `main`):

- `Model::evaluate` restarted the whole evaluation for every out-of-order
  dynamic-array anchor (n² restarts). A workbook GridPath had saved with 25k
  agent formulas never finished its first evaluate. Now 7.2 s.
- `calc_engine_unload`, `get_cell`, `status` ran on the main thread; while
  an evaluate held the engine lock, closing a tab froze the app.

## 3. Work remaining

Estimates are working days for one person. Order is the recommended order.

### Phase 0 — ship the engine fixes (hours)

- Cherry-pick 33f63b3 and 26845f2 to `main`, push the private origin, rebuild
  installers. Any GridPath-saved workbook with many agent formulas hangs on
  reopen without the first one.
- Add a regression test: a sheet with ~2,000 scalar dynamic-array formulas
  referencing each other evaluates in bounded time.

### Phase 1 — agent parity (1 day) — done 2026-10-09 except the stage_data items

The tools are the same on both grids; the opening context is not. Until this
is done the agent is blinder on the viewport grid.

Done: the preview is captured from engine windows (`getPreviewSnapshot` →
`captureWorkbookContextWith`) in the Univer-path format with the same memo
and base+delta (changed cells anywhere in the workbook ride the delta via
the grid's dirty set); file-saved cells come from the window's `fileSaved`
flag (`getFileSavedCell`, read_range `file_saved`); calc health is an
engine command (`calc_engine_calc_health`, served from the pruned cache
map). Measured on the 2.3M-cell model in the harness: preview assembled in
1.4 s (635k cells), captured in 160 ms, 733 KB preview (25 sheets × 32 KB
cap), a later edit ships as a 1.5 s delta on the same base.

- Opening preview from engine windows in the exact format the Univer path
  produced (`captureContext.ts` builder fed by a snapshot-shaped object
  assembled from tiles), so prompt bytes are identical and the cached block
  still hits. Budgeted per sheet as before; hidden sheets stay label-only.
- Base + delta on later prompts from the grid's edit log instead of a
  snapshot diff.
- `getFileSavedCell` / `getExternalPin` from the window's `fileSaved` cells
  (rule 17b live-vs-file warnings).
- `getCalcHealth` computed engine-side (live values unusable vs file cache).
- `captureCellBand` with real formats (review previews). (open)
- `getSheetSnapshot` / `restoreSheetSnapshot` for stage_data, or an
  engine-side equivalent. (open)

**Capture parity, measured 2026-10-10.** The harness now mounts the real
Univer grid (`.harness-univer/univer-real.html`) and `compare-capture.mjs`
diffs the agent context from both grids on the same file. Datacenter v4:
15 of 25 sheets byte-identical; the other 10 differ only by (a) `wrap`
in format annotations, which the Univer grid dropped and the viewport grid
reports, and (b) display rounding of exact .5 ties, where the viewport
grid now matches Excel (half away from zero; engine fix 49777c9) and
Univer's JS formatter rounds to even. Three parity fixes came out of the
diff: font family/size/colour in annotations, snapshot capacity counts,
and the `@` marker stripped from agent-facing formulas.

### Phase 2 — daily-use usability (1 day) — done 2026-10-09

All five items below are in (commits 5a67cd3, f84c125), verified in the
harness: Ctrl+arrow block jumps with on-demand tile fetch, Ctrl+End,
click-to-insert references with dashed outlines while typing a formula,
drag auto-scroll, Cmd+X, right-click menu (cut/copy/paste/clear, insert and
delete rows/columns recorded into the save mirror), selection
count/sum/avg. Row/column inserts and deletes now shift the session's edit
log, overlay, dirty set and size overrides.

- Ctrl+arrow to the edge of a data block, Ctrl+Home/End, Ctrl+Shift+arrow.
- Click-to-insert cell references while typing a formula, with coloured
  reference highlights.
- Auto-scroll when dragging a selection past the viewport edge.
- Cut; right-click menu with insert/delete rows and columns, clear, copy,
  paste.
- Selection sum / count / average in the status bar.

### Phase 3 — feature completeness (3–4 days) — in progress

Done 2026-10-09 (ce10237): mergeCells/unmerge, freezePanes/unfreeze,
setRangeBorders, notes (marker + hover), sortRange, wrapped text,
defineName/getDefinedNameRef (display/lookup; the engine does not learn new
names yet). Also done: find bar (engine `calc_engine_find`), sheet-tab menu
with inline rename, hide/unhide from headers, fill handle, delete-sheet
undo snapshot, in-session formats applied in the engine
(`calc_engine_set_style`, so number formats render), defined names created
in-session reach the engine (`calc_engine_define_name` + formula re-parse).
Also done: border styles (weight/colour/dash/double), conditional
formatting (matched fill/font deltas and data bars come with each window
cell as `cf`; drawn over the cell's style), hide/unhide sheets end to end,
tab-bar arrows, tab colours, white fills over gridlines, batched tile
fetches with direction prefetch, paste with formats (own copies and HTML
tables), Replace, zoom (50–200%), multi-range selection, Ctrl+A,
Ctrl/Shift+Space. Also done 2026-10-10: context menus close on outside
click/Escape, Insert sheet + Tab color (patcher SheetOp::SetTabColor),
Ctrl+arrow to the grid edge / across merges, and outline groups: the
engine imports outlineLevel + outlinePr (GRIDPATH-CHANGES.md), the layout
carries rowOutline/colOutline, the grid draws Excel's gutters (level
buttons, bars, +/- at the summary row/column) on their own canvas and
collapses/expands through the hide path, so the state persists as hidden
rows/cols (the patcher keeps outlineLevel). Not done: writing `collapsed`
on the summary row, creating/removing groups, expanding an outer group
re-shows collapsed inner groups (Excel keeps them collapsed).
Saved view (2026-10-10): the file opens on Excel's selected tab, and each
sheet's first visit scrolls to the pane's topLeftCell and selects the
saved active cell (engine imports frozenSplit as frozen + topLeftCell);
selection and scroll are kept per sheet across tab switches.
AutoFilter (2026-10-10): the engine imports the sheet's autoFilter (range
+ value-list / custom criteria), the grid draws Excel's dropdown buttons
on the header row (funnel when filtered), the popup has sort, clear,
search and a value checklist, filtering hides/shows the range's rows
through the size path (undoable, persists as hidden rows), and the
criteria are written back through the patcher's AutoFilter op (range or
range+columns; null removes). Filter on/off: toolbar button, context menu
(Filter / Clear filter / Remove filter) and Cmd+Shift+L over the current
region. Not done: custom-filter editor (file criteria are honoured and
shown as checked items; the checklist writes a value list), top10/colour/
dynamic criteria (kept as "other", rows keep their file state), table
filters, shifting the range when rows are inserted above it.
Remaining: data-validation markers (the engine has no DV model yet), icon sets, rich text runs, text
rotation/indent, charts and images, formula autocomplete, drag-to-move.

Handle members still stubbed (each is one engine command plus grid support):

| member | needs |
|---|---|
| mergeCells / unmerge | engine merge op + layout refetch |
| setFreezePanes | engine freeze op + layout refetch |
| setRangeBorders | overlay today; engine style op for persistence |
| setNote / deleteNote / getNote | engine notes + hover rendering |
| defineName / getDefinedNameRef | engine name ops (exist in WASM) |
| sortRange | engine sort or TS sort over window + writes |
| applyOutlineLevel | engine row/col outline |
| number_format on in-session formats | engine style op so display strings update |
| exportBytes (full export) | keep throwing; the in-place patch is the save path |

Rendering: wrapped text (multi-line), rich text runs, conditional formatting
(the engine already runs `evaluate_conditional_formatting`; expose results in
windows), data-validation dropdown markers, hyperlinks, notes markers,
images/charts as placeholders.

Grid UI: fill handle, find (Cmd+F), zoom, sheet-tab add/rename/reorder/
right-click, unhide rows and columns, multi-range selection, in-cell
multi-line editing.

### Phase 4 — retire Univer (2 days) — default flipped 2026-10-10

The viewport grid now serves every xlsx tab (a0f3ece); Univer stays behind
`gridpath.viewportGrid = "0"` for one release and still serves untitled
tabs. The Excel keyboard set is in (Cmd+D/R, Cmd+Enter, B/I/U, number
formats, date/time, copy-above, hide/unhide, insert/delete rows, find and
replace, Shift+Enter/Tab, Option+PageUp/Down). Next: a week of daily use,
then delete Univer and its presets, patches and calc worker; give
untitled tabs an engine-backed viewport grid.

Only once Phases 1–3 are in and the viewport grid has run for a week on
normal models.

- Make the viewport grid the default; keep the Univer path behind the flag
  for one release, then delete `UniverGrid.tsx`, the Univer presets, the
  patch-package patches, the calc worker and the no-worker plugin, the
  snapshot mapper (`mapSnapshot`) and the Univer-shaped snapshot code in the
  engine (keep the index and query paths).
- Review tab, pending-edit tints and the format painter re-pointed at the
  viewport grid.
- Startup: no 0.8 GB worker, smaller bundle, faster first paint.
- Risks to check before deleting: anything the Univer presets rendered that
  the viewport grid does not yet (filters, CF, DV, drawings, find/replace,
  sort UI) is either implemented or consciously dropped.

### Phase 5 — open time (1–2 days) — parallel parse done 2026-10-09

Worksheet parts now parse on all cores (see GRIDPATH-CHANGES.md): v4 parse
2.8–4.1 s → 1.1–1.5 s, snapshot byte-identical to the sequential parse.
The cached-value map (the ~2 s of load bookkeeping) now builds when the
deferred settle runs, after the first tiles are on screen, not before the
layout is returned. Remaining, if needed: lazy per-sheet parse.

Native open of v4 today: parse 2.8 s, formula parse 0.13 s, then ~2 s of
load bookkeeping in the desktop path (cached-value map, per-sheet index).

- Parse sheet XML parts in parallel (independent parts; 25 sheets across
  cores → under 1 s).
- Build the cached-value map and the index lazily.
- Later: lazy per-sheet parse like Excel, if the first two are not enough.

### Phase 6 — robustness and polish (1 day, spread out)

- Tile reads during a long evaluate wait on the engine lock. Either serve
  windows from a read-only copy of the evaluated-before state, or make the
  evaluate cancellable; at minimum, show "calculating" in the status line.
- Degraded engine (load failure): the viewport grid has nothing to show.
  Fall back to Univer for that tab automatically.
- Webview `engineLog` lines still do not reach `GridPath.log`.
- Dynamic-array save writes spills with literal spill cells (would be
  #SPILL! in Excel) — needs engine-reported spill ranges.
- WASM package: the same evaluate fix applies; republish `gridpath` so MCP
  users get it.

## 4. Testing

- Unit: `src/screens/SpreadsheetScreen/__tests__/viewportModel.test.ts`
  (geometry, tile cache, helpers, index context). Add tests for the edit
  log → patch mapping and the overlay merge.
- Harness (`.harness-univer/`, git-excluded): `viewport.html` runs the real
  component against the WASM engine in headless Chromium; `shot.mjs`,
  `probe3.mjs` (set/undo/redo/paste/patch), `probe4.mjs` (formats, sizes,
  visibility, clearRange), `scrollperf.mjs` (frames, heap). `?file=/small.xlsx`
  for fast runs, `?nosettle=1` to skip the first evaluate.
- Desktop: the datacenter v3/v4 files, project-159 (Canalyst, 4.4k INDIRECT),
  one normal 8 MB model; check memory toast, open time, scroll, an agent build
  of a sheet, save and reopen in Excel.
- Eval: rerun the SpreadsheetBench shard and the MCP fixtures once Phase 1 is
  in, since the opening context changes what the agent sees.

## 5. Sequencing and totals

| phase | days | gate |
|---|---|---|
| 0 engine fixes to main | 0.5 | none — do now |
| 1 agent parity | 1 | needed before any agent eval on the new grid |
| 2 daily-use usability | 1 | needed before you use it as your grid |
| 3 feature completeness | 3–4 | needed before Univer can go |
| 4 retire Univer | 2 | a week of daily use on normal models |
| 5 open time | 1–2 | independent; any time |
| 6 robustness | 1 | spread across the above |

About two working weeks to a single grid, with the agent at parity after the
first two days.

## 6. Decisions needed

1. Retire Univer, or keep two grids (Univer for small files, viewport for
   large)? Recommendation: retire; two grids is the clunky end state.
2. Order of Phase 2 versus Phase 5: comfort in the grid first, or open time
   first? Recommendation: Phase 2 first; open time is 5 s today, usable.
3. Whether the opening preview should stay byte-identical to the Univer
   format (cache-friendly, known quality) or be redesigned around the index
   plus on-demand reads (smaller, needs eval). Recommendation: identical
   first, redesign as a separate experiment.
