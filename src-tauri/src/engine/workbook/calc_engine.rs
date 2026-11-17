//! IronCalc-backed calculation engine (see vendor/ironcalc/).
//!
//! Holds one loaded workbook model PER TAB. Two frontend consumers:
//!
//! * shadow mode (`ironcalcShadow.ts`): mirrors edits and cross-checks
//!   against Univer's engine, nothing user-visible;
//! * engine mode (`ironcalcEngine.ts`): IronCalc is the calculator — Univer's
//!   formula execution is disabled and the deltas returned from
//!   `calc_engine_edit` / `calc_engine_snapshot` are written into the grid.
//!
//! All model-touching commands are async + spawn_blocking (multi-second work
//! must not block the main thread) and wrap the model in catch_unwind: a
//! panicking engine drops that tab's model and surfaces a command error
//! instead of aborting the app.

use std::collections::HashMap;
use std::panic::AssertUnwindSafe;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex, MutexGuard};
use std::time::Instant;

use ironcalc::base::cell::CellValue;
use ironcalc::base::expressions::token::get_error_by_english_name;
use ironcalc::base::new_empty::is_link_sheet_name;
use ironcalc::base::types::FormulaValue;
use ironcalc::base::Model;
use ironcalc::import::load_from_xlsx;
use serde::{Deserialize, Serialize};
use tauri::State;

use super::workspace::{self, LinkEntry, LinkReport, Stamp, TabDeltas};

/// Every loaded model: one per tab, under the tab's id, and the workbooks
/// that open ones read and that have no tab, under `bg:<path>` (see
/// `workspace.rs`).
pub(super) type Models = Arc<Mutex<HashMap<String, Loaded>>>;

#[derive(Default)]
pub struct CalcEngineState {
    pub(super) inner: Models,
    pub(super) workspace: workspace::Shared,
}

pub(super) struct Loaded {
    pub(super) model: Model<'static>,
    pub(super) path: String,
    pub(super) sheets: Vec<String>,
    /// File-cached values of formula cells, captured before the first
    /// evaluation (see `gridpath_engine_core::Cached`). Empty for models
    /// that did not come from a file.
    pub(super) cached: gridpath_engine_core::Cached,
    /// Per-sheet memo of the structural index served to `describe_workbook`
    /// / `find_rows`; mutating commands mark the sheets they touch dirty.
    pub(super) index: gridpath_engine_core::index::IndexCache,
    /// Loaded with `defer_eval`: the model still holds the file's cached
    /// values and no evaluation has run yet. Settled by
    /// `calc_engine_evaluate_pending` (the open path, after the grid is
    /// up) or by the first command that needs computed values.
    pub(super) pending_eval: bool,
    /// The hidden copies this workbook keeps of sheets of other workbooks,
    /// one per sheet its formulas read (see `workspace.rs`).
    pub(super) links: Vec<LinkEntry>,
    /// Opened with links: formulas typed here may name other workbooks.
    pub(super) linkable: bool,
    /// The file as it was when this model was read from it.
    pub(super) stamp: Option<Stamp>,
    /// Its formulas followed rows that moved, or a sheet that was renamed,
    /// in a workbook it reads, and the file does not have that yet.
    pub(super) needs_saving: bool,
    /// That happened at some point since the file was read: a save writes
    /// every formula that names another workbook, not only the edited ones.
    pub(super) rewrite_links: bool,
    /// Order of loading. Should two tabs hold the same file, the later one
    /// stands for it.
    pub(super) seq: u64,
    /// The file's native charts as the grid draws them (see
    /// `xlsx_patch::charts`): read once when a tab opens the file, kept in
    /// step with row, column and sheet changes, their data read from the
    /// model whenever it is asked for. Empty for a model without a tab.
    pub(super) charts: Vec<super::xlsx_patch::charts::Chart>,
}

static NEXT_SEQ: AtomicU64 = AtomicU64::new(1);

pub(super) fn next_seq() -> u64 {
    NEXT_SEQ.fetch_add(1, Ordering::Relaxed)
}

impl Loaded {
    pub(super) fn new(model: Model<'static>, path: String, pending_eval: bool) -> Self {
        let sheets = model.workbook.get_worksheet_names();
        Loaded {
            model,
            path,
            sheets,
            cached: Default::default(),
            index: Default::default(),
            pending_eval,
            links: Vec::new(),
            linkable: false,
            stamp: None,
            needs_saving: false,
            rewrite_links: false,
            seq: next_seq(),
            charts: Vec::new(),
        }
    }
}

/// The native charts of the file at `path`, for a tab that has just opened
/// it. A file whose charts cannot be read opens without them: nothing here
/// may fail the open, or (a panic) take the tab's model with it.
fn read_charts(path: &str) -> Vec<super::xlsx_patch::charts::Chart> {
    let t = Instant::now();
    let read = std::panic::catch_unwind(|| {
        std::fs::read(path)
            .map_err(|e| e.to_string())
            .and_then(|bytes| super::xlsx_patch::charts::read(&bytes).map_err(|e| e.to_string()))
    })
    .unwrap_or_else(|_| Err("the chart reader failed".to_string()));
    match read {
        Ok(charts) => {
            if !charts.is_empty() {
                let drawn = charts.iter().filter(|c| c.unsupported.is_none()).count();
                log::info!("[calc-engine] charts: {} ({} drawn) read in {} ms", charts.len(), drawn, t.elapsed().as_millis());
            }
            charts
        }
        Err(e) => {
            log::warn!("[calc-engine] charts of {path} not read: {e}");
            Vec::new()
        }
    }
}

/// Fills `charts` from the live model: series names and category labels as
/// the sheet shows them, values as numbers, and, for a chart made in this
/// session, the number formats of the cells it plots.
fn fill_charts(loaded: &Loaded, charts: &mut [super::xlsx_patch::charts::Chart]) {
    use super::xlsx_patch::charts::Range;
    let m = &loaded.model;
    let sheet_of = |range: &Range| m.workbook.worksheets.iter().position(|s| s.name.eq_ignore_ascii_case(&range.sheet)).map(|i| i as u32);
    let read = |range: &Range, text: bool| -> Option<Vec<serde_json::Value>> {
        let sheet = sheet_of(range)?;
        Some(
            range
                .cells()
                .map(|(r, c)| {
                    let (row, column) = (r as i32 + 1, c as i32 + 1);
                    if text {
                        m.get_formatted_cell_value(sheet, row, column).map(serde_json::Value::String).unwrap_or(serde_json::Value::Null)
                    } else {
                        m.get_cell_value_by_index(sheet, row, column).map(cell_value_to_json).unwrap_or(serde_json::Value::Null)
                    }
                })
                .collect(),
        )
    };
    super::xlsx_patch::charts::fill_live(charts, &read);
    let format_of = |range: &Range| -> Option<String> {
        let sheet = sheet_of(range)?;
        m.get_style_for_cell(sheet, range.r0 as i32 + 1, range.c0 as i32 + 1).ok().map(|style| style.num_fmt)
    };
    super::xlsx_patch::charts::link_formats(charts, &format_of);
}

/// The tab's charts with their data read from the live model. Before the
/// first evaluate the file's own cached chart data stands.
#[tauri::command]
pub async fn calc_engine_charts(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<Vec<super::xlsx_patch::charts::Chart>, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let mut charts = loaded.charts.clone();
        if charts.is_empty() || loaded.pending_eval {
            return Ok(charts);
        }
        fill_charts(loaded, &mut charts);
        Ok(charts)
    })
    .await
}

/// Adds a chart to the tab's workbook: one made in this session, which the
/// next save writes into the file (see `xlsx_patch::chart_write`). Gives
/// the chart as it is drawn, with the data it plots, or says why it cannot
/// be made. An id that is already there is replaced (a redo).
pub(super) fn add_chart(
    inner: &Models,
    key: &str,
    spec: &super::xlsx_patch::charts::ChartSpec,
) -> Result<super::xlsx_patch::charts::Chart, String> {
    let mut guard = lock_recover(inner);
    let loaded = guard.get_mut(key).ok_or("no workbook loaded for tab")?;
    // Sheets under the names the workbook spells them; never the hidden
    // copy of another workbook's sheet.
    let own: Vec<String> = loaded.model.workbook.worksheets.iter().map(|s| s.name.clone()).filter(|n| !n.starts_with('[')).collect();
    let sheet_named = |name: &str| own.iter().find(|n| n.eq_ignore_ascii_case(name)).cloned();
    let chart = super::xlsx_patch::charts::build(spec, &sheet_named).map_err(|e| {
        if e.contains("there is no sheet") {
            format!("{e}. Sheets: {}", own.join(", "))
        } else {
            e
        }
    })?;
    if chart.id.trim().is_empty() {
        return Err("the chart needs an id".into());
    }
    let mut filled = vec![chart.clone()];
    fill_charts(loaded, &mut filled);
    let filled = filled.remove(0);
    if let Some(why) = &filled.unsupported {
        return Err(why.clone());
    }
    loaded.charts.retain(|c| c.id != chart.id);
    loaded.charts.push(chart);
    Ok(filled)
}

#[tauri::command]
pub async fn calc_engine_add_chart(
    state: State<'_, CalcEngineState>,
    tab: String,
    spec: super::xlsx_patch::charts::ChartSpec,
) -> Result<super::xlsx_patch::charts::Chart, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| add_chart(inner, &key, &spec)).await
}

/// Takes a chart made in this session away again (its batch was rejected
/// or undone). A chart of the file is not removed: `false`.
#[tauri::command]
pub async fn calc_engine_remove_chart(state: State<'_, CalcEngineState>, tab: String, id: String) -> Result<bool, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let before = loaded.charts.len();
        loaded.charts.retain(|c| !(c.id == id && c.part.is_empty()));
        Ok(loaded.charts.len() < before)
    })
    .await
}

/// Locks the model map, recovering from lock poisoning: after a caught engine
/// panic the affected tab is removed anyway, so the poison flag carries no
/// information worth propagating.
pub(super) fn lock_recover(inner: &Models) -> MutexGuard<'_, HashMap<String, Loaded>> {
    inner.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

#[derive(Serialize)]
pub struct CalcLoadResult {
    pub sheets: Vec<String>,
    pub load_ms: u64,
    pub eval_ms: u64,
    /// What changed in other open workbooks: the ones that read this file.
    pub others: Vec<TabDeltas>,
    /// The model was held without a tab and has changes the file lacks.
    pub needs_saving: bool,
}

/// One computed cell value. Row/column are 1-based (IronCalc convention).
#[derive(Serialize, Debug)]
pub struct CellDelta {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
    pub value: serde_json::Value,
}

#[derive(Serialize)]
pub struct CalcEditResult {
    pub recalc_ms: u64,
    pub changed: Vec<CellDelta>,
    /// What changed in other open workbooks: the ones that read this one.
    pub others: Vec<TabDeltas>,
    /// The sheet list, when the edit added a hidden copy of a linked sheet.
    pub sheets: Option<Vec<String>>,
    /// What became of the links to other workbooks the edit spelled.
    pub links: Option<LinkReport>,
}

#[derive(Serialize, Default)]
pub struct CalcSnapshot {
    pub ms: u64,
    pub cells: Vec<CellDelta>,
    /// What changed in other open workbooks: the ones that read this one.
    pub others: Vec<TabDeltas>,
    /// Tabs whose formulas moved to follow this workbook: they have changes
    /// to save now.
    pub moved: Vec<String>,
}

#[derive(Serialize)]
pub struct CalcTabStatus {
    pub tab: String,
    pub path: String,
    pub sheets: Vec<String>,
}

pub(super) fn cell_value_to_json(value: CellValue) -> serde_json::Value {
    match value {
        CellValue::None => serde_json::Value::Null,
        CellValue::String(s) => serde_json::Value::String(s),
        CellValue::Number(n) => serde_json::Number::from_f64(n)
            .map(serde_json::Value::Number)
            // NaN/inf have no JSON representation; surface them as strings.
            .unwrap_or_else(|| serde_json::Value::String(n.to_string())),
        CellValue::Boolean(b) => serde_json::Value::Bool(b),
    }
}

/// Values of every formula cell in the model. Literal cells are already
/// correct in the grid (they came from the file); formula cells are what the
/// engine owns — some vendor exports don't even carry cached values for them.
fn formula_cell_deltas(loaded: &Loaded) -> Vec<CellDelta> {
    let mut out = Vec::new();
    for (sheet_index, ws) in loaded.model.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_none() {
                    continue;
                }
                let value = loaded
                    .model
                    .get_cell_value_by_index(sheet_index as u32, *row, *column)
                    .map(cell_value_to_json)
                    .unwrap_or(serde_json::Value::Null);
                out.push(CellDelta {
                    sheet: sheet_index as u32,
                    row: *row,
                    column: *column,
                    value,
                });
            }
        }
    }
    out
}

pub(super) fn deltas_for(loaded: &Loaded, keys: impl IntoIterator<Item = (u32, i32, i32)>) -> Vec<CellDelta> {
    keys.into_iter()
        .map(|(s, r, c)| CellDelta {
            sheet: s,
            row: r,
            column: c,
            value: loaded
                .model
                .get_cell_value_by_index(s, r, c)
                .map(cell_value_to_json)
                .unwrap_or(serde_json::Value::Null),
        })
        .collect()
}

/// Runs the deferred first evaluate if one is pending and returns the
/// formula cells whose computed value differs from the file's cached value
/// (usually a handful: constructs our engine evaluates differently, stale
/// caches from manual-calc files). No-op once settled.
pub(super) fn settle_pending(loaded: &mut Loaded) -> Vec<CellDelta> {
    if !loaded.pending_eval {
        return Vec::new();
    }
    let t = Instant::now();
    if loaded.cached.is_empty() {
        loaded.cached = gridpath_engine_core::build_cached(&loaded.model);
        log::info!("[calc-engine] cached-value map (deferred): {} ms", t.elapsed().as_millis());
    }
    loaded.model.evaluate();
    loaded.pending_eval = false;
    loaded.index.mark_all_dirty();
    // The file cache IS the before side: no value map is built at all.
    let changed: Vec<CellDelta> = gridpath_engine_core::diff::changed_vs_cached(&loaded.model, &loaded.cached)
        .into_iter()
        .map(|((sheet, row, column), v)| CellDelta { sheet, row, column, value: v.to_json() })
        .collect();
    let pruned = gridpath_engine_core::diff::prune_cached(&loaded.model, &mut loaded.cached);
    log::info!(
        "[calc-engine] deferred first evaluate: {} ms, {} values differ from the file's cache; cache map pruned by {} to {} entries",
        t.elapsed().as_millis(),
        changed.len(),
        pruned,
        loaded.cached.len()
    );
    changed
}

/// Formula cells whose value differs from `before` (see
/// `gridpath_engine_core::diff::changed_since`), as grid deltas.
fn changed_deltas(
    before: gridpath_engine_core::diff::ValueMap,
    loaded: &Loaded,
    remap: impl Fn((u32, i32, i32)) -> Option<(u32, i32, i32)>,
) -> Vec<CellDelta> {
    gridpath_engine_core::diff::changed_since(before, &loaded.model, remap)
        .into_iter()
        .map(|((sheet, row, column), v)| CellDelta { sheet, row, column, value: v.to_json() })
        .collect()
}

/// Runs `f` on a blocking thread with panic containment; a panic drops the
/// tab's model (it may be mid-mutation) and returns an error.
pub(super) async fn run_contained<T, F>(inner: Models, tab: String, f: F) -> Result<T, String>
where
    T: Send + 'static,
    F: FnOnce(&Models) -> Result<T, String> + Send + 'static,
{
    tauri::async_runtime::spawn_blocking(move || {
        let result = std::panic::catch_unwind(AssertUnwindSafe(|| f(&inner)));
        result.unwrap_or_else(|_| {
            lock_recover(&inner).remove(&tab);
            Err(format!(
                "ironcalc engine panicked (tab {tab}); model dropped — reopen the workbook"
            ))
        })
    })
    .await
    .map_err(|e| format!("calc task join error: {e}"))?
}

/// Loads `path` into a fresh model for `tab` and runs the initial full
/// evaluation. Replaces any model previously held for that tab.
#[tauri::command]
pub async fn calc_engine_load(
    state: State<'_, CalcEngineState>,
    tab: String,
    path: String,
    defer_eval: Option<bool>,
    links: Option<bool>,
) -> Result<CalcLoadResult, String> {
    let key = tab.clone();
    let defer = defer_eval.unwrap_or(false);
    let shared = state.workspace.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        // With links: the workbooks this file was saved reading are found
        // and loaded too, and formulas typed later may name others.
        if links.unwrap_or(false) {
            let opened = workspace::open_tab(inner, &shared, &key, &path, defer)?;
            let charts = read_charts(&path);
            if let Some(loaded) = lock_recover(inner).get_mut(&key) {
                loaded.charts = charts;
            }
            return Ok(CalcLoadResult {
                sheets: opened.sheets,
                load_ms: opened.load_ms,
                eval_ms: opened.eval_ms,
                others: opened.others,
                needs_saving: opened.needs_saving,
            });
        }
        let t = Instant::now();
        let mut model = load_from_xlsx(&path, "en", "UTC", "en")
            .map_err(|e| format!("ironcalc load: {e}"))?;
        let load_ms = t.elapsed().as_millis() as u64;
        // The file-cached value map is the "before" side of the deferred
        // first evaluate and the fallback for unevaluable cells. Building it
        // walks every formula cell (seconds on a 900k-formula model), so a
        // deferred open builds it when the settle runs — after the grid's
        // first tiles are on screen — not before the layout is returned.
        let t_cached = Instant::now();
        let cached = if defer { gridpath_engine_core::Cached::new() } else { gridpath_engine_core::build_cached(&model) };
        if !defer {
            log::info!("[calc-engine] cached-value map: {} ms", t_cached.elapsed().as_millis());
        }

        // Deferred first evaluate: the file's cached values go to the grid
        // at once (what Excel shows on open too); the evaluate runs after
        // the grid is up and only cells whose value differs are re-sent.
        let t = Instant::now();
        let mut cached = cached;
        if !defer {
            model.evaluate();
            gridpath_engine_core::diff::prune_cached(&model, &mut cached);
        }
        let eval_ms = t.elapsed().as_millis() as u64;
        log::info!("[calc-engine] load {}: parse {} ms, evaluate {} ms{}", path, load_ms, eval_ms, if defer { " (deferred)" } else { "" });

        let charts = read_charts(&path);
        let mut loaded = Loaded::new(model, path, defer);
        loaded.cached = cached;
        loaded.charts = charts;
        let sheets = loaded.sheets.clone();
        let others = workspace::install_tab(&mut lock_recover(inner), &key, loaded);
        Ok(CalcLoadResult {
            sheets,
            load_ms,
            eval_ms,
            others,
            needs_saving: false,
        })
    })
    .await
}

/// Applies one cell edit (`value` uses spreadsheet input syntax: `=SUM(..)`,
/// `42`, `hello`, empty string clears) and incrementally recalculates.
/// Returns the cells whose values changed (dirty closure + the edited cell).
/// A formula that names another workbook, `='[NOW.xlsx]Model'!B2`, is
/// linked to it (see `workspace::apply_edits`).
#[tauri::command]
pub async fn calc_engine_edit(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheet: u32,
    row: i32,
    column: i32,
    value: String,
) -> Result<CalcEditResult, String> {
    let key = tab.clone();
    let shared = state.workspace.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        workspace::apply_edits(inner, &shared, &key, vec![EditCell { sheet, row, column, value }])
    })
    .await
}

#[derive(Deserialize)]
pub struct EditCell {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
    pub value: String,
}

/// Applies several cell edits as ONE recalculation — an undo replay or a
/// paste is a single dirty-closure pass instead of one per cell.
#[tauri::command]
pub async fn calc_engine_edit_batch(
    state: State<'_, CalcEngineState>,
    tab: String,
    edits: Vec<EditCell>,
) -> Result<CalcEditResult, String> {
    let key = tab.clone();
    let shared = state.workspace.clone();
    run_contained(state.inner.clone(), tab, move |inner| workspace::apply_edits(inner, &shared, &key, edits)).await
}

/// Full formula-cell value snapshot — applied to the grid after load, and
/// after structural operations (everything may have shifted).
#[tauri::command]
pub async fn calc_engine_snapshot(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<CalcSnapshot, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let cells = formula_cell_deltas(loaded);
        Ok(CalcSnapshot {
            ms: t.elapsed().as_millis() as u64,
            cells,
            ..Default::default()
        })
    })
    .await
}

/// The whole workbook in the engine's snapshot shape (sheets, cells with
/// formula / value / display / style index, style table, defined names —
/// see `gridpath_engine_core::snapshot_json`), as raw JSON bytes. This is
/// what the grid is built from on open, instead of a second parse of the
/// file with ExcelJS in the webview. Returned as an IPC `Response` so the
/// 100+ MB payload is not JSON-encoded a second time on the way across.
/// `lazy_threshold`: above this many cells, sheets other than the first
/// visible one are sent as skeletons flagged `lazy` (see
/// `gridpath_engine_core::SnapshotOptions`); the grid fetches each through
/// `calc_engine_sheet_cells` when it is shown or written to. On a 2.3M-cell
/// model the grid otherwise holds ~2 GB of cell objects for sheets nobody is
/// looking at.
#[tauri::command]
pub async fn calc_engine_univer_snapshot(
    state: State<'_, CalcEngineState>,
    tab: String,
    lazy_threshold: Option<u64>,
    lazy_sheet_min_cells: Option<u64>,
) -> Result<tauri::ipc::Response, String> {
    let key = tab.clone();
    let bytes = run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let first_visible = loaded
            .model
            .workbook
            .worksheets
            .iter()
            .position(|ws| matches!(ws.state, ironcalc::base::types::SheetState::Visible))
            .unwrap_or(0) as u32;
        let mut opts = gridpath_engine_core::SnapshotOptions {
            lazy_threshold: lazy_threshold.map(|t| t as usize),
            eager_sheets: vec![first_visible],
            // Before the deferred first evaluate, formulas the file has no
            // cached value for are "not computed yet", not `#ERROR!`.
            blank_unevaluated: loaded.pending_eval,
            ..Default::default()
        };
        if let Some(min) = lazy_sheet_min_cells {
            opts.lazy_sheet_min_cells = min as usize;
        }
        let json = gridpath_engine_core::snapshot_json_with(&loaded.model, &loaded.cached, &opts)?;
        log::info!(
            "[calc-engine] univer snapshot: {} MB in {} ms (lazy threshold {:?})",
            json.len() / 1_000_000,
            t.elapsed().as_millis(),
            lazy_threshold
        );
        Ok(json.into_bytes())
    })
    .await?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// Cells and styles of one sheet, for a sheet the snapshot sent lazily.
#[tauri::command]
pub async fn calc_engine_sheet_cells(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheet: u32,
) -> Result<tauri::ipc::Response, String> {
    let key = tab.clone();
    let bytes = run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let json = gridpath_engine_core::sheet_cells_json_with(&loaded.model, &loaded.cached, sheet, loaded.pending_eval)?;
        log::info!(
            "[calc-engine] sheet {} cells: {} MB in {} ms",
            sheet,
            json.len() / 1_000_000,
            t.elapsed().as_millis()
        );
        Ok(json.into_bytes())
    })
    .await?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// Structural operation: `op` is one of `insert_rows`, `delete_rows`,
/// `insert_columns`, `delete_columns`; `index` is 1-based. References shift
/// like in Excel; a full re-evaluation follows (the dependency map is stale
/// after a shift) and the values that changed are returned.
#[tauri::command]
pub async fn calc_engine_structural(
    state: State<'_, CalcEngineState>,
    tab: String,
    op: String,
    sheet: u32,
    index: i32,
    count: i32,
) -> Result<CalcSnapshot, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| structural(inner, &key, &op, sheet, index, count)).await
}

/// What `calc_engine_structural` does, for the model under `key`.
pub(super) fn structural(inner: &Models, key: &str, op: &str, sheet: u32, index: i32, count: i32) -> Result<CalcSnapshot, String> {
    let mut guard = lock_recover(inner);
    let loaded = guard.get_mut(key).ok_or("no workbook loaded for tab")?;
    let t = Instant::now();
    let before = gridpath_engine_core::diff::formula_values(&loaded.model);
    match op {
        "insert_rows" => loaded.model.insert_rows(sheet, index, count)?,
        "delete_rows" => loaded.model.delete_rows(sheet, index, count)?,
        "insert_columns" => loaded.model.insert_columns(sheet, index, count)?,
        "delete_columns" => loaded.model.delete_columns(sheet, index, count)?,
        other => return Err(format!("unknown structural op: {other}")),
    }
    // The file's saved values follow their cells.
    gridpath_engine_core::shift_cached(
        &mut loaded.cached,
        sheet,
        index,
        if op.starts_with("insert") { count } else { -count },
        op.ends_with("rows"),
    );
    let t_eval = Instant::now();
    loaded.model.evaluate();
    loaded.pending_eval = false;
    let eval_ms = t_eval.elapsed().as_millis();
    loaded.index.mark_all_dirty();
    // Where each pre-op formula cell lives now (grid did the same shift).
    let (rows, insert) = match op {
        "insert_rows" => (true, true),
        "delete_rows" => (true, false),
        "insert_columns" => (false, true),
        _ => (false, false),
    };
    let cells = changed_deltas(before, loaded, |(s, r, c)| {
        if s != sheet {
            return Some((s, r, c));
        }
        let axis = if rows { r } else { c };
        let moved = if insert {
            if axis >= index { axis + count } else { axis }
        } else if axis >= index && axis < index + count {
            return None;
        } else if axis >= index + count {
            axis - count
        } else {
            axis
        };
        Some(if rows { (s, moved, c) } else { (s, r, moved) })
    });
    log::info!(
        "[calc-engine] structural {op} sheet {sheet} @{index} x{count}: full evaluate {eval_ms} ms, {} values changed",
        cells.len()
    );
    // The workbooks that read this one: links into the sheet follow the
    // cells they pointed at, and every value read is compared again.
    let sheet_name = loaded.model.workbook.worksheets.get(sheet as usize).map(|ws| ws.name.clone()).unwrap_or_default();
    // The charts' ranges and anchors follow too.
    if !loaded.charts.is_empty() {
        let shift = super::xlsx_patch::refs::RowColShift { rows, start: (index - 1).max(0) as u32, count: count.max(0) as u32, insert };
        if let Err(e) = super::xlsx_patch::charts::shift(&mut loaded.charts, &sheet_name, &shift) {
            log::warn!("[calc-engine] charts not moved with {op} on {sheet_name}: {e}");
        }
    }
    let path = loaded.path.clone();
    let (mut others, moved) = workspace::mirror_structural(&mut guard, key, &sheet_name, op, index, count);
    workspace::merge(&mut others, workspace::propagate(&mut guard, &path, workspace::Change::All));
    Ok(CalcSnapshot {
        ms: t.elapsed().as_millis() as u64,
        cells,
        others,
        moved,
    })
}

/// Creates an EMPTY model for a tab — untitled/blank workbooks that the
/// agent builds into have no file to import from. `sheets` is the grid's
/// current sheet-name list (usually just ["Sheet1"]).
#[tauri::command]
pub async fn calc_engine_new(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheets: Vec<String>,
) -> Result<CalcLoadResult, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut model = Model::new_empty("untitled", "en", "UTC", "en")?;
        let existing = model.workbook.get_worksheet_names();
        if let (Some(target), Some(current)) = (sheets.first(), existing.first()) {
            if target != current {
                model.rename_sheet(current, target)?;
            }
        }
        for name in sheets.iter().skip(1) {
            model.add_sheet(name)?;
        }
        model.evaluate();
        let loaded = Loaded::new(model, String::new(), false);
        let names = loaded.sheets.clone();
        lock_recover(inner).insert(key, loaded);
        Ok(CalcLoadResult {
            sheets: names,
            load_ms: 0,
            eval_ms: 0,
            others: Vec::new(),
            needs_saving: false,
        })
    })
    .await
}

#[derive(Serialize)]
pub struct CalcSheetsResult {
    pub sheets: Vec<String>,
    pub cells: Vec<CellDelta>,
    /// What changed in other open workbooks: the ones that read this one.
    pub others: Vec<TabDeltas>,
    /// Tabs whose formulas were respelled to follow a renamed sheet.
    pub moved: Vec<String>,
}

#[derive(Deserialize)]
pub struct ResyncCell {
    pub row: i32,    // 1-based
    pub column: i32, // 1-based
    /// Formula text ("=SUM(..)"); when set, `value` is the grid's current
    /// (cached) value for the cell.
    pub formula: Option<String>,
    #[serde(default)]
    pub value: serde_json::Value,
}

#[derive(Deserialize)]
pub struct ResyncSheet {
    pub name: String,
    pub cells: Vec<ResyncCell>,
    /// Lazy sheet (grid holds a skeleton only): leave the engine's cells as
    /// they are instead of clearing and re-seeding from the grid.
    #[serde(default)]
    pub keep: bool,
}

#[derive(Serialize)]
pub struct CalcResyncResult {
    pub sheets: Vec<String>,
    pub cells: Vec<CellDelta>,
    pub seeded: usize,
    pub ms: u64,
    /// What changed in other open workbooks: the ones that read this one.
    pub others: Vec<TabDeltas>,
}

/// Cached value for a reseeded formula cell. Strings that name an Excel
/// error become real errors (they re-evaluate honestly rather than freeze —
/// same rule the importer applies); other strings are text.
fn json_to_formula_value(v: &serde_json::Value, origin: String) -> FormulaValue {
    match v {
        serde_json::Value::Bool(b) => FormulaValue::Boolean(*b),
        serde_json::Value::Number(n) => n
            .as_f64()
            .map(FormulaValue::Number)
            .unwrap_or(FormulaValue::Unevaluated),
        serde_json::Value::String(s) => match get_error_by_english_name(s) {
            Some(ei) => FormulaValue::Error {
                ei,
                o: origin,
                m: s.clone(),
            },
            None => FormulaValue::Text(s.clone()),
        },
        _ => FormulaValue::Unevaluated,
    }
}

fn json_to_cell_value(v: serde_json::Value) -> Option<CellValue> {
    match v {
        serde_json::Value::Bool(b) => Some(CellValue::Boolean(b)),
        serde_json::Value::Number(n) => n.as_f64().map(CellValue::Number),
        serde_json::Value::String(s) => {
            if s.is_empty() {
                None
            } else {
                Some(CellValue::String(s))
            }
        }
        _ => None,
    }
}

/// Rebuilds the tab's model IN PLACE from a grid content dump — the recovery
/// path for operations the command bridge cannot mirror (sort, range moves,
/// unrecognized commands). Sheets are reconciled by name, cell content is
/// cleared and reseeded (formulas keep their grid value as the cached value,
/// so unevaluable add-in/external-ref formulas stay frozen at real data),
/// defined names and styles survive because the model is never rebuilt from
/// scratch. Ends with a full evaluation and returns the formula-cell
/// snapshot for the grid.
#[tauri::command]
pub async fn calc_engine_resync(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheets: Vec<ResyncSheet>,
) -> Result<CalcResyncResult, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        if sheets.is_empty() {
            return Err("resync needs at least one sheet".into());
        }
        let t = Instant::now();
        let mut guard = lock_recover(inner);
        // A tab whose load failed has no model — resync can heal that too.
        if !guard.contains_key(&key) {
            let model = Model::new_empty("resynced", "en", "UTC", "en")?;
            guard.insert(key.clone(), Loaded::new(model, String::new(), false));
        }
        let loaded = guard.get_mut(&key).expect("just inserted");
        let model = &mut loaded.model;

        // Reconcile the sheet set by name, grid as truth. Add before delete
        // so the "cannot delete only sheet" guard can never fire mid-way.
        let want: Vec<&str> = sheets.iter().map(|s| s.name.as_str()).collect();
        let existing = model.workbook.get_worksheet_names();
        for name in &want {
            if !existing.iter().any(|e| e == name) {
                model.add_sheet(name)?;
            }
        }
        // The hidden copies of linked sheets are the engine's own: the grid
        // does not list them, and they stay as they are.
        for name in existing {
            if !want.contains(&name.as_str()) && !is_link_sheet_name(&name) {
                model.delete_sheet_by_name(&name)?;
            }
        }

        let names = model.workbook.get_worksheet_names();
        let kept: Vec<u32> = sheets
            .iter()
            .filter(|s| s.keep)
            .filter_map(|s| names.iter().position(|n| n == &s.name).map(|i| i as u32))
            .chain(names.iter().enumerate().filter(|(_, n)| is_link_sheet_name(n)).map(|(i, _)| i as u32))
            .collect();
        model.resync_clear_contents_except(&kept);

        let mut seeded = 0usize;
        for sheet in &sheets {
            if sheet.keep {
                continue;
            }
            let idx = names
                .iter()
                .position(|n| n == &sheet.name)
                .ok_or_else(|| format!("resync: sheet \"{}\" vanished", sheet.name))?
                as u32;
            for cell in &sheet.cells {
                match &cell.formula {
                    Some(f) => {
                        let origin = format!("{}!R{}C{}", sheet.name, cell.row, cell.column);
                        model.seed_formula_with_cached_value(
                            idx,
                            cell.row,
                            cell.column,
                            f,
                            json_to_formula_value(&cell.value, origin),
                        )?;
                        seeded += 1;
                    }
                    None => {
                        if let Some(v) = json_to_cell_value(cell.value.clone()) {
                            model.seed_cell_literal(idx, cell.row, cell.column, v)?;
                            seeded += 1;
                        }
                    }
                }
            }
        }
        let t_eval = Instant::now();
        model.evaluate();
        loaded.pending_eval = false;
        log::info!("[calc-engine] resync: {} sheets, {} cells seeded, full evaluate {} ms", sheets.len(), seeded, t_eval.elapsed().as_millis());
        loaded.sheets = model.workbook.get_worksheet_names();
        loaded.index.mark_all_dirty();
        let cells = formula_cell_deltas(loaded);
        let names = loaded.sheets.clone();
        let path = loaded.path.clone();
        let others = workspace::propagate(&mut guard, &path, workspace::Change::All);
        Ok(CalcResyncResult {
            sheets: names,
            cells,
            seeded,
            ms: t.elapsed().as_millis() as u64,
            others,
        })
    })
    .await
}

/// Sheet structure: `op` is `add`, `rename` (with `new_name`) or `delete`,
/// all by name. Renames rewrite referencing formulas and deletes turn them
/// into errors, so a full re-evaluation + snapshot follows (sheet ops are
/// rare enough that simplicity wins).
#[tauri::command]
pub async fn calc_engine_sheet_op(
    state: State<'_, CalcEngineState>,
    tab: String,
    op: String,
    name: String,
    new_name: Option<String>,
) -> Result<CalcSheetsResult, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| sheet_op(inner, &key, &op, &name, new_name.as_deref())).await
}

/// What `calc_engine_sheet_op` does, for the model under `key`.
pub(super) fn sheet_op(inner: &Models, key: &str, op: &str, name: &str, new_name: Option<&str>) -> Result<CalcSheetsResult, String> {
    let mut guard = lock_recover(inner);
    let loaded = guard.get_mut(key).ok_or("no workbook loaded for tab")?;
    let before = gridpath_engine_core::diff::formula_values(&loaded.model);
    let deleted_index = if op == "delete" { loaded.sheets.iter().position(|n| n == name).map(|i| i as u32) } else { None };
    match op {
        "add" => loaded.model.add_sheet(name)?,
        "rename" => loaded
            .model
            .rename_sheet(name, new_name.ok_or("rename needs new_name")?)?,
        "delete" => loaded.model.delete_sheet_by_name(name)?,
        other => return Err(format!("unknown sheet op: {other}")),
    }
    if let Some(deleted) = deleted_index {
        gridpath_engine_core::remove_sheet_from_cached(&mut loaded.cached, deleted);
    }
    let t_eval = Instant::now();
    loaded.model.evaluate();
    loaded.pending_eval = false;
    let eval_ms = t_eval.elapsed().as_millis();
    loaded.sheets = loaded.model.workbook.get_worksheet_names();
    loaded.index.mark_all_dirty();
    // Adding appends (no index moves); deleting shifts later sheets down.
    let cells = changed_deltas(before, loaded, |(s, r, c)| match deleted_index {
        Some(d) if s == d => None,
        Some(d) if s > d => Some((s - 1, r, c)),
        _ => Some((s, r, c)),
    });
    log::info!("[calc-engine] sheet op {op} {name:?}: full evaluate {eval_ms} ms, {} values changed", cells.len());
    match (op, new_name) {
        ("rename", Some(new)) => {
            if let Err(e) = super::xlsx_patch::charts::rename_sheet(&mut loaded.charts, name, new) {
                log::warn!("[calc-engine] charts not renamed with {name:?}: {e}");
            }
        }
        ("delete", _) => super::xlsx_patch::charts::remove_sheet(&mut loaded.charts, name),
        _ => {}
    }
    let names = loaded.sheets.clone();
    let path = loaded.path.clone();
    // The workbooks that read this one: a renamed sheet is renamed in
    // their links, and every value read is compared again. A deleted
    // sheet keeps its last values there (the list shows the link as
    // broken).
    let (mut others, moved) = match (op, new_name) {
        ("rename", Some(new)) => workspace::mirror_rename(&mut guard, key, name, new),
        _ => (Vec::new(), Vec::new()),
    };
    workspace::merge(&mut others, workspace::propagate(&mut guard, &path, workspace::Change::All));
    Ok(CalcSheetsResult {
        sheets: names,
        cells,
        others,
        moved,
    })
}

/// Settles a deferred first evaluate (see `calc_engine_load`): runs it if
/// still pending and returns only the formula cells whose computed value
/// differs from the file's cached value. Idempotent.
#[tauri::command]
pub async fn calc_engine_evaluate_pending(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<CalcSnapshot, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let cells = settle_pending(loaded);
        let path = loaded.path.clone();
        let keys = cells.iter().map(|d| (d.sheet, d.row, d.column)).collect();
        let others = workspace::propagate(&mut guard, &path, workspace::Change::Cells(keys));
        Ok(CalcSnapshot { ms: t.elapsed().as_millis() as u64, cells, others, moved: Vec::new() })
    })
    .await
}

/// Layout for the viewport grid (`gridpath_engine_core::window`): per
/// sheet the frozen panes, merges, custom widths/heights and used extent,
/// plus the workbook style table. Small; fetched once per open.
#[tauri::command]
pub async fn calc_engine_workbook_layout(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<tauri::ipc::Response, String> {
    let key = tab.clone();
    let bytes = run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let json = gridpath_engine_core::window::workbook_layout_json(&loaded.model)?;
        log::info!("[calc-engine] workbook layout: {} KB in {} ms", json.len() / 1000, t.elapsed().as_millis());
        Ok(json.into_bytes())
    })
    .await?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// One rectangle of cells for the viewport grid (1-based, inclusive).
#[tauri::command]
pub async fn calc_engine_sheet_window(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheet: u32,
    r0: i32,
    r1: i32,
    c0: i32,
    c1: i32,
) -> Result<tauri::ipc::Response, String> {
    let key = tab.clone();
    let bytes = run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let json = gridpath_engine_core::window::sheet_window_json(
            &loaded.model,
            &loaded.cached,
            sheet,
            r0,
            r1,
            c0,
            c1,
            loaded.pending_eval,
        )?;
        Ok(json.into_bytes())
    })
    .await?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// Structural index of the workbook (`gridpath_engine_core::index`): per
/// sheet the used range, header row, row labels, sections and formula
/// runs — the same JSON shape `buildWorkbookIndex` produces in TypeScript,
/// so `describe_workbook` / `find_rows` consume either. Memoized per sheet:
/// only sheets written since the last call are rebuilt.
#[tauri::command]
pub async fn calc_engine_workbook_index(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<tauri::ipc::Response, String> {
    let key = tab.clone();
    let bytes = run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let (json, rebuilt) = loaded.index.workbook_index_json(&loaded.model)?;
        log::info!(
            "[calc-engine] workbook index: {} sheets rebuilt, {} KB in {} ms",
            rebuilt,
            json.len() / 1000,
            t.elapsed().as_millis()
        );
        Ok(json.into_bytes())
    })
    .await?;
    Ok(tauri::ipc::Response::new(bytes))
}

/// `query_rows`: filter / project / aggregate one sheet's rows in the
/// engine (`gridpath_engine_core::query`). `query` is the tool's input
/// object with `sheet` already resolved to an index.
#[tauri::command]
pub async fn calc_engine_query_rows(
    state: State<'_, CalcEngineState>,
    tab: String,
    query: serde_json::Value,
) -> Result<serde_json::Value, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let input: gridpath_engine_core::query::QueryInput =
            serde_json::from_value(query).map_err(|e| format!("bad query: {e}"))?;
        let t = Instant::now();
        let out = gridpath_engine_core::query::query_rows(&loaded.model, &input)?;
        log::info!(
            "[calc-engine] query_rows sheet {}: {} matched in {} ms",
            input.sheet,
            out.get("matched").and_then(|v| v.as_u64()).unwrap_or(0),
            t.elapsed().as_millis()
        );
        Ok(out)
    })
    .await
}

#[derive(Serialize)]
pub struct CalcHealthSheet {
    pub sheet: String,
    pub divergent: usize,
}

/// Mirrors `packages/gridpath/src/core/calcHealth.ts` (`CalcHealth`).
#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CalcHealthReport {
    pub checked: usize,
    pub divergent: usize,
    pub by_sheet: Vec<CalcHealthSheet>,
    pub truncated: bool,
}

/// Live-vs-file divergence for the agent's calc-health line (rule 17b):
/// formula cells whose live result is unusable (empty, error, or 0 where
/// the file cached a number) while the file carried a usable value. After
/// the first evaluate the cached-value map holds exactly the cells whose
/// live value is missing (`prune_cached`), so the report is a scan of that
/// map. None while the first evaluate is still pending.
#[tauri::command]
pub async fn calc_engine_calc_health(
    state: State<'_, CalcEngineState>,
    tab: String,
) -> Result<Option<CalcHealthReport>, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        if loaded.pending_eval {
            return Ok(None);
        }
        let mut checked = 0usize;
        for ws in &loaded.model.workbook.worksheets {
            for row in ws.sheet_data.values() {
                checked += row.values().filter(|c| c.get_formula().is_some()).count();
            }
        }
        let mut per_sheet: HashMap<u32, usize> = HashMap::new();
        for ((sheet, _row, _col), saved) in loaded.cached.iter() {
            let usable = match saved {
                serde_json::Value::Null => false,
                serde_json::Value::String(s) => !(s.is_empty() || s.starts_with('#')),
                _ => true,
            };
            if usable {
                *per_sheet.entry(*sheet).or_insert(0) += 1;
            }
        }
        let mut by_sheet: Vec<CalcHealthSheet> = per_sheet
            .into_iter()
            .map(|(sheet, divergent)| CalcHealthSheet {
                sheet: loaded.sheets.get(sheet as usize).cloned().unwrap_or_else(|| sheet.to_string()),
                divergent,
            })
            .collect();
        by_sheet.sort_by(|a, b| b.divergent.cmp(&a.divergent).then_with(|| a.sheet.cmp(&b.sheet)));
        let divergent = by_sheet.iter().map(|s| s.divergent).sum();
        Ok(Some(CalcHealthReport { checked, divergent, by_sheet, truncated: false }))
    })
    .await
}

#[derive(Serialize)]
pub struct FindHit {
    pub sheet: u32,
    pub row: i32,
    pub column: i32,
    pub text: String,
}

/// Workbook-wide text search for the viewport grid's find bar (the grid
/// holds only tiles). Case-insensitive substring match over each cell's
/// formula text when it has one, else its value; `sheet` limits the scan
/// to one sheet. Results in sheet, row, column order, capped.
#[tauri::command]
pub async fn calc_engine_find(
    state: State<'_, CalcEngineState>,
    tab: String,
    query: String,
    sheet: Option<u32>,
    limit: Option<usize>,
) -> Result<Vec<FindHit>, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        let needle = query.to_lowercase();
        let cap = limit.unwrap_or(500).min(5_000);
        let mut hits = Vec::new();
        if needle.is_empty() {
            return Ok(hits);
        }
        let m = &loaded.model;
        for (si, ws) in m.workbook.worksheets.iter().enumerate() {
            if let Some(only) = sheet {
                if only as usize != si {
                    continue;
                }
            }
            // Hidden copies of sheets of other workbooks are not the user's.
            if is_link_sheet_name(&ws.name) {
                continue;
            }
            let mut rows: Vec<&i32> = ws.sheet_data.keys().collect();
            rows.sort_unstable();
            for r in rows {
                let row = &ws.sheet_data[r];
                let mut cols: Vec<&i32> = row.keys().collect();
                cols.sort_unstable();
                for c in cols {
                    let cell = &row[c];
                    let text = if cell.get_formula().is_some() {
                        m.get_cell_formula(si as u32, *r, *c).ok().flatten().unwrap_or_default()
                    } else {
                        match m.get_cell_value_by_index(si as u32, *r, *c) {
                            Ok(CellValue::String(s)) => s,
                            Ok(CellValue::Number(n)) => n.to_string(),
                            Ok(CellValue::Boolean(b)) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                            _ => String::new(),
                        }
                    };
                    if text.to_lowercase().contains(&needle) {
                        hits.push(FindHit { sheet: si as u32, row: *r, column: *c, text: text.chars().take(120).collect() });
                        if hits.len() >= cap {
                            return Ok(hits);
                        }
                    }
                }
            }
        }
        Ok(hits)
    })
    .await
}

/// Partial style change from the grid or the agent (CellFormatShape fields).
#[derive(Deserialize, Default)]
pub struct StylePatchWire {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    pub font_color: Option<String>,
    /// Some(None) clears the fill; Some(Some(rgb)) sets it.
    #[serde(default, deserialize_with = "deserialize_double_option")]
    pub background_color: Option<Option<String>>,
    pub font_size: Option<i32>,
    pub font_family: Option<String>,
    pub horizontal_align: Option<String>,
    pub vertical_align: Option<String>,
    pub number_format: Option<String>,
    pub wrap_text: Option<bool>,
}

fn deserialize_double_option<'de, D: serde::Deserializer<'de>>(d: D) -> Result<Option<Option<String>>, D::Error> {
    Ok(Some(Option::<String>::deserialize(d)?))
}

/// Applies a partial style to a rectangle (1-based inclusive) so the engine's
/// display strings and style table reflect in-session formatting — the
/// viewport grid draws from those, and the save mirror carries the change to
/// the file independently.
#[tauri::command]
pub async fn calc_engine_set_style(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheet: u32,
    r0: i32,
    r1: i32,
    c0: i32,
    c1: i32,
    patch: StylePatchWire,
) -> Result<usize, String> {
    use ironcalc::base::types::{Alignment, Color, HorizontalAlignment, VerticalAlignment};
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        if r1 < r0 || c1 < c0 || r0 < 1 || c0 < 1 {
            return Err("bad range".into());
        }
        if (r1 - r0 + 1) as i64 * (c1 - c0 + 1) as i64 > 200_000 {
            return Err("range too large for a style change".into());
        }
        let mut n = 0usize;
        for r in r0..=r1 {
            for c in c0..=c1 {
                let mut st = loaded.model.get_style_for_cell(sheet, r, c)?;
                if let Some(b) = patch.bold { st.font.b = b; }
                if let Some(i) = patch.italic { st.font.i = i; }
                if let Some(u) = patch.underline { st.font.u = u; }
                if let Some(s) = patch.strike { st.font.strike = s; }
                if let Some(col) = &patch.font_color { st.font.color = Color::Rgb(col.clone()); }
                if let Some(bg) = &patch.background_color {
                    st.fill.color = match bg { Some(rgb) if !rgb.is_empty() => Color::Rgb(rgb.clone()), _ => Color::None };
                }
                if let Some(sz) = patch.font_size { st.font.sz = sz; }
                if let Some(name) = &patch.font_family { st.font.name = name.clone(); }
                if let Some(nf) = &patch.number_format { st.num_fmt = nf.clone(); }
                if patch.horizontal_align.is_some() || patch.vertical_align.is_some() || patch.wrap_text.is_some() {
                    let mut al = st.alignment.clone().unwrap_or(Alignment { horizontal: HorizontalAlignment::General, vertical: VerticalAlignment::Bottom, wrap_text: false });
                    if let Some(h) = &patch.horizontal_align {
                        al.horizontal = match h.as_str() { "left" => HorizontalAlignment::Left, "center" => HorizontalAlignment::Center, "right" => HorizontalAlignment::Right, _ => HorizontalAlignment::General };
                    }
                    if let Some(v) = &patch.vertical_align {
                        al.vertical = match v.as_str() { "top" => VerticalAlignment::Top, "middle" => VerticalAlignment::Center, _ => VerticalAlignment::Bottom };
                    }
                    if let Some(w) = patch.wrap_text { al.wrap_text = w; }
                    st.alignment = Some(al);
                }
                loaded.model.set_cell_style(sheet, r, c, &st)?;
                n += 1;
            }
        }
        Ok(n)
    })
    .await
}

/// Creates or replaces a workbook-scoped defined name in the engine and
/// recalculates, returning the formula cells whose value moved (formulas
/// that referenced the name resolve now). Full evaluate: names are rare.
#[tauri::command]
pub async fn calc_engine_define_name(
    state: State<'_, CalcEngineState>,
    tab: String,
    name: String,
    formula: String,
) -> Result<CalcSnapshot, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        let f = formula.trim().trim_start_matches('=').to_string();
        let existing = loaded.model.get_defined_name_list().into_iter().any(|(n, scope, _)| scope.is_none() && n.eq_ignore_ascii_case(&name));
        if existing {
            loaded.model.update_defined_name(&name, None, &name, None, &f)?;
        } else {
            loaded.model.new_defined_name(&name, None, &f)?;
        }
        // Formulas that referenced the name before it existed were parsed
        // as #NAME?; re-parse everything (and evaluate).
        loaded.model.reset_parsed_structures();
        loaded.pending_eval = false;
        loaded.index.mark_all_dirty();
        let cells: Vec<CellDelta> = gridpath_engine_core::diff::changed_since(before, &loaded.model, Some)
            .into_iter()
            .map(|((sheet, row, column), v)| CellDelta { sheet, row, column, value: v.to_json() })
            .collect();
        log::info!("[calc-engine] defined name {name} = {f}: {} ms, {} values changed", t.elapsed().as_millis(), cells.len());
        let path = loaded.path.clone();
        let keys = cells.iter().map(|d| (d.sheet, d.row, d.column)).collect();
        let others = workspace::propagate(&mut guard, &path, workspace::Change::Cells(keys));
        Ok(CalcSnapshot { ms: t.elapsed().as_millis() as u64, cells, others, moved: Vec::new() })
    })
    .await
}

/// Value of a single cell — used by the shadow cross-check.
///
/// Every command that takes the engine lock is async and runs on the
/// blocking pool: a sync command runs on the main thread, and while a
/// multi-second evaluate holds the lock that would freeze the whole app
/// (the close-tab unload used to do exactly that).
#[tauri::command]
pub async fn calc_engine_get_cell(
    state: State<'_, CalcEngineState>,
    tab: String,
    sheet: u32,
    row: i32,
    column: i32,
) -> Result<serde_json::Value, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let guard = lock_recover(inner);
        let loaded = guard.get(&key).ok_or("no workbook loaded for tab")?;
        loaded
            .model
            .get_cell_value_by_index(sheet, row, column)
            .map(cell_value_to_json)
    })
    .await
}

#[tauri::command]
pub async fn calc_engine_status(state: State<'_, CalcEngineState>) -> Result<Vec<CalcTabStatus>, String> {
    let inner = state.inner.clone();
    tauri::async_runtime::spawn_blocking(move || {
        let guard = lock_recover(&inner);
        guard
            .iter()
            // Tabs only: the workbooks loaded without one are listed by
            // `workspace_status`.
            .filter(|(tab, _)| !workspace::is_bg(tab))
            .map(|(tab, loaded)| CalcTabStatus {
                tab: tab.clone(),
                path: loaded.path.clone(),
                sheets: loaded.sheets.clone(),
            })
            .collect::<Vec<_>>()
    })
    .await
    .map_err(|e| format!("calc task join error: {e}"))
}

/// Drops the model for a tab (tab closed). A workbook that other open
/// workbooks read stays loaded without the tab: as it is, or, when the tab
/// had unsaved changes (`dirty`), as its file has it. Returns what that
/// changed in the workbooks that read it.
#[tauri::command]
pub async fn calc_engine_unload(
    state: State<'_, CalcEngineState>,
    tab: String,
    dirty: Option<bool>,
) -> Result<Vec<TabDeltas>, String> {
    let inner = state.inner.clone();
    let shared = state.workspace.clone();
    tauri::async_runtime::spawn_blocking(move || {
        let closing = tab.clone();
        std::panic::catch_unwind(AssertUnwindSafe(|| workspace::close_tab(&inner, &shared, &closing, dirty.unwrap_or(false))))
            .unwrap_or_else(|_| {
                lock_recover(&inner).remove(&tab);
                Vec::new()
            })
    })
    .await
    .map_err(|e| format!("calc task join error: {e}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn loaded_from(model: Model<'static>) -> Loaded {
        Loaded::new(model, String::new(), false)
    }

    /// A deferred load serves the file's cached values; settling runs the
    /// evaluate once and reports only cells whose value moved. Here the
    /// "cached" values are the unevaluated state, so every formula moves;
    /// a second settle is a no-op.
    #[test]
    fn deferred_evaluate_settles_once_and_reports_moved_values() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.set_user_input(0, 1, 1, "2".to_string()).unwrap();
        m.set_user_input(0, 1, 2, "=A1*3".to_string()).unwrap();
        m.set_user_input(0, 2, 2, "=B1+1".to_string()).unwrap();
        let mut loaded = loaded_from(m);
        loaded.pending_eval = true;
        let changed = settle_pending(&mut loaded);
        let mut got: Vec<(i32, i32, serde_json::Value)> = changed.iter().map(|d| (d.row, d.column, d.value.clone())).collect();
        got.sort_by(|a, b| (a.0, a.1).cmp(&(b.0, b.1)));
        assert_eq!(got, vec![(1, 2, serde_json::json!(6.0)), (2, 2, serde_json::json!(7.0))]);
        assert!(!loaded.pending_eval);
        assert!(settle_pending(&mut loaded).is_empty());
    }

    /// A chart made in a session: built from ranges, drawn with the live
    /// data, kept in step with rows that move, and taken away again.
    #[test]
    fn a_chart_made_in_the_session_plots_live_cells_and_follows_them() {
        use super::super::xlsx_patch::charts::{ChartSpec, SeriesSpec};
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.rename_sheet_by_index(0, "P&L").unwrap();
        for (c, year) in ["FY24", "FY25", "FY26"].iter().enumerate() {
            let col = c as i32 + 2;
            m.set_user_input(0, 1, col, year.to_string()).unwrap();
            m.set_user_input(0, 2, col, format!("{}", 100 * (c + 1))).unwrap();
            m.set_user_input(0, 3, col, format!("={}2*0.25/{}2", ["B", "C", "D"][c], ["B", "C", "D"][c])).unwrap();
        }
        m.set_user_input(0, 2, 1, "Revenue".to_string()).unwrap();
        m.set_user_input(0, 3, 4, "30%".to_string()).unwrap();
        m.evaluate();
        let inner: Models = Arc::new(Mutex::new(HashMap::new()));
        inner.lock().unwrap().insert("t".to_string(), loaded_from(m));

        let spec = ChartSpec {
            id: "new:1".into(),
            sheet: "p&l".into(),
            row: 5,
            col: 1,
            title: "Revenue".into(),
            kind: "column".into(),
            categories: Some("B1:D1".into()),
            series: vec![
                SeriesSpec { name_ref: Some("A2".into()), values: "B2:D2".into(), ..Default::default() },
                SeriesSpec { name: Some("Margin".into()), values: "'P&L'!B3:D3".into(), kind: Some("line".into()), secondary: true, ..Default::default() },
            ],
            ..Default::default()
        };
        let chart = add_chart(&inner, "t", &spec).unwrap();
        assert_eq!(chart.sheet, "P&L");
        assert_eq!(chart.series[0].name, "Revenue");
        assert_eq!(chart.series[0].categories, vec![serde_json::json!("FY24"), serde_json::json!("FY25"), serde_json::json!("FY26")]);
        assert_eq!(chart.series[0].values, vec![Some(100.), Some(200.), Some(300.)]);
        assert_eq!(chart.series[1].values, vec![Some(0.25), Some(0.25), Some(0.3)]);
        assert_eq!(chart.series[1].values_ref.as_deref(), Some("'P&L'!$B$3:$D$3"));
        // The same id again replaces the chart (a redo), it does not add one.
        add_chart(&inner, "t", &spec).unwrap();
        assert_eq!(inner.lock().unwrap()["t"].charts.len(), 1);

        // What cannot be made says why, and adds nothing.
        let on = |change: &dyn Fn(&mut ChartSpec)| {
            let mut s = spec.clone();
            s.id = "new:2".into();
            change(&mut s);
            add_chart(&inner, "t", &s).unwrap_err()
        };
        assert!(on(&|s| s.sheet = "Charts".into()).contains("there is no sheet \"Charts\" to put the chart on. Sheets: P&L"));
        assert!(on(&|s| s.series[0].values = "Other!B2:D2".into()).contains("series 1: values: there is no sheet \"Other\""));
        assert!(on(&|s| s.series[0].values = "B2:D3".into()).contains("one row or one column"));
        assert_eq!(inner.lock().unwrap()["t"].charts.len(), 1);

        // A row inserted above the data: the chart's ranges and place follow.
        structural(&inner, "t", "insert_rows", 0, 1, 1).unwrap();
        let moved = inner.lock().unwrap()["t"].charts[0].clone();
        assert_eq!(moved.series[0].values_ref.as_deref(), Some("'P&L'!$B$3:$D$3"));
        assert_eq!(moved.series[0].name_ref.as_deref(), Some("'P&L'!$A$3"));
        assert_eq!(moved.anchor.from.row, 6);
    }

    /// A sheet add changes no value; a row insert moves formulas but only
    /// the cells whose VALUE changed come back (here: the SUM that now
    /// spans the inserted blank row keeps its value, the moved cells keep
    /// theirs, so nothing is re-sent); a delete that drops a referenced
    /// row changes exactly the dependents.
    #[test]
    fn sheet_and_structural_ops_return_only_changed_values() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        for r in 1..=3 {
            m.set_user_input(0, r, 1, format!("{r}")).unwrap();
            m.set_user_input(0, r, 2, format!("=A{r}*10")).unwrap();
        }
        m.set_user_input(0, 5, 1, "=SUM(A1:A3)".to_string()).unwrap();
        m.evaluate();
        let mut loaded = loaded_from(m);

        // add sheet → no value changes
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        loaded.model.add_sheet("New").unwrap();
        loaded.model.evaluate();
        assert!(changed_deltas(before, &loaded, |k| Some(k)).is_empty());

        // insert a row at 2 → formulas move down one; values identical
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        loaded.model.insert_rows(0, 2, 1).unwrap();
        loaded.model.evaluate();
        let changed = changed_deltas(before, &loaded, |(s, r, c)| Some(if s == 0 && r >= 2 { (s, r + 1, c) } else { (s, r, c) }));
        assert!(changed.is_empty(), "{changed:?}");

        // delete row 3 (A3=2 after the insert) → the SUM drops 6 → 4; the
        // deleted row's formula vanishes; every other formula moves up with
        // its value intact and is not re-sent.
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        loaded.model.delete_rows(0, 3, 1).unwrap();
        loaded.model.evaluate();
        let changed = changed_deltas(before, &loaded, |(s, r, c)| {
            if s != 0 { return Some((s, r, c)); }
            if r == 3 { None } else if r > 3 { Some((s, r - 1, c)) } else { Some((s, r, c)) }
        });
        let mut keys: Vec<(i32, i32, serde_json::Value)> = changed.iter().map(|d| (d.row, d.column, d.value.clone())).collect();
        keys.sort_by(|a, b| (a.0, a.1).cmp(&(b.0, b.1)));
        assert_eq!(keys, vec![(5, 1, serde_json::json!(4.0))]);
    }

    /// End-to-end sanity on a small generated workbook: build via ironcalc,
    /// save, reload, edit incrementally and confirm the dirty closure carries
    /// the recomputed dependent; then a structural insert shifts references.
    #[test]
    fn incremental_edit_returns_dependents() {
        let dir = std::env::temp_dir().join("gridpath-calc-engine-test");
        std::fs::create_dir_all(&dir).unwrap();
        let path = dir.join("small.xlsx");
        let _ = std::fs::remove_file(&path);

        let mut model = Model::new_empty("small", "en", "UTC", "en").unwrap();
        model.set_user_input(0, 1, 1, "2".to_string()).unwrap(); // A1
        model.set_user_input(0, 1, 2, "=A1*10".to_string()).unwrap(); // B1
        model.set_user_input(0, 1, 3, "=B1+5".to_string()).unwrap(); // C1
        model.evaluate();
        ironcalc::export::save_to_xlsx(&model, path.to_str().unwrap()).unwrap();

        let mut model = load_from_xlsx(path.to_str().unwrap(), "en", "UTC", "en").unwrap();
        model.evaluate();
        model.set_user_input(0, 1, 1, "3".to_string()).unwrap();
        let changed = model.evaluate_edited(&[(0, 1, 1)]);

        assert!(changed.contains(&(0, 1, 2)), "B1 should be dirty: {changed:?}");
        assert!(changed.contains(&(0, 1, 3)), "C1 should be dirty: {changed:?}");
        assert_eq!(
            model.get_cell_value_by_index(0, 1, 3).unwrap(),
            CellValue::Number(35.0)
        );

        // Structural: inserting a row above shifts everything down one; the
        // formulas must follow.
        model.insert_rows(0, 1, 1).unwrap();
        model.evaluate();
        assert_eq!(
            model.get_cell_value_by_index(0, 2, 3).unwrap(),
            CellValue::Number(35.0)
        );
    }

    /// Resync semantics at the vendor level: clearing keeps defined names,
    /// reseeded evaluable formulas recompute, unevaluable ones freeze at the
    /// cached value they were seeded with.
    #[test]
    fn resync_reseed_preserves_names_and_frozen_values() {
        let mut model = Model::new_empty("resync", "en", "UTC", "en").unwrap();
        model
            .new_defined_name("GROWTH", None, "Sheet1!$D$1")
            .expect("defined name");
        model.set_user_input(0, 1, 4, "1.5".to_string()).unwrap(); // D1 (GROWTH)
        model.set_user_input(0, 1, 1, "10".to_string()).unwrap(); // A1
        model.set_user_input(0, 1, 2, "=A1*GROWTH".to_string()).unwrap(); // B1
        model.evaluate();
        assert_eq!(
            model.get_cell_value_by_index(0, 1, 2).unwrap(),
            CellValue::Number(15.0)
        );

        // Simulate drift + resync: clear, reseed from "the grid" with an
        // extra add-in style formula carrying a cached market value.
        model.resync_clear_contents();
        model.seed_cell_literal(0, 1, 4, CellValue::Number(1.5)).unwrap();
        model.seed_cell_literal(0, 1, 1, CellValue::Number(20.0)).unwrap();
        model
            .seed_formula_with_cached_value(
                0,
                1,
                2,
                "=A1*GROWTH",
                FormulaValue::Unevaluated,
            )
            .unwrap();
        model
            .seed_formula_with_cached_value(
                0,
                1,
                3,
                "=_xll.BDP(\"ACME\",\"PX_LAST\")",
                FormulaValue::Number(432.5),
            )
            .unwrap();
        // Literal string that looks like a number must stay text.
        model
            .seed_cell_literal(0, 2, 1, CellValue::String("123".to_string()))
            .unwrap();
        model.evaluate();

        // Defined name survived the clear; the formula recomputed.
        assert_eq!(
            model.get_cell_value_by_index(0, 1, 2).unwrap(),
            CellValue::Number(30.0)
        );
        // Unevaluable add-in call froze at its cached value.
        assert_eq!(
            model.get_cell_value_by_index(0, 1, 3).unwrap(),
            CellValue::Number(432.5)
        );
        assert_eq!(
            model.get_cell_value_by_index(0, 2, 1).unwrap(),
            CellValue::String("123".to_string())
        );
    }
}
