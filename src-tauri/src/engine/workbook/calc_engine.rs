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
use std::sync::{Arc, Mutex, MutexGuard};
use std::time::Instant;

use ironcalc::base::cell::CellValue;
use ironcalc::base::expressions::token::get_error_by_english_name;
use ironcalc::base::types::FormulaValue;
use ironcalc::base::Model;
use ironcalc::import::load_from_xlsx;
use serde::{Deserialize, Serialize};
use tauri::State;

type Models = Arc<Mutex<HashMap<String, Loaded>>>;

#[derive(Default)]
pub struct CalcEngineState {
    inner: Models,
}

struct Loaded {
    model: Model<'static>,
    path: String,
    sheets: Vec<String>,
    /// File-cached values of formula cells, captured before the first
    /// evaluation (see `gridpath_engine_core::Cached`). Empty for models
    /// that did not come from a file.
    cached: gridpath_engine_core::Cached,
    /// Per-sheet memo of the structural index served to `describe_workbook`
    /// / `find_rows`; mutating commands mark the sheets they touch dirty.
    index: gridpath_engine_core::index::IndexCache,
    /// Loaded with `defer_eval`: the model still holds the file's cached
    /// values and no evaluation has run yet. Settled by
    /// `calc_engine_evaluate_pending` (the open path, after the grid is
    /// up) or by the first command that needs computed values.
    pending_eval: bool,
}

/// Locks the model map, recovering from lock poisoning: after a caught engine
/// panic the affected tab is removed anyway, so the poison flag carries no
/// information worth propagating.
fn lock_recover(inner: &Models) -> MutexGuard<'_, HashMap<String, Loaded>> {
    inner.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

#[derive(Serialize)]
pub struct CalcLoadResult {
    pub sheets: Vec<String>,
    pub load_ms: u64,
    pub eval_ms: u64,
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
}

#[derive(Serialize)]
pub struct CalcSnapshot {
    pub ms: u64,
    pub cells: Vec<CellDelta>,
}

#[derive(Serialize)]
pub struct CalcTabStatus {
    pub tab: String,
    pub path: String,
    pub sheets: Vec<String>,
}

fn cell_value_to_json(value: CellValue) -> serde_json::Value {
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
/// engine owns — Canalyst exports don't even carry cached values for them.
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

fn deltas_for(loaded: &Loaded, keys: impl IntoIterator<Item = (u32, i32, i32)>) -> Vec<CellDelta> {
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
fn settle_pending(loaded: &mut Loaded) -> Vec<CellDelta> {
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
async fn run_contained<T, F>(inner: Models, tab: String, f: F) -> Result<T, String>
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
) -> Result<CalcLoadResult, String> {
    let key = tab.clone();
    let defer = defer_eval.unwrap_or(false);
    run_contained(state.inner.clone(), tab, move |inner| {
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

        let sheets = model.workbook.get_worksheet_names();
        lock_recover(inner).insert(
            key,
            Loaded {
                model,
                path,
                sheets: sheets.clone(),
                cached,
                index: Default::default(),
                pending_eval: defer,
            },
        );
        Ok(CalcLoadResult {
            sheets,
            load_ms,
            eval_ms,
        })
    })
    .await
}

/// Applies one cell edit (`value` uses spreadsheet input syntax: `=SUM(..)`,
/// `42`, `hello`, empty string clears) and incrementally recalculates.
/// Returns the cells whose values changed (dirty closure + the edited cell).
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
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;

        let settled = settle_pending(loaded);
        let t = Instant::now();
        loaded.model.set_user_input(sheet, row, column, value)?;
        let changed_keys = loaded.model.evaluate_edited(&[(sheet, row, column)]);
        let recalc_ms = t.elapsed().as_millis() as u64;
        loaded.index.mark_dirty(sheet);
        for k in &changed_keys {
            loaded.index.mark_dirty(k.0);
        }

        let mut changed = settled;
        changed.extend(deltas_for(
            loaded,
            std::iter::once((sheet, row, column)).chain(changed_keys),
        ));
        Ok(CalcEditResult { recalc_ms, changed })
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
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;

        let settled = settle_pending(loaded);
        let t = Instant::now();
        let mut keys = Vec::with_capacity(edits.len());
        for e in &edits {
            loaded
                .model
                .set_user_input(e.sheet, e.row, e.column, e.value.clone())?;
            keys.push((e.sheet, e.row, e.column));
        }
        let changed_keys = loaded.model.evaluate_edited(&keys);
        let recalc_ms = t.elapsed().as_millis() as u64;
        for k in keys.iter().chain(changed_keys.iter()) {
            loaded.index.mark_dirty(k.0);
        }
        log::info!(
            "[calc-engine] edit batch: {} cells → recalc {} ms, {} dependents changed",
            keys.len(),
            recalc_ms,
            changed_keys.len()
        );

        let mut changed = settled;
        changed.extend(deltas_for(loaded, keys.into_iter().chain(changed_keys)));
        Ok(CalcEditResult { recalc_ms, changed })
    })
    .await
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
/// after a shift) and the full snapshot is returned.
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
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let t = Instant::now();
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        match op.as_str() {
            "insert_rows" => loaded.model.insert_rows(sheet, index, count)?,
            "delete_rows" => loaded.model.delete_rows(sheet, index, count)?,
            "insert_columns" => loaded.model.insert_columns(sheet, index, count)?,
            "delete_columns" => loaded.model.delete_columns(sheet, index, count)?,
            other => return Err(format!("unknown structural op: {other}")),
        }
        let t_eval = Instant::now();
        loaded.model.evaluate();
        loaded.pending_eval = false;
        let eval_ms = t_eval.elapsed().as_millis();
        loaded.index.mark_all_dirty();
        // Where each pre-op formula cell lives now (grid did the same shift).
        let (rows, insert) = match op.as_str() {
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
        Ok(CalcSnapshot {
            ms: t.elapsed().as_millis() as u64,
            cells,
        })
    })
    .await
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
        let names = model.workbook.get_worksheet_names();
        lock_recover(inner).insert(
            key,
            Loaded {
                model,
                path: String::new(),
                sheets: names.clone(),
                cached: Default::default(),
                index: Default::default(),
                pending_eval: false,
            },
        );
        Ok(CalcLoadResult {
            sheets: names,
            load_ms: 0,
            eval_ms: 0,
        })
    })
    .await
}

#[derive(Serialize)]
pub struct CalcSheetsResult {
    pub sheets: Vec<String>,
    pub cells: Vec<CellDelta>,
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
            guard.insert(
                key.clone(),
                Loaded {
                    model,
                    path: String::new(),
                    sheets: Vec::new(),
                    cached: Default::default(),
                    index: Default::default(),
                    pending_eval: false,
                },
            );
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
        for name in existing {
            if !want.contains(&name.as_str()) {
                model.delete_sheet_by_name(&name)?;
            }
        }

        let names = model.workbook.get_worksheet_names();
        let kept: Vec<u32> = sheets
            .iter()
            .filter(|s| s.keep)
            .filter_map(|s| names.iter().position(|n| n == &s.name).map(|i| i as u32))
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
        Ok(CalcResyncResult {
            sheets: loaded.sheets.clone(),
            cells,
            seeded,
            ms: t.elapsed().as_millis() as u64,
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
    run_contained(state.inner.clone(), tab, move |inner| {
        let mut guard = lock_recover(inner);
        let loaded = guard.get_mut(&key).ok_or("no workbook loaded for tab")?;
        let before = gridpath_engine_core::diff::formula_values(&loaded.model);
        let deleted_index = if op == "delete" { loaded.sheets.iter().position(|n| *n == name).map(|i| i as u32) } else { None };
        match op.as_str() {
            "add" => loaded.model.add_sheet(&name)?,
            "rename" => loaded
                .model
                .rename_sheet(&name, new_name.as_deref().ok_or("rename needs new_name")?)?,
            "delete" => loaded.model.delete_sheet_by_name(&name)?,
            other => return Err(format!("unknown sheet op: {other}")),
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
        Ok(CalcSheetsResult {
            sheets: loaded.sheets.clone(),
            cells,
        })
    })
    .await
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
        Ok(CalcSnapshot { ms: t.elapsed().as_millis() as u64, cells })
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
        Ok(CalcSnapshot { ms: t.elapsed().as_millis() as u64, cells })
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

/// Drops the model for a tab (tab closed).
#[tauri::command]
pub async fn calc_engine_unload(state: State<'_, CalcEngineState>, tab: String) -> Result<(), String> {
    let inner = state.inner.clone();
    tauri::async_runtime::spawn_blocking(move || {
        lock_recover(&inner).remove(&tab);
    })
    .await
    .map_err(|e| format!("calc task join error: {e}"))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn loaded_from(model: Model<'static>) -> Loaded {
        let sheets = model.workbook.get_worksheet_names();
        Loaded { model, path: String::new(), sheets, cached: Default::default(), index: Default::default(), pending_eval: false }
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
                "=_xll.BDP(\"MSFT\",\"PX_LAST\")",
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
