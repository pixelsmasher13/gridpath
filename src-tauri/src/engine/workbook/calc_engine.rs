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
#[derive(Serialize)]
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
) -> Result<CalcLoadResult, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let t = Instant::now();
        let mut model = load_from_xlsx(&path, "en", "UTC", "en")
            .map_err(|e| format!("ironcalc load: {e}"))?;
        let load_ms = t.elapsed().as_millis() as u64;

        let t = Instant::now();
        model.evaluate();
        let eval_ms = t.elapsed().as_millis() as u64;

        let sheets = model.workbook.get_worksheet_names();
        lock_recover(inner).insert(
            key,
            Loaded {
                model,
                path,
                sheets: sheets.clone(),
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

        let t = Instant::now();
        loaded.model.set_user_input(sheet, row, column, value)?;
        let changed_keys = loaded.model.evaluate_edited(&[(sheet, row, column)]);
        let recalc_ms = t.elapsed().as_millis() as u64;

        let changed = deltas_for(
            loaded,
            std::iter::once((sheet, row, column)).chain(changed_keys),
        );
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

        let changed = deltas_for(loaded, keys.into_iter().chain(changed_keys));
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
        match op.as_str() {
            "insert_rows" => loaded.model.insert_rows(sheet, index, count)?,
            "delete_rows" => loaded.model.delete_rows(sheet, index, count)?,
            "insert_columns" => loaded.model.insert_columns(sheet, index, count)?,
            "delete_columns" => loaded.model.delete_columns(sheet, index, count)?,
            other => return Err(format!("unknown structural op: {other}")),
        }
        loaded.model.evaluate();
        let cells = formula_cell_deltas(loaded);
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

        model.resync_clear_contents();

        let names = model.workbook.get_worksheet_names();
        let mut seeded = 0usize;
        for sheet in &sheets {
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
        model.evaluate();
        loaded.sheets = model.workbook.get_worksheet_names();
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
        match op.as_str() {
            "add" => loaded.model.add_sheet(&name)?,
            "rename" => loaded
                .model
                .rename_sheet(&name, new_name.as_deref().ok_or("rename needs new_name")?)?,
            "delete" => loaded.model.delete_sheet_by_name(&name)?,
            other => return Err(format!("unknown sheet op: {other}")),
        }
        loaded.model.evaluate();
        loaded.sheets = loaded.model.workbook.get_worksheet_names();
        let cells = formula_cell_deltas(loaded);
        Ok(CalcSheetsResult {
            sheets: loaded.sheets.clone(),
            cells,
        })
    })
    .await
}

/// Value of a single cell — used by the shadow cross-check.
#[tauri::command]
pub fn calc_engine_get_cell(
    state: State<CalcEngineState>,
    tab: String,
    sheet: u32,
    row: i32,
    column: i32,
) -> Result<serde_json::Value, String> {
    let guard = lock_recover(&state.inner);
    let loaded = guard.get(&tab).ok_or("no workbook loaded for tab")?;
    loaded
        .model
        .get_cell_value_by_index(sheet, row, column)
        .map(cell_value_to_json)
}

#[tauri::command]
pub fn calc_engine_status(state: State<CalcEngineState>) -> Result<Vec<CalcTabStatus>, String> {
    let guard = lock_recover(&state.inner);
    Ok(guard
        .iter()
        .map(|(tab, loaded)| CalcTabStatus {
            tab: tab.clone(),
            path: loaded.path.clone(),
            sheets: loaded.sheets.clone(),
        })
        .collect())
}

/// Drops the model for a tab (tab closed).
#[tauri::command]
pub fn calc_engine_unload(state: State<CalcEngineState>, tab: String) -> Result<(), String> {
    lock_recover(&state.inner).remove(&tab);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

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
