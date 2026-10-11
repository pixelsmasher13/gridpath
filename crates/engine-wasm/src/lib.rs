//! GridPath engine for Node and the browser.
//!
//! One `Workbook` = one IronCalc `UserModel` built from the bytes of an .xlsx
//! file. Tool handlers mutate the model (IronCalc shifts references on
//! structural edits), `evaluate` recomputes, `snapshot` serialises the state
//! for the TypeScript side, and `patch` writes a batch back into the ORIGINAL
//! bytes through the surgical patcher so untouched package parts stay
//! byte-identical.
//!
//! Rows and columns are 1-based here, as in IronCalc. The TypeScript layer
//! converts from the 0-based snapshot addressing the tools use.

use gridpath_engine_core::{build_cached, cell_value_json, is_missing, snapshot_json, write_cell, Cached, StyleTable};
use ironcalc_base::expressions::types::Area;
use ironcalc_base::types::Color;
use ironcalc_base::{Model, UserModel};
use serde_json::json;
use wasm_bindgen::prelude::*;

fn js_err(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

#[wasm_bindgen]
pub struct Workbook {
    model: UserModel<'static>,
    /// The file's last-saved values for formula cells, captured before the
    /// first evaluation. Some constructs (add-in `_xll.*` calls, external
    /// links, implicit-intersection INDEX/MATCH chains) evaluate to errors
    /// here but were fine in Excel; for those, the saved value is what the
    /// user expects to see, flagged as `file_saved`.
    cached: Cached,
    /// Cells written since the last evaluation (inputs for the incremental
    /// recalc). Automatic evaluation is paused on the UserModel, so writes
    /// only record themselves here until `evaluateEdited` / `evaluate`.
    edited: Vec<(u32, i32, i32)>,
    /// Cells whose style changed since the last evaluation: no recalc
    /// needed, but their snapshot entry must be refreshed.
    restyled: Vec<(u32, i32, i32)>,
    /// A structural edit (rows/columns/sheets/defined names) invalidated the
    /// dependency map: the next evaluation must be a full one.
    needs_full_eval: bool,
    /// Sheet-level properties (names, widths, freeze, hidden…) changed: the
    /// TypeScript side must take a fresh full snapshot.
    needs_full_snapshot: bool,
    /// `openDeferred` and no `evaluatePending` yet: windows blank formulas
    /// the file cached nothing for, instead of showing the placeholder.
    pending_eval: bool,
}

#[wasm_bindgen]
impl Workbook {
    /// Load an .xlsx from bytes and run a full evaluation.
    #[wasm_bindgen]
    pub fn open(bytes: &[u8]) -> Result<Workbook, JsError> {
        let wb = ironcalc::import::load_from_xlsx_bytes(bytes, "workbook", "en", "UTC")
            .map_err(|e| js_err(format!("{e:?}")))?;
        let model = Model::from_workbook(wb, "en").map_err(js_err)?;
        let cached = build_cached(&model);
        let mut model = UserModel::from_model(model);
        model.evaluate();
        // Every UserModel write would otherwise run a FULL recalculation of
        // its own (5+ minutes on an 878k-formula model). Writes are batched
        // by the caller and evaluated once, incrementally where possible.
        model.pause_evaluation();
        Ok(Workbook {
            model,
            cached,
            edited: Vec::new(),
            restyled: Vec::new(),
            needs_full_eval: false,
            needs_full_snapshot: false,
            pending_eval: false,
        })
    }

    /// Load an .xlsx WITHOUT evaluating — the desktop's big-model open.
    /// Reads serve the file's cached values until `evaluatePending` runs
    /// the first full evaluation and reports what moved.
    #[wasm_bindgen(js_name = openDeferred)]
    pub fn open_deferred(bytes: &[u8]) -> Result<Workbook, JsError> {
        let wb = ironcalc::import::load_from_xlsx_bytes(bytes, "workbook", "en", "UTC")
            .map_err(|e| js_err(format!("{e:?}")))?;
        let model = Model::from_workbook(wb, "en").map_err(js_err)?;
        let cached = build_cached(&model);
        let mut model = UserModel::from_model(model);
        model.pause_evaluation();
        Ok(Workbook {
            model,
            cached,
            edited: Vec::new(),
            restyled: Vec::new(),
            needs_full_eval: true,
            needs_full_snapshot: false,
            pending_eval: true,
        })
    }

    /// The deferred first evaluation: full evaluate, then the formula cells
    /// whose live value differs from the file's cache, as
    /// `{"ms":n,"cells":[{"sheet","row","column","value"}]}` (1-based).
    #[wasm_bindgen(js_name = evaluatePending)]
    pub fn evaluate_pending(&mut self) -> Result<String, JsError> {
        if !self.pending_eval {
            return Ok("{\"ms\":0,\"cells\":[]}".to_string());
        }
        self.model.evaluate();
        self.reset_tracking();
        self.pending_eval = false;
        let m = self.model.get_model();
        let changed = gridpath_engine_core::diff::changed_vs_cached(m, &self.cached);
        let cells: Vec<serde_json::Value> = changed
            .into_iter()
            .map(|((sheet, row, column), v)| json!({"sheet": sheet, "row": row, "column": column, "value": v.to_json()}))
            .collect();
        let dropped = gridpath_engine_core::diff::prune_cached(m, &mut self.cached);
        let _ = dropped;
        serde_json::to_string(&json!({"ms": 0, "cells": cells})).map_err(js_err)
    }

    /// Full recalculation of every formula. `evaluateEdited` is the normal
    /// path; this remains for callers that want the unconditional pass.
    pub fn evaluate(&mut self) {
        self.model.evaluate();
        self.reset_tracking();
    }

    /// Recalculates what the writes since the last evaluation affected and
    /// returns JSON describing what the snapshot consumer must refresh:
    /// `{"full": true}` after a structural or sheet-level change (take a new
    /// `snapshot()`), otherwise `{"full": false, "cells": [...], "styles": {...}}`
    /// with every edited, restyled or recomputed cell in snapshot shape plus
    /// a `sheet` index, and the style table entries those cells reference.
    #[wasm_bindgen(js_name = evaluateEdited)]
    pub fn evaluate_edited(&mut self) -> Result<String, JsError> {
        let edited = std::mem::take(&mut self.edited);
        let restyled = std::mem::take(&mut self.restyled);
        let mut changed: Vec<(u32, i32, i32)> = Vec::new();
        if self.needs_full_eval {
            self.model.evaluate();
        } else if !edited.is_empty() {
            changed = self.model.get_model_mut().evaluate_edited(&edited);
        }
        let full = self.needs_full_snapshot;
        self.needs_full_eval = false;
        self.needs_full_snapshot = false;
        if full {
            return Ok(json!({ "full": true }).to_string());
        }
        let mut keys: Vec<(u32, i32, i32)> = Vec::with_capacity(edited.len() + restyled.len() + changed.len());
        keys.extend(edited);
        keys.extend(restyled);
        keys.extend(changed);
        keys.sort_unstable();
        keys.dedup();
        let m = self.model.get_model();
        let sheet_count = m.workbook.worksheets.len() as u32;
        let mut table = StyleTable::new(m);
        let mut out = String::from("{\"full\":false,\"cells\":[");
        let mut first = true;
        for (sheet, r, c) in keys {
            if sheet >= sheet_count || r < 1 || c < 1 {
                continue;
            }
            if !first {
                out.push(',');
            }
            first = false;
            write_cell(&mut out, m, &self.cached, sheet, r, c, true, &mut table).map_err(js_err)?;
        }
        out.push_str("],\"styles\":");
        out.push_str(&serde_json::to_string(&table.styles).map_err(js_err)?);
        out.push('}');
        Ok(out)
    }

    fn reset_tracking(&mut self) {
        self.edited.clear();
        self.restyled.clear();
        self.needs_full_eval = false;
        self.needs_full_snapshot = false;
    }

    /// Rows/columns/sheets/names changed: references shifted and the
    /// dependency map is stale, so the next evaluation is a full one and the
    /// consumer needs a fresh snapshot.
    fn mark_structural(&mut self) {
        self.needs_full_eval = true;
        self.needs_full_snapshot = true;
    }

    /// Sheet-level properties changed (nothing to recalc, but the snapshot's
    /// per-sheet data must be rebuilt).
    fn mark_sheet_props(&mut self) {
        self.needs_full_snapshot = true;
    }

    #[wasm_bindgen(js_name = sheetNames)]
    pub fn sheet_names(&self) -> Vec<String> {
        self.model
            .get_worksheets_properties()
            .into_iter()
            .map(|p| p.name)
            .collect()
    }

    /// JSON snapshot of the whole workbook: sheets (properties, cols, rows,
    /// merges, cells with formula / value / display / style index) plus a
    /// style table keyed by index and the defined names.
    pub fn snapshot(&self) -> Result<String, JsError> {
        snapshot_json(self.model.get_model(), &self.cached).map_err(js_err)
    }

    /// Snapshot with lazy sheets above `lazy_threshold` cells (first visible
    /// sheet eager); see `gridpath_engine_core::SnapshotOptions`. Headless
    /// counterpart of the desktop's `calc_engine_univer_snapshot`.
    #[wasm_bindgen(js_name = snapshotLazy)]
    pub fn snapshot_lazy(&self, lazy_threshold: u32, lazy_sheet_min_cells: u32) -> Result<String, JsError> {
        let m = self.model.get_model();
        let first_visible = m
            .workbook
            .worksheets
            .iter()
            .position(|ws| matches!(ws.state, ironcalc_base::types::SheetState::Visible))
            .unwrap_or(0) as u32;
        let opts = gridpath_engine_core::SnapshotOptions {
            lazy_threshold: Some(lazy_threshold as usize),
            eager_sheets: vec![first_visible],
            lazy_sheet_min_cells: lazy_sheet_min_cells as usize,
            ..Default::default()
        };
        gridpath_engine_core::snapshot_json_with(m, &self.cached, &opts).map_err(js_err)
    }

    /// One sheet's cells and styles (the lazy second half).
    #[wasm_bindgen(js_name = sheetCells)]
    pub fn sheet_cells(&self, sheet: u32) -> Result<String, JsError> {
        gridpath_engine_core::sheet_cells_json(self.model.get_model(), &self.cached, sheet).map_err(js_err)
    }

    /// Structural index of the workbook (`gridpath_engine_core::index`), the
    /// JSON `buildWorkbookIndex` would produce from the snapshot. Not memoized
    /// here: the headless session rebuilds on demand and its workbooks are
    /// small compared with the desktop's.
    #[wasm_bindgen(js_name = workbookIndex)]
    pub fn workbook_index(&self) -> Result<String, JsError> {
        let idx = gridpath_engine_core::index::build_workbook_index(self.model.get_model()).map_err(js_err)?;
        serde_json::to_string(&idx).map_err(js_err)
    }

    /// Viewport grid: workbook-wide case-insensitive text search over formula
    /// text (when present) or value; JSON `[{sheet,row,column,text}]`, 1-based.
    #[wasm_bindgen(js_name = find)]
    pub fn find(&self, query: &str, sheet: Option<u32>, limit: Option<usize>) -> Result<String, JsError> {
        let m = self.model.get_model();
        let needle = query.to_lowercase();
        let cap = limit.unwrap_or(500).min(5_000);
        let mut hits: Vec<serde_json::Value> = Vec::new();
        if needle.is_empty() {
            return Ok("[]".to_string());
        }
        'outer: for (si, ws) in m.workbook.worksheets.iter().enumerate() {
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
                            Ok(ironcalc_base::cell::CellValue::String(s)) => s,
                            Ok(ironcalc_base::cell::CellValue::Number(n)) => n.to_string(),
                            Ok(ironcalc_base::cell::CellValue::Boolean(b)) => if b { "TRUE".to_string() } else { "FALSE".to_string() },
                            _ => String::new(),
                        }
                    };
                    if text.to_lowercase().contains(&needle) {
                        hits.push(json!({"sheet": si, "row": r, "column": c, "text": text.chars().take(120).collect::<String>()}));
                        if hits.len() >= cap {
                            break 'outer;
                        }
                    }
                }
            }
        }
        serde_json::to_string(&hits).map_err(js_err)
    }

    /// Viewport grid: layout for every sheet plus the style table.
    #[wasm_bindgen(js_name = workbookLayout)]
    pub fn workbook_layout(&self) -> Result<String, JsError> {
        gridpath_engine_core::window::workbook_layout_json(self.model.get_model()).map_err(js_err)
    }

    /// Viewport grid: one rectangle of cells (1-based, inclusive).
    #[wasm_bindgen(js_name = sheetWindow)]
    pub fn sheet_window(&self, sheet: u32, r0: i32, r1: i32, c0: i32, c1: i32) -> Result<String, JsError> {
        gridpath_engine_core::window::sheet_window_json(self.model.get_model(), &self.cached, sheet, r0, r1, c0, c1, self.pending_eval).map_err(js_err)
    }

    /// `query_rows` over one sheet (`gridpath_engine_core::query`): `query` is
    /// the tool input as JSON with `sheet` as an index; returns the reply JSON.
    #[wasm_bindgen(js_name = queryRows)]
    pub fn query_rows(&self, query: &str) -> Result<String, JsError> {
        gridpath_engine_core::query::query_rows_json(self.model.get_model(), query).map_err(js_err)
    }

    /// One cell: `{content, value, display}` — content is the formula text
    /// (with `=`) or the literal as typed.
    pub fn cell(&self, sheet: u32, row: i32, col: i32) -> Result<String, JsError> {
        let m = self.model.get_model();
        let content = self.model.get_cell_content(sheet, row, col).map_err(js_err)?;
        let value = cell_value_json(m.get_cell_value_by_index(sheet, row, col).map_err(js_err)?);
        let display = m.get_formatted_cell_value(sheet, row, col).map_err(js_err)?;
        let mut out = json!({"content": content, "value": value, "display": display});
        if is_missing(&value) {
            if let Some(saved) = self.cached.get(&(sheet, row, col)) {
                out["file_saved"] = saved.clone();
            }
        }
        serde_json::to_string(&out).map_err(js_err)
    }

    // ---- writes (no evaluate — the caller batches, then calls evaluateEdited) ----

    #[wasm_bindgen(js_name = setInput)]
    pub fn set_input(&mut self, sheet: u32, row: i32, col: i32, value: &str) -> Result<(), JsError> {
        self.cached.remove(&(sheet, row, col));
        self.model.set_user_input(sheet, row, col, value).map_err(js_err)?;
        self.edited.push((sheet, row, col));
        Ok(())
    }

    #[wasm_bindgen(js_name = clearContents)]
    pub fn clear_contents(&mut self, sheet: u32, row: i32, col: i32, height: i32, width: i32) -> Result<(), JsError> {
        self.model
            .range_clear_contents(&Area { sheet, row, column: col, width, height })
            .map_err(js_err)?;
        for r in row..row + height.max(0) {
            for c in col..col + width.max(0) {
                self.cached.remove(&(sheet, r, c));
                self.edited.push((sheet, r, c));
            }
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = setRangeStyle)]
    pub fn set_range_style(&mut self, sheet: u32, row: i32, col: i32, height: i32, width: i32, path: &str, value: &str) -> Result<(), JsError> {
        self.model
            .update_range_style(&Area { sheet, row, column: col, width, height }, path, value)
            .map_err(js_err)?;
        // Styles never change values; the cells just need their snapshot
        // entries refreshed. A whole-sheet restyle is cheaper as a full one.
        if (height.max(0) as i64) * (width.max(0) as i64) > 50_000 {
            self.needs_full_snapshot = true;
        } else {
            for r in row..row + height.max(0) {
                for c in col..col + width.max(0) {
                    self.restyled.push((sheet, r, c));
                }
            }
        }
        Ok(())
    }

    #[wasm_bindgen(js_name = insertRows)]
    pub fn insert_rows(&mut self, sheet: u32, row: i32, count: i32) -> Result<(), JsError> {
        self.model.insert_rows(sheet, row, count).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }
    #[wasm_bindgen(js_name = deleteRows)]
    pub fn delete_rows(&mut self, sheet: u32, row: i32, count: i32) -> Result<(), JsError> {
        self.model.delete_rows(sheet, row, count).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }
    #[wasm_bindgen(js_name = insertColumns)]
    pub fn insert_columns(&mut self, sheet: u32, col: i32, count: i32) -> Result<(), JsError> {
        self.model.insert_columns(sheet, col, count).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }
    #[wasm_bindgen(js_name = deleteColumns)]
    pub fn delete_columns(&mut self, sheet: u32, col: i32, count: i32) -> Result<(), JsError> {
        self.model.delete_columns(sheet, col, count).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }

    /// Appends a new sheet and returns its index.
    #[wasm_bindgen(js_name = newSheet)]
    pub fn new_sheet(&mut self) -> Result<u32, JsError> {
        self.model.new_sheet().map_err(js_err)?;
        self.mark_structural();
        Ok(self.model.get_worksheets_properties().len() as u32 - 1)
    }
    #[wasm_bindgen(js_name = renameSheet)]
    pub fn rename_sheet(&mut self, sheet: u32, name: &str) -> Result<(), JsError> {
        self.model.rename_sheet(sheet, name).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }
    #[wasm_bindgen(js_name = deleteSheet)]
    pub fn delete_sheet(&mut self, sheet: u32) -> Result<(), JsError> {
        self.model.delete_sheet(sheet).map_err(js_err)?;
        self.mark_structural();
        Ok(())
    }
    #[wasm_bindgen(js_name = setSheetColor)]
    pub fn set_sheet_color(&mut self, sheet: u32, rgb: &str) -> Result<(), JsError> {
        let color = if rgb.is_empty() { Color::None } else { Color::Rgb(rgb.to_string()) };
        self.model.set_sheet_color(sheet, &color).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }

    #[wasm_bindgen(js_name = setFrozen)]
    pub fn set_frozen(&mut self, sheet: u32, rows: i32, cols: i32) -> Result<(), JsError> {
        self.model.set_frozen_rows_count(sheet, rows).map_err(js_err)?;
        self.model.set_frozen_columns_count(sheet, cols).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }
    #[wasm_bindgen(js_name = setColumnsWidth)]
    pub fn set_columns_width(&mut self, sheet: u32, start: i32, end: i32, width: f64) -> Result<(), JsError> {
        self.model.set_columns_width(sheet, start, end, width).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }
    #[wasm_bindgen(js_name = setRowsHeight)]
    pub fn set_rows_height(&mut self, sheet: u32, start: i32, end: i32, height: f64) -> Result<(), JsError> {
        self.model.set_rows_height(sheet, start, end, height).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }
    #[wasm_bindgen(js_name = setColumnsHidden)]
    pub fn set_columns_hidden(&mut self, sheet: u32, start: i32, end: i32, hidden: bool) -> Result<(), JsError> {
        self.model.set_columns_hidden(sheet, start, end, hidden).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }
    #[wasm_bindgen(js_name = setRowsHidden)]
    pub fn set_rows_hidden(&mut self, sheet: u32, start: i32, end: i32, hidden: bool) -> Result<(), JsError> {
        self.model.set_rows_hidden(sheet, start, end, hidden).map_err(js_err)?;
        self.mark_sheet_props();
        Ok(())
    }

    #[wasm_bindgen(js_name = newDefinedName)]
    pub fn new_defined_name(&mut self, name: &str, formula: &str) -> Result<(), JsError> {
        self.model.new_defined_name(name, None, formula).map_err(js_err)?;
        // Re-parse so formulas that already referenced the name resolve.
        self.model.get_model_mut().reset_parsed_structures();
        self.reset_tracking();
        self.needs_full_snapshot = true;
        Ok(())
    }
}

/// Apply a patch (see `xlsx_patch::Patch`) to the ORIGINAL file bytes.
/// Untouched package parts are copied through byte-for-byte.
#[wasm_bindgen]
pub fn patch(base: &[u8], patch_json: &str) -> Result<Vec<u8>, JsError> {
    xlsx_patch::apply_patch_json(base, patch_json).map_err(|e| js_err(format!("{e:?}")))
}
