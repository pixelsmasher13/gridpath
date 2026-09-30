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

use ironcalc_base::cell::CellValue;
use ironcalc_base::expressions::types::Area;
use ironcalc_base::types::{Color, SheetState};
use ironcalc_base::{Model, UserModel};
use serde_json::{json, Map, Value};
use std::collections::BTreeMap;
use wasm_bindgen::prelude::*;

fn js_err(e: impl std::fmt::Display) -> JsError {
    JsError::new(&e.to_string())
}

#[wasm_bindgen]
pub struct Workbook {
    model: UserModel<'static>,
}

#[wasm_bindgen]
impl Workbook {
    /// Load an .xlsx from bytes and run a full evaluation.
    #[wasm_bindgen]
    pub fn open(bytes: &[u8]) -> Result<Workbook, JsError> {
        let wb = ironcalc::import::load_from_xlsx_bytes(bytes, "workbook", "en", "UTC")
            .map_err(|e| js_err(format!("{e:?}")))?;
        let model = Model::from_workbook(wb, "en").map_err(js_err)?;
        let mut model = UserModel::from_model(model);
        model.evaluate();
        Ok(Workbook { model })
    }

    pub fn evaluate(&mut self) {
        self.model.evaluate();
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
        let m = self.model.get_model();
        let mut styles: BTreeMap<i32, Value> = BTreeMap::new();
        let mut sheets = Vec::new();
        for (idx, ws) in m.workbook.worksheets.iter().enumerate() {
            let sheet = idx as u32;
            let mut cells = Vec::new();
            let mut rows: Vec<&i32> = ws.sheet_data.keys().collect();
            rows.sort();
            for r in rows {
                let row = &ws.sheet_data[r];
                let mut cols: Vec<&i32> = row.keys().collect();
                cols.sort();
                for c in cols {
                    let mut cell = Map::new();
                    cell.insert("r".into(), json!(r));
                    cell.insert("c".into(), json!(c));
                    if let Ok(Some(f)) = m.get_cell_formula(sheet, *r, *c) {
                        // IronCalc returns the formula with its leading '='; keep exactly one.
                        let f = f.strip_prefix('=').unwrap_or(&f).to_string();
                        cell.insert("f".into(), json!(format!("={f}")));
                    }
                    let v = match m.get_cell_value_by_index(sheet, *r, *c).map_err(js_err)? {
                        CellValue::None => Value::Null,
                        CellValue::Number(n) => json!(n),
                        CellValue::String(s) => json!(s),
                        CellValue::Boolean(b) => json!(b),
                    };
                    cell.insert("v".into(), v);
                    if let Ok(d) = m.get_formatted_cell_value(sheet, *r, *c) {
                        cell.insert("d".into(), json!(d));
                    }
                    let s = m.get_cell_style_index(sheet, *r, *c).unwrap_or(0);
                    if s != 0 {
                        cell.insert("s".into(), json!(s));
                        if !styles.contains_key(&s) {
                            if let Ok(style) = m.get_style_for_cell(sheet, *r, *c) {
                                styles.insert(s, serde_json::to_value(style).map_err(js_err)?);
                            }
                        }
                    }
                    cells.push(Value::Object(cell));
                }
            }
            sheets.push(json!({
                "name": ws.name,
                "hidden": !matches!(ws.state, SheetState::Visible),
                "color": color_json(&ws.color),
                "frozenRows": ws.frozen_rows,
                "frozenColumns": ws.frozen_columns,
                "showGridLines": ws.show_grid_lines,
                "merges": ws.merge_cells,
                "cols": ws.cols.iter().map(|c| json!({"min": c.min, "max": c.max, "width": c.width, "hidden": c.hidden})).collect::<Vec<_>>(),
                "rows": ws.rows.iter().map(|r| json!({"r": r.r, "height": r.height, "hidden": r.hidden})).collect::<Vec<_>>(),
                "comments": ws.comments.iter().map(|c| json!({"ref": c.cell_ref, "text": c.text, "author": c.author_name})).collect::<Vec<_>>(),
                "cells": cells,
            }));
        }
        let defined_names: Vec<Value> = m
            .get_defined_name_list()
            .into_iter()
            .map(|(name, scope, formula)| json!({"name": name, "formula": formula, "sheetId": scope}))
            .collect();
        let out = json!({
            "sheets": sheets,
            "styles": styles,
            "definedNames": defined_names,
        });
        serde_json::to_string(&out).map_err(js_err)
    }

    /// One cell: `{content, value, display}` — content is the formula text
    /// (with `=`) or the literal as typed.
    pub fn cell(&self, sheet: u32, row: i32, col: i32) -> Result<String, JsError> {
        let m = self.model.get_model();
        let content = self.model.get_cell_content(sheet, row, col).map_err(js_err)?;
        let value = match m.get_cell_value_by_index(sheet, row, col).map_err(js_err)? {
            CellValue::None => Value::Null,
            CellValue::Number(n) => json!(n),
            CellValue::String(s) => json!(s),
            CellValue::Boolean(b) => json!(b),
        };
        let display = m.get_formatted_cell_value(sheet, row, col).map_err(js_err)?;
        serde_json::to_string(&json!({"content": content, "value": value, "display": display})).map_err(js_err)
    }

    // ---- writes (no evaluate — the caller batches, then calls evaluate) ----

    #[wasm_bindgen(js_name = setInput)]
    pub fn set_input(&mut self, sheet: u32, row: i32, col: i32, value: &str) -> Result<(), JsError> {
        self.model.set_user_input(sheet, row, col, value).map_err(js_err)
    }

    #[wasm_bindgen(js_name = clearContents)]
    pub fn clear_contents(&mut self, sheet: u32, row: i32, col: i32, height: i32, width: i32) -> Result<(), JsError> {
        self.model
            .range_clear_contents(&Area { sheet, row, column: col, width, height })
            .map_err(js_err)
    }

    #[wasm_bindgen(js_name = setRangeStyle)]
    pub fn set_range_style(&mut self, sheet: u32, row: i32, col: i32, height: i32, width: i32, path: &str, value: &str) -> Result<(), JsError> {
        self.model
            .update_range_style(&Area { sheet, row, column: col, width, height }, path, value)
            .map_err(js_err)
    }

    #[wasm_bindgen(js_name = insertRows)]
    pub fn insert_rows(&mut self, sheet: u32, row: i32, count: i32) -> Result<(), JsError> {
        self.model.insert_rows(sheet, row, count).map_err(js_err)
    }
    #[wasm_bindgen(js_name = deleteRows)]
    pub fn delete_rows(&mut self, sheet: u32, row: i32, count: i32) -> Result<(), JsError> {
        self.model.delete_rows(sheet, row, count).map_err(js_err)
    }
    #[wasm_bindgen(js_name = insertColumns)]
    pub fn insert_columns(&mut self, sheet: u32, col: i32, count: i32) -> Result<(), JsError> {
        self.model.insert_columns(sheet, col, count).map_err(js_err)
    }
    #[wasm_bindgen(js_name = deleteColumns)]
    pub fn delete_columns(&mut self, sheet: u32, col: i32, count: i32) -> Result<(), JsError> {
        self.model.delete_columns(sheet, col, count).map_err(js_err)
    }

    /// Appends a new sheet and returns its index.
    #[wasm_bindgen(js_name = newSheet)]
    pub fn new_sheet(&mut self) -> Result<u32, JsError> {
        self.model.new_sheet().map_err(js_err)?;
        Ok(self.model.get_worksheets_properties().len() as u32 - 1)
    }
    #[wasm_bindgen(js_name = renameSheet)]
    pub fn rename_sheet(&mut self, sheet: u32, name: &str) -> Result<(), JsError> {
        self.model.rename_sheet(sheet, name).map_err(js_err)
    }
    #[wasm_bindgen(js_name = deleteSheet)]
    pub fn delete_sheet(&mut self, sheet: u32) -> Result<(), JsError> {
        self.model.delete_sheet(sheet).map_err(js_err)
    }
    #[wasm_bindgen(js_name = setSheetColor)]
    pub fn set_sheet_color(&mut self, sheet: u32, rgb: &str) -> Result<(), JsError> {
        let color = if rgb.is_empty() { Color::None } else { Color::Rgb(rgb.to_string()) };
        self.model.set_sheet_color(sheet, &color).map_err(js_err)
    }

    #[wasm_bindgen(js_name = setFrozen)]
    pub fn set_frozen(&mut self, sheet: u32, rows: i32, cols: i32) -> Result<(), JsError> {
        self.model.set_frozen_rows_count(sheet, rows).map_err(js_err)?;
        self.model.set_frozen_columns_count(sheet, cols).map_err(js_err)
    }
    #[wasm_bindgen(js_name = setColumnsWidth)]
    pub fn set_columns_width(&mut self, sheet: u32, start: i32, end: i32, width: f64) -> Result<(), JsError> {
        self.model.set_columns_width(sheet, start, end, width).map_err(js_err)
    }
    #[wasm_bindgen(js_name = setRowsHeight)]
    pub fn set_rows_height(&mut self, sheet: u32, start: i32, end: i32, height: f64) -> Result<(), JsError> {
        self.model.set_rows_height(sheet, start, end, height).map_err(js_err)
    }
    #[wasm_bindgen(js_name = setColumnsHidden)]
    pub fn set_columns_hidden(&mut self, sheet: u32, start: i32, end: i32, hidden: bool) -> Result<(), JsError> {
        self.model.set_columns_hidden(sheet, start, end, hidden).map_err(js_err)
    }
    #[wasm_bindgen(js_name = setRowsHidden)]
    pub fn set_rows_hidden(&mut self, sheet: u32, start: i32, end: i32, hidden: bool) -> Result<(), JsError> {
        self.model.set_rows_hidden(sheet, start, end, hidden).map_err(js_err)
    }

    #[wasm_bindgen(js_name = newDefinedName)]
    pub fn new_defined_name(&mut self, name: &str, formula: &str) -> Result<(), JsError> {
        self.model.new_defined_name(name, None, formula).map_err(js_err)
    }
}

fn color_json(c: &Color) -> Value {
    match c {
        Color::Rgb(s) => json!(s),
        _ => Value::Null,
    }
}

/// Apply a patch (see `xlsx_patch::Patch`) to the ORIGINAL file bytes.
/// Untouched package parts are copied through byte-for-byte.
#[wasm_bindgen]
pub fn patch(base: &[u8], patch_json: &str) -> Result<Vec<u8>, JsError> {
    xlsx_patch::apply_patch_json(base, patch_json).map_err(|e| js_err(format!("{e:?}")))
}
