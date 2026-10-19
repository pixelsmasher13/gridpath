//! Shared engine plumbing over the vendored IronCalc model.
//!
//! The one thing both GridPath front ends need from the engine is the same
//! JSON snapshot of a workbook — sheets (properties, columns, rows, merges,
//! comments, cells with formula / value / display / style index), a style
//! table keyed by index, and the defined names. The wasm engine hands it to
//! the headless session; the desktop app hands it straight to the Univer
//! grid instead of parsing the file a second time with ExcelJS.
//!
//! Rows and columns are 1-based (IronCalc). Colours are resolved to
//! `#RRGGBB` against the workbook theme so consumers never see theme slots.
//! The serialiser streams into one `String`: building a serde_json tree for
//! 2.3M cells took 147 s inside the wasm heap, the stream takes 5.7 s.

pub mod diff;
pub mod index;
pub mod query;
pub mod window;

use ironcalc_base::cell::CellValue;
use ironcalc_base::types::{Color, SheetState, Style, Theme};
use ironcalc_base::Model;
use serde_json::{json, Value};
use std::collections::{BTreeMap, HashMap};
use std::fmt::Write;

/// The file's last-saved values for formula cells, captured before the first
/// evaluation. Some constructs (add-in `_xll.*` calls, external links,
/// implicit-intersection INDEX/MATCH chains) evaluate to errors here but
/// were fine in Excel; for those, the saved value is what the user expects
/// to see, flagged as `fileSaved`.
pub type Cached = HashMap<(u32, i32, i32), Value>;

/// Capture every formula cell's imported (file-cached) value. Call on a
/// freshly loaded model, BEFORE `evaluate()`.
pub fn build_cached(model: &Model) -> Cached {
    // Walk the sheet data directly: `get_cell_formula` would stringify every
    // formula's AST just to answer "is this a formula" — on a 900k-formula
    // model that was 1.4 s of the deferred first evaluate.
    let mut cached: Cached = HashMap::new();
    for (sheet_index, ws) in model.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_none() {
                    continue;
                }
                let Ok(v) = model.get_cell_value_by_index(sheet_index as u32, *row, *column) else { continue };
                let j = cell_value_json(v);
                if !is_missing(&j) {
                    cached.insert((sheet_index as u32, *row, *column), j);
                }
            }
        }
    }
    cached
}

pub fn cell_value_json(v: CellValue) -> Value {
    match v {
        CellValue::None => Value::Null,
        CellValue::Number(n) => json!(n),
        CellValue::String(s) => json!(s),
        CellValue::Boolean(b) => json!(b),
    }
}

/// Null, empty, or an Excel error value — a live result that isn't usable.
pub fn is_missing(v: &Value) -> bool {
    match v {
        Value::Null => true,
        Value::String(s) => s.is_empty() || s.starts_with('#'),
        _ => false,
    }
}

pub fn value_missing(v: &CellValue) -> bool {
    match v {
        CellValue::None => true,
        CellValue::String(s) => s.is_empty() || s.starts_with('#'),
        _ => false,
    }
}

/// Appends `s` as a JSON string literal (RFC 8259 escaping) without an
/// intermediate allocation.
pub fn push_json_str(out: &mut String, s: &str) {
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => {
                let _ = write!(out, "\\u{:04x}", c as u32);
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

/// Appends a cell value as JSON: numbers as JSON numbers (non-finite → null),
/// strings (including Excel errors like "#N/A") as strings.
pub fn push_cell_value(out: &mut String, v: &CellValue) {
    match v {
        CellValue::None => out.push_str("null"),
        CellValue::Number(n) if n.is_finite() => {
            let _ = write!(out, "{n}");
        }
        CellValue::Number(_) => out.push_str("null"),
        CellValue::String(s) => push_json_str(out, s),
        CellValue::Boolean(b) => out.push_str(if *b { "true" } else { "false" }),
    }
}

pub fn color_json(c: &Color, theme: &Theme) -> Value {
    match c {
        Color::None => Value::Null,
        other => json!(other.to_rgb(theme)),
    }
}

/// A style as the snapshot ships it: IronCalc's `Style` serialisation with
/// every colour (fill, font, borders) resolved to `#RRGGBB`.
pub fn style_json(style: &Style, theme: &Theme) -> Result<Value, String> {
    let mut v = serde_json::to_value(style).map_err(|e| e.to_string())?;
    if let Some(fill) = v.get_mut("fill") {
        fill["color"] = color_json(&style.fill.color, theme);
    }
    if let Some(font) = v.get_mut("font") {
        font["color"] = color_json(&style.font.color, theme);
    }
    if let Some(border) = v.get_mut("border").and_then(|b| b.as_object_mut()) {
        let sides = [
            ("left", style.border.left.as_ref()),
            ("right", style.border.right.as_ref()),
            ("top", style.border.top.as_ref()),
            ("bottom", style.border.bottom.as_ref()),
            ("diagonal", style.border.diagonal.as_ref()),
        ];
        for (name, item) in sides {
            if let (Some(side), Some(item)) = (border.get_mut(name), item) {
                side["color"] = color_json(&item.color, theme);
            }
        }
    }
    Ok(v)
}

/// Per-snapshot memo: the style table being emitted plus each style index's
/// number format, so display text is produced without constructing a
/// `Style` per cell.
pub struct StyleTable {
    pub styles: BTreeMap<i32, Value>,
    num_fmts: HashMap<i32, String>,
    theme: Theme,
    /// Emit formula cells that have not been evaluated yet as blank (`v`
    /// null) instead of IronCalc's `#ERROR!` placeholder. Set for snapshots
    /// taken before a deferred first evaluate: a formula the file carries no
    /// cached value for is not an error, it is simply not computed yet.
    pub blank_unevaluated: bool,
}

impl StyleTable {
    pub fn new(model: &Model) -> Self {
        StyleTable { styles: BTreeMap::new(), num_fmts: HashMap::new(), theme: model.get_theme(), blank_unevaluated: false }
    }

    pub fn with_blank_unevaluated(mut self, on: bool) -> Self {
        self.blank_unevaluated = on;
        self
    }

    pub fn theme(&self) -> &Theme {
        &self.theme
    }

    fn num_fmt(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) -> &str {
        if !self.num_fmts.contains_key(&s) {
            let fmt = m
                .get_style_for_cell(sheet, r, c)
                .map(|st| st.num_fmt)
                .unwrap_or_else(|_| "General".to_string());
            self.num_fmts.insert(s, fmt);
        }
        &self.num_fmts[&s]
    }

    fn record(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) -> Result<(), String> {
        if !self.styles.contains_key(&s) {
            if let Ok(style) = m.get_style_for_cell(sheet, r, c) {
                let v = style_json(&style, &self.theme)?;
                self.styles.insert(s, v);
            }
        }
        Ok(())
    }
}

/// Appends one cell in snapshot shape: `r`, `c` (1-based), optional `f`
/// (formula with exactly one leading '='), `v`, `d` (display text, numbers
/// and booleans only — strings and errors display as themselves), `s`
/// (style index; the style itself is added to the table), and `sheet` when
/// `sheet_tag` is set. A formula whose live result is unusable but which has
/// a file-saved value gets that value as `v`, with `live` and
/// `fileSaved: true` alongside.
#[allow(clippy::too_many_arguments)]
pub fn write_cell(
    out: &mut String,
    m: &Model,
    cached: &Cached,
    sheet: u32,
    r: i32,
    c: i32,
    sheet_tag: bool,
    table: &mut StyleTable,
) -> Result<(), String> {
    let _ = write!(out, "{{\"r\":{r},\"c\":{c}");
    if let Ok(Some(f)) = m.get_cell_formula(sheet, r, c) {
        let f = f.strip_prefix('=').unwrap_or(&f);
        out.push_str(",\"f\":");
        let mut with_eq = String::with_capacity(f.len() + 1);
        with_eq.push('=');
        with_eq.push_str(f);
        push_json_str(out, &with_eq);
    }
    let mut value = m.get_cell_value_by_index(sheet, r, c)?;
    if table.blank_unevaluated && is_unevaluated(m, sheet, r, c) {
        value = CellValue::None;
    }
    let mut wrote_v = false;
    if value_missing(&value) {
        if let Some(saved) = cached.get(&(sheet, r, c)) {
            out.push_str(",\"live\":");
            push_cell_value(out, &value);
            out.push_str(",\"fileSaved\":true,\"v\":");
            out.push_str(&serde_json::to_string(saved).map_err(|e| e.to_string())?);
            wrote_v = true;
        }
    }
    if !wrote_v {
        out.push_str(",\"v\":");
        push_cell_value(out, &value);
    }
    let s = m.get_cell_style_index(sheet, r, c).unwrap_or(0);
    if matches!(value, CellValue::Number(_) | CellValue::Boolean(_)) {
        let fmt = table.num_fmt(m, sheet, r, c, s).to_string();
        if let Ok(d) = m.get_formatted_cell_value_using(sheet, r, c, &fmt) {
            out.push_str(",\"d\":");
            push_json_str(out, &d);
        }
    }
    if s != 0 {
        let _ = write!(out, ",\"s\":{s}");
        table.record(m, sheet, r, c, s)?;
    }
    if sheet_tag {
        let _ = write!(out, ",\"sheet\":{sheet}");
    }
    // Conditional formatting that matched this cell in the last evaluation:
    // the fill and font the rule applies (only what differs from the cell's
    // own style) and a data bar if the rule draws one.
    if m.cell_has_conditional_format(sheet, r, c) {
        if let (Ok(ext), Ok(base)) = (m.get_extended_style_for_cell(sheet, r, c), m.get_style_for_cell(sheet, r, c)) {
            let theme = m.get_theme();
            let mut cf = serde_json::Map::new();
            if ext.style.fill.color != base.fill.color {
                cf.insert("fill".into(), color_json(&ext.style.fill.color, &theme));
            }
            if ext.style.font.color != base.font.color {
                cf.insert("color".into(), color_json(&ext.style.font.color, &theme));
            }
            if ext.style.font.b != base.font.b { cf.insert("bold".into(), Value::Bool(ext.style.font.b)); }
            if ext.style.font.i != base.font.i { cf.insert("italic".into(), Value::Bool(ext.style.font.i)); }
            if ext.style.font.u != base.font.u { cf.insert("underline".into(), Value::Bool(ext.style.font.u)); }
            if ext.style.font.strike != base.font.strike { cf.insert("strike".into(), Value::Bool(ext.style.font.strike)); }
            if let Some(db) = &ext.data_bar {
                cf.insert("dataBar".into(), serde_json::json!({
                    "value": db.value,
                    "axis": db.axis_position,
                    "positive": color_json(&db.positive_color, &theme),
                    "negative": color_json(&db.negative_color, &theme),
                    "gradient": db.is_gradient,
                    "showValue": db.show_value,
                }));
            }
            if !cf.is_empty() {
                out.push_str(",\"cf\":");
                out.push_str(&Value::Object(cf).to_string());
            }
        }
    }
    out.push('}');
    Ok(())
}

/// A formula cell whose value has never been computed (fresh model, or a
/// file that carried no cached value for it, before the first evaluate).
pub fn is_unevaluated(m: &Model, sheet: u32, r: i32, c: i32) -> bool {
    use ironcalc_base::types::{Cell, FormulaValue};
    matches!(
        m.workbook.worksheets.get(sheet as usize).and_then(|ws| ws.cell(r, c)),
        Some(Cell::CellFormula { v: FormulaValue::Unevaluated, .. } | Cell::ArrayFormula { v: FormulaValue::Unevaluated, .. })
    )
}

/// How much of each sheet a snapshot carries.
#[derive(Debug, Clone)]
pub struct SnapshotOptions {
    /// When the workbook holds more cells than this, sheets outside
    /// `eager_sheets` are emitted LAZY: header, extent and a skeleton of
    /// cells (the first `skeleton_rows` rows and `skeleton_cols` columns —
    /// enough for the grid's tab strip, the agent's structural index and a
    /// preview), flagged `"lazy": true`. Their cells come later through
    /// `sheet_cells_json`. `None` disables laziness.
    pub lazy_threshold: Option<usize>,
    /// Sheets always emitted in full (the one the user sees first).
    pub eager_sheets: Vec<u32>,
    /// In lazy mode, a sheet with this many cells or fewer is still sent in
    /// full: switching to it must stay instant, and small sheets are not
    /// where the memory goes.
    pub lazy_sheet_min_cells: usize,
    pub skeleton_rows: i32,
    pub skeleton_cols: i32,
    /// See `StyleTable::blank_unevaluated`.
    pub blank_unevaluated: bool,
}

impl Default for SnapshotOptions {
    fn default() -> Self {
        SnapshotOptions { lazy_threshold: None, eager_sheets: Vec::new(), lazy_sheet_min_cells: 100_000, skeleton_rows: 60, skeleton_cols: 12, blank_unevaluated: false }
    }
}

fn sheet_cell_count(ws: &ironcalc_base::types::Worksheet) -> usize {
    ws.sheet_data.values().map(|r| r.len()).sum()
}

/// JSON snapshot of the whole workbook: sheets (properties, cols, rows,
/// merges, comments, cells), the style table keyed by index, and the
/// defined names. Every sheet in full.
pub fn snapshot_json(m: &Model, cached: &Cached) -> Result<String, String> {
    snapshot_json_with(m, cached, &SnapshotOptions::default())
}

fn sheet_header(ws: &ironcalc_base::types::Worksheet, theme: &Theme) -> Value {
    let (mut cell_count, mut max_row, mut max_col) = (0usize, 0i32, 0i32);
    for (r, row) in &ws.sheet_data {
        cell_count += row.len();
        if *r > max_row {
            max_row = *r;
        }
        for c in row.keys() {
            if *c > max_col {
                max_col = *c;
            }
        }
    }
    json!({
        "name": ws.name,
        "hidden": !matches!(ws.state, SheetState::Visible),
        "color": color_json(&ws.color, theme),
        "frozenRows": ws.frozen_rows,
        "frozenColumns": ws.frozen_columns,
        "showGridLines": ws.show_grid_lines,
        "merges": ws.merge_cells,
        "cols": ws.cols.iter().map(|c| json!({"min": c.min, "max": c.max, "width": c.width, "hidden": c.hidden})).collect::<Vec<_>>(),
        "rows": ws.rows.iter().map(|r| json!({"r": r.r, "height": r.height, "hidden": r.hidden})).collect::<Vec<_>>(),
        "comments": ws.comments.iter().map(|c| json!({"ref": c.cell_ref, "text": c.text, "author": c.author_name})).collect::<Vec<_>>(),
        "cellCount": cell_count,
        "maxRow": max_row,
        "maxCol": max_col,
    })
}

/// Appends a sheet's cells (all of them, or only the skeleton) as a JSON
/// array body (no brackets).
fn write_sheet_cells(
    out: &mut String,
    m: &Model,
    cached: &Cached,
    sheet: u32,
    skeleton: Option<(i32, i32)>,
    table: &mut StyleTable,
) -> Result<(), String> {
    let ws = &m.workbook.worksheets[sheet as usize];
    let mut rows: Vec<&i32> = ws.sheet_data.keys().collect();
    rows.sort();
    let mut first = true;
    for r in rows {
        let row = &ws.sheet_data[r];
        let mut cols: Vec<&i32> = row.keys().collect();
        cols.sort();
        for c in cols {
            if let Some((skel_rows, skel_cols)) = skeleton {
                if *r > skel_rows && *c > skel_cols {
                    continue;
                }
            }
            if !first {
                out.push(',');
            }
            first = false;
            write_cell(out, m, cached, sheet, *r, *c, false, table)?;
        }
    }
    Ok(())
}

pub fn snapshot_json_with(m: &Model, cached: &Cached, opts: &SnapshotOptions) -> Result<String, String> {
    let mut table = StyleTable::new(m).with_blank_unevaluated(opts.blank_unevaluated);
    let total_cells: usize = m.workbook.worksheets.iter().map(sheet_cell_count).sum();
    let lazy_mode = opts.lazy_threshold.is_some_and(|t| total_cells > t);
    // Size the buffer up front: a 142 MB snapshot grown by doubling peaked at
    // 284 MB of live heap on top of the model. The measured average is ~60
    // bytes per cell on large models; 72 leaves headroom so the last cells
    // never trigger a final doubling (60 did, by a few kilobytes).
    let mut out = String::with_capacity((total_cells.saturating_mul(72)).max(1 << 20));
    out.push_str("{\"sheets\":[");
    for (idx, ws) in m.workbook.worksheets.iter().enumerate() {
        let sheet = idx as u32;
        if idx > 0 {
            out.push(',');
        }
        let lazy = lazy_mode && !opts.eager_sheets.contains(&sheet) && sheet_cell_count(ws) > opts.lazy_sheet_min_cells;
        let mut header = sheet_header(ws, table.theme());
        if lazy {
            header["lazy"] = json!(true);
        }
        let header = serde_json::to_string(&header).map_err(|e| e.to_string())?;
        // Reopen the header object and append the cell array.
        out.push_str(&header[..header.len() - 1]);
        out.push_str(",\"cells\":[");
        let skeleton = if lazy { Some((opts.skeleton_rows, opts.skeleton_cols)) } else { None };
        write_sheet_cells(&mut out, m, cached, sheet, skeleton, &mut table)?;
        out.push_str("]}");
    }
    let defined_names: Vec<Value> = m
        .get_defined_name_list()
        .into_iter()
        .map(|(name, scope, formula)| json!({"name": name, "formula": formula, "sheetId": scope}))
        .collect();
    out.push_str("],\"styles\":");
    out.push_str(&serde_json::to_string(&table.styles).map_err(|e| e.to_string())?);
    out.push_str(",\"definedNames\":");
    out.push_str(&serde_json::to_string(&defined_names).map_err(|e| e.to_string())?);
    out.push('}');
    Ok(out)
}

/// One sheet's cells in full, with the styles they reference:
/// `{"sheet": i, "cells": [...], "styles": {...}}`. The second half of a lazy
/// snapshot, fetched when the sheet is first shown or written to.
pub fn sheet_cells_json(m: &Model, cached: &Cached, sheet: u32) -> Result<String, String> {
    sheet_cells_json_with(m, cached, sheet, false)
}

/// `sheet_cells_json` with the unevaluated-as-blank rule (see `StyleTable`).
pub fn sheet_cells_json_with(m: &Model, cached: &Cached, sheet: u32, blank_unevaluated: bool) -> Result<String, String> {
    if sheet as usize >= m.workbook.worksheets.len() {
        return Err(format!("sheet index {sheet} out of range"));
    }
    let mut table = StyleTable::new(m).with_blank_unevaluated(blank_unevaluated);
    let cells = sheet_cell_count(&m.workbook.worksheets[sheet as usize]);
    let mut out = String::with_capacity((cells.saturating_mul(72)).max(1 << 16));
    let _ = write!(out, "{{\"sheet\":{sheet},\"cells\":[");
    write_sheet_cells(&mut out, m, cached, sheet, None, &mut table)?;
    out.push_str("],\"styles\":");
    out.push_str(&serde_json::to_string(&table.styles).map_err(|e| e.to_string())?);
    out.push('}');
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model() -> Model<'static> {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.add_sheet("Big").unwrap();
        for r in 1..=100 {
            for c in 1..=20 {
                m.set_user_input(1, r, c, format!("{}", r * c)).unwrap();
            }
        }
        m.set_user_input(0, 1, 1, "=Big!A1*2".to_string()).unwrap();
        m.evaluate();
        m
    }

    #[test]
    fn lazy_snapshot_keeps_skeleton_and_flags_sheet() {
        let m = model();
        let full: Value = serde_json::from_str(&snapshot_json(&m, &Cached::new()).unwrap()).unwrap();
        assert_eq!(full["sheets"][1]["cells"].as_array().unwrap().len(), 2000);
        assert!(full["sheets"][1].get("lazy").is_none());
        assert_eq!(full["sheets"][1]["cellCount"], 2000);
        assert_eq!(full["sheets"][1]["maxRow"], 100);

        let opts = SnapshotOptions { lazy_threshold: Some(1000), eager_sheets: vec![0], lazy_sheet_min_cells: 500, skeleton_rows: 10, skeleton_cols: 2, blank_unevaluated: false };
        let lazy: Value = serde_json::from_str(&snapshot_json_with(&m, &Cached::new(), &opts).unwrap()).unwrap();
        assert_eq!(lazy["sheets"][1]["lazy"], true);
        assert!(lazy["sheets"][0].get("lazy").is_none());
        // Skeleton: rows 1..=10 (all 20 columns) + columns 1..=2 for the other 90 rows.
        assert_eq!(lazy["sheets"][1]["cells"].as_array().unwrap().len(), 10 * 20 + 90 * 2);
        assert_eq!(lazy["sheets"][1]["cellCount"], 2000);

        let cells: Value = serde_json::from_str(&sheet_cells_json(&m, &Cached::new(), 1).unwrap()).unwrap();
        assert_eq!(cells["sheet"], 1);
        assert_eq!(cells["cells"].as_array().unwrap().len(), 2000);

        // A sheet at or under the per-sheet floor stays eager even in lazy mode.
        let small = SnapshotOptions { lazy_sheet_min_cells: 2000, ..opts.clone() };
        let eager: Value = serde_json::from_str(&snapshot_json_with(&m, &Cached::new(), &small).unwrap()).unwrap();
        assert!(eager["sheets"][1].get("lazy").is_none());
        assert_eq!(eager["sheets"][1]["cells"].as_array().unwrap().len(), 2000);

        // Below the workbook threshold nothing is lazy.
        let opts = SnapshotOptions { lazy_threshold: Some(10_000), ..opts };
        let eager: Value = serde_json::from_str(&snapshot_json_with(&m, &Cached::new(), &opts).unwrap()).unwrap();
        assert!(eager["sheets"][1].get("lazy").is_none());
    }
}

#[cfg(test)]
mod unevaluated_probe {
    use super::*;
    /// Before the first evaluate a formula without a cached value is
    /// `#ERROR!` to IronCalc; the deferred-open snapshot must show it blank.
    #[test]
    fn deferred_snapshot_blanks_unevaluated_formulas() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.set_user_input(0, 1, 1, "=1+1".to_string()).unwrap();
        let cached = Cached::new();
        let plain: Value = serde_json::from_str(&snapshot_json(&m, &cached).unwrap()).unwrap();
        assert_eq!(plain["sheets"][0]["cells"][0]["v"], "#ERROR!");
        let opts = SnapshotOptions { blank_unevaluated: true, ..Default::default() };
        let blank: Value = serde_json::from_str(&snapshot_json_with(&m, &cached, &opts).unwrap()).unwrap();
        assert_eq!(blank["sheets"][0]["cells"][0]["v"], Value::Null);
        assert_eq!(blank["sheets"][0]["cells"][0]["f"], "=1+1");
        let cells: Value = serde_json::from_str(&sheet_cells_json_with(&m, &cached, 0, true).unwrap()).unwrap();
        assert_eq!(cells["cells"][0]["v"], Value::Null);
        m.evaluate();
        let after: Value = serde_json::from_str(&snapshot_json_with(&m, &cached, &opts).unwrap()).unwrap();
        assert_eq!(after["sheets"][0]["cells"][0]["v"], 2.0);
    }
}

#[cfg(test)]
mod new_sheet_probe {
    use ironcalc_base::Model;
    #[test]
    fn edits_on_a_sheet_added_after_the_full_evaluate_are_computed_incrementally() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.set_user_input(0, 1, 1, "5".to_string()).unwrap();
        m.evaluate();
        m.add_sheet("New").unwrap();
        m.set_user_input(1, 10, 1, "=1+1".to_string()).unwrap();
        m.set_user_input(1, 11, 1, "=Sheet1!A1*3".to_string()).unwrap();
        let changed = m.evaluate_edited(&[(1, 10, 1), (1, 11, 1)]);
        let v10 = m.get_cell_value_by_index(1, 10, 1).unwrap();
        let v11 = m.get_cell_value_by_index(1, 11, 1).unwrap();
        eprintln!("changed={changed:?} A10={v10:?} A11={v11:?}");
        assert_eq!(format!("{v10:?}"), "Number(2.0)");
        assert_eq!(format!("{v11:?}"), "Number(15.0)");
    }
}
