//! Viewport reads for the big-model grid: the renderer holds only what is
//! on screen and asks the engine for rectangles of cells as the user
//! scrolls. Two serialisers:
//!
//! * `workbook_layout_json` — everything the renderer needs BEFORE any cell:
//!   per sheet the name, hidden flag, tab colour, frozen panes, merges,
//!   custom column widths / row heights, the used extent, plus the workbook
//!   style table resolved to `#RRGGBB`. Small (tens of KB) even for a
//!   2.3M-cell model.
//! * `sheet_window_json` — the cells of one rectangle, in the same per-cell
//!   shape as the full snapshot (`r`, `c`, `v`, `d`, `f`, `s`, `fileSaved`),
//!   with the style table entries those cells reference. Rows and columns
//!   are 1-based and inclusive.

use crate::{style_json, write_cell, Cached, StyleTable};
use ironcalc_base::types::SheetState;
use ironcalc_base::Model;
use serde_json::{json, Value};
use std::fmt::Write;

/// Cells per window the caller may ask for; a 64×32 tile is 2,048.
pub const MAX_WINDOW_CELLS: i64 = 65_536;

pub fn workbook_layout_json(m: &Model) -> Result<String, String> {
    let theme = m.get_theme();
    let mut sheets = Vec::with_capacity(m.workbook.worksheets.len());
    let selected_sheet = m.workbook.views.get(&0).map(|v| v.sheet as usize);
    for (sheet_index, ws) in m.workbook.worksheets.iter().enumerate() {
        // `maxRow`/`maxCol`: any cell, including styled blanks (grid extent).
        // `usedRow`/`usedCol`: cells with content, formula or value (what
        // the agent's used range and Ctrl+End mean).
        let (mut cell_count, mut max_row, mut max_col, mut used_row, mut used_col) = (0usize, 0i32, 0i32, 0i32, 0i32);
        for (r, row) in &ws.sheet_data {
            cell_count += row.len();
            if *r > max_row {
                max_row = *r;
            }
            for (c, cell) in row {
                if *c > max_col {
                    max_col = *c;
                }
                let has_content = cell.get_formula().is_some()
                    || !matches!(cell, ironcalc_base::types::Cell::EmptyCell { .. });
                if has_content {
                    if *r > used_row {
                        used_row = *r;
                    }
                    if *c > used_col {
                        used_col = *c;
                    }
                }
            }
        }
        fn sorted_outline(m: &std::collections::HashMap<i32, u8>) -> Vec<(i32, u8)> {
            let mut v: Vec<(i32, u8)> = m.iter().map(|(k, l)| (*k, *l)).collect();
            v.sort_unstable();
            v
        }
        sheets.push(json!({
            "name": ws.name,
            "hidden": !matches!(ws.state, SheetState::Visible),
            "selected": selected_sheet == Some(sheet_index),
            "color": crate::color_json(&ws.color, &theme),
            "frozenRows": ws.frozen_rows,
            "frozenColumns": ws.frozen_columns,
            "showGridLines": ws.show_grid_lines,
            "zoom": ws.views.get(&0).map(|v| v.zoom).unwrap_or(100),
            // Saved view: first row/column of the scrolling pane and the active cell (1-based).
            "topRow": ws.views.get(&0).map(|v| v.top_row).unwrap_or(1),
            "leftColumn": ws.views.get(&0).map(|v| v.left_column).unwrap_or(1),
            "activeRow": ws.views.get(&0).map(|v| v.row).unwrap_or(1),
            "activeCol": ws.views.get(&0).map(|v| v.column).unwrap_or(1),
            "merges": ws.merge_cells,
            "cols": ws.cols.iter().map(|c| json!({"min": c.min, "max": c.max, "width": c.width, "hidden": c.hidden})).collect::<Vec<_>>(),
            "rows": ws.rows.iter().filter(|r| r.custom_height || r.hidden).map(|r| json!({"r": r.r, "height": r.height, "hidden": r.hidden, "custom": r.custom_height})).collect::<Vec<_>>(),
            // Outline groups: [1-based index, level] pairs, sorted.
            "rowOutline": sorted_outline(&ws.row_outline),
            "colOutline": sorted_outline(&ws.col_outline),
            "autoFilter": ws.auto_filter.as_ref().map(|af| json!({
                "r0": af.range[0], "c0": af.range[1], "r1": af.range[2], "c1": af.range[3],
                "columns": af.columns.iter().map(|c| json!({
                    "col": c.col_id,
                    "values": c.values,
                    "blank": c.blank,
                    "custom": c.custom.as_ref().map(|(and, items)| json!({"and": and, "items": items.iter().map(|i| json!({"op": i.operator, "val": i.value})).collect::<Vec<_>>()})),
                    "other": c.other,
                })).collect::<Vec<_>>(),
            })),
            "summaryBelow": ws.outline_summary_below,
            "summaryRight": ws.outline_summary_right,
            "cellCount": cell_count,
            "maxRow": max_row,
            "maxCol": max_col,
            "usedRow": used_row,
            "usedCol": used_col,
        }));
    }
    // The whole style table: 650 entries on the large benchmark model, a few
    // hundred KB at most, and it saves a styles merge per window.
    let mut styles = serde_json::Map::new();
    let count = m.workbook.styles.cell_xfs.len();
    for (i, _) in m.workbook.styles.cell_xfs.iter().enumerate() {
        if let Ok(st) = m.workbook.styles.get_style(i as i32) {
            styles.insert(i.to_string(), style_json(&st, &theme)?);
        }
    }
    let defined_names: Vec<Value> = m
        .get_defined_name_list()
        .into_iter()
        .map(|(name, scope, formula)| json!({"name": name, "formula": formula, "sheetId": scope}))
        .collect();
    serde_json::to_string(&json!({
        "sheets": sheets,
        "styles": styles,
        "styleCount": count,
        "definedNames": defined_names,
    }))
    .map_err(|e| e.to_string())
}

/// Cells of `sheet` within rows `r0..=r1` and columns `c0..=c1` (1-based).
pub fn sheet_window_json(
    m: &Model,
    cached: &Cached,
    sheet: u32,
    r0: i32,
    r1: i32,
    c0: i32,
    c1: i32,
    blank_unevaluated: bool,
) -> Result<String, String> {
    let ws = m
        .workbook
        .worksheets
        .get(sheet as usize)
        .ok_or_else(|| format!("sheet index {sheet} out of range"))?;
    if r1 < r0 || c1 < c0 || r0 < 1 || c0 < 1 {
        return Err("bad window".into());
    }
    let area = (r1 - r0 + 1) as i64 * (c1 - c0 + 1) as i64;
    if area > MAX_WINDOW_CELLS {
        return Err(format!("window of {area} cells exceeds {MAX_WINDOW_CELLS}"));
    }
    let mut table = StyleTable::new(m).with_blank_unevaluated(blank_unevaluated);
    let mut out = String::with_capacity(64 << 10);
    let _ = write!(out, "{{\"sheet\":{sheet},\"r0\":{r0},\"r1\":{r1},\"c0\":{c0},\"c1\":{c1},\"cells\":[");
    let mut first = true;
    // Rows are a hash map: iterate the requested rows, not the map.
    for r in r0..=r1 {
        let Some(row) = ws.sheet_data.get(&r) else { continue };
        let mut cols: Vec<&i32> = row.keys().filter(|c| **c >= c0 && **c <= c1).collect();
        cols.sort_unstable();
        for c in cols {
            if !first {
                out.push(',');
            }
            first = false;
            write_cell(&mut out, m, cached, sheet, r, *c, false, &mut table)?;
        }
    }
    out.push_str("],\"styles\":");
    out.push_str(&serde_json::to_string(&table.styles).map_err(|e| e.to_string())?);
    out.push('}');
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn layout_and_window() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        for r in 1..=100 {
            for c in 1..=10 {
                m.set_user_input(0, r, c, format!("{}", r * c)).unwrap();
            }
        }
        m.set_user_input(0, 101, 1, "=SUM(A1:A100)".to_string()).unwrap();
        m.evaluate();
        let layout: Value = serde_json::from_str(&workbook_layout_json(&m).unwrap()).unwrap();
        assert_eq!(layout["sheets"][0]["name"], "Sheet1");
        assert_eq!(layout["sheets"][0]["maxRow"], 101);
        assert_eq!(layout["sheets"][0]["maxCol"], 10);
        assert!(layout["styles"].as_object().unwrap().len() >= 1);

        let w: Value = serde_json::from_str(&sheet_window_json(&m, &Cached::new(), 0, 50, 60, 2, 4, false).unwrap()).unwrap();
        let cells = w["cells"].as_array().unwrap();
        assert_eq!(cells.len(), 11 * 3);
        assert_eq!(cells[0]["r"], 50);
        assert_eq!(cells[0]["c"], 2);
        assert_eq!(cells[0]["v"], 100.0);
        // Window past the data is empty, not an error; too-large windows are refused.
        let e: Value = serde_json::from_str(&sheet_window_json(&m, &Cached::new(), 0, 500, 600, 1, 5, false).unwrap()).unwrap();
        assert_eq!(e["cells"].as_array().unwrap().len(), 0);
        assert!(sheet_window_json(&m, &Cached::new(), 0, 1, 1000, 1, 1000, false).is_err());
        let f: Value = serde_json::from_str(&sheet_window_json(&m, &Cached::new(), 0, 101, 101, 1, 1, false).unwrap()).unwrap();
        assert_eq!(f["cells"][0]["f"], "=SUM(A1:A100)");
        assert_eq!(f["cells"][0]["v"], 5050.0);
    }
}
