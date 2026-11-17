//! Links between workbooks that are open together.
//!
//! A formula `=[NOW.xlsx]Model!B23` in one workbook reads a sheet of another.
//! The reading workbook holds a hidden copy of that sheet's VALUES — a "link
//! sheet" named exactly as formulas spell the prefix (`[NOW.xlsx]Model`, see
//! `Model::add_link_sheet`) — so the link is evaluated, tracked and
//! incrementally recalculated like a reference to any other sheet. This
//! module keeps the copy in step with the workbook it mirrors:
//!
//! * `sync_link_sheet` — bring the whole copy up to date (first attach, after
//!   a full recalculation or a structural edit of the source),
//! * `push_link_cells` — carry over only the cells the source just changed.
//!
//! Both return the cells of the link sheet whose value changed; the caller
//! hands them to `Model::evaluate_edited` as the edited cells.
//!
//! Only values cross. A source cell the engine cannot compute (an add-in
//! call, a link of its own to a workbook that is not open) gives the value
//! its file was saved with — the same value its own readers see.
//!
//! Saving: an .xlsx keeps, for every other workbook it reads, the last value
//! of each cell it reads there, so its numbers are right before the other
//! file is found. `link_cache` is that list, taken from the copies.
//! `unresolved_links` is the other direction: the links of a file that was
//! just opened and that no copy stands behind yet (`[1]Model`).

use std::collections::{BTreeMap, HashMap, HashSet};

use crate::Cached;
use ironcalc_base::expressions::parser::Node;
use ironcalc_base::expressions::token::Error;
use ironcalc_base::new_empty::{is_link_sheet_name, LinkValue};
use ironcalc_base::Model;
use serde::Serialize;
use serde_json::{json, Value};

const LAST_ROW: i32 = 1_048_576;
const LAST_COLUMN: i32 = 16_384;

/// Whether sheet `sheet` of `m` is the hidden copy of another workbook's
/// sheet.
pub fn is_link_sheet(m: &Model, sheet: u32) -> bool {
    m.workbook
        .worksheets
        .get(sheet as usize)
        .is_some_and(|ws| is_link_sheet_name(&ws.name))
}

/// The index of the link sheet called `name` (sheet names match without
/// case), if the workbook has one.
pub fn link_sheet_index(m: &Model, name: &str) -> Option<u32> {
    if !is_link_sheet_name(name) {
        return None;
    }
    let upper = name.to_uppercase();
    m.workbook
        .worksheets
        .iter()
        .position(|ws| ws.name.to_uppercase() == upper)
        .map(|i| i as u32)
}

/// What a cell of the source workbook shows: its live value, or the value the
/// file was saved with when the live one is missing (empty, or an error) and
/// the file had one. The rule `write_cell` applies for the cell itself.
pub fn source_value(source: &Model, source_cached: &Cached, sheet: u32, row: i32, column: i32) -> LinkValue {
    let live = source.link_value(sheet, row, column);
    let missing = match &live {
        LinkValue::Empty | LinkValue::Error(_) => true,
        LinkValue::Text(s) => s.is_empty(),
        LinkValue::Number(_) | LinkValue::Boolean(_) => false,
    };
    if missing {
        match source_cached.get(&(sheet, row, column)) {
            Some(Value::Number(n)) => {
                if let Some(n) = n.as_f64() {
                    return LinkValue::Number(n);
                }
            }
            Some(Value::String(s)) if !s.is_empty() => return LinkValue::Text(s.clone()),
            Some(Value::Bool(b)) => return LinkValue::Boolean(*b),
            _ => {}
        }
    }
    live
}

fn check(target: &Model, link_sheet: u32, source: &Model, source_sheet: u32) -> Result<(), String> {
    if !is_link_sheet(target, link_sheet) {
        return Err(format!("sheet index {link_sheet} is not a link sheet"));
    }
    if source_sheet as usize >= source.workbook.worksheets.len() {
        return Err(format!("source sheet index {source_sheet} out of range"));
    }
    Ok(())
}

/// Makes the link sheet `link_sheet` of `target` hold exactly the values of
/// sheet `source_sheet` of `source`. Returns the (row, column) of every link
/// cell that changed, including the ones removed because the source no
/// longer has them. Does not evaluate.
pub fn sync_link_sheet(
    target: &mut Model,
    link_sheet: u32,
    source: &Model,
    source_cached: &Cached,
    source_sheet: u32,
) -> Result<Vec<(i32, i32)>, String> {
    check(target, link_sheet, source, source_sheet)?;
    let mut changed: Vec<(i32, i32)> = Vec::new();
    let source_ws = &source.workbook.worksheets[source_sheet as usize];
    for (row, columns) in &source_ws.sheet_data {
        for column in columns.keys() {
            let value = source_value(source, source_cached, source_sheet, *row, *column);
            if target.link_value(link_sheet, *row, *column) != value {
                target.set_link_cell(link_sheet, *row, *column, value)?;
                changed.push((*row, *column));
            }
        }
    }
    // Cells the copy still has and the source does not (cleared there, or
    // moved away by a structural edit).
    let stale: Vec<(i32, i32)> = target.workbook.worksheets[link_sheet as usize]
        .sheet_data
        .iter()
        .flat_map(|(row, columns)| columns.keys().map(move |column| (*row, *column)))
        .filter(|(row, column)| {
            source_ws
                .sheet_data
                .get(row)
                .is_none_or(|columns| !columns.contains_key(column))
        })
        .collect();
    for (row, column) in stale {
        target.set_link_cell(link_sheet, row, column, LinkValue::Empty)?;
        changed.push((row, column));
    }
    Ok(changed)
}

/// Carries the listed cells of `source_sheet` over to the link sheet. Returns
/// the ones whose value in the copy changed. Does not evaluate.
pub fn push_link_cells(
    target: &mut Model,
    link_sheet: u32,
    source: &Model,
    source_cached: &Cached,
    source_sheet: u32,
    cells: &[(i32, i32)],
) -> Result<Vec<(i32, i32)>, String> {
    check(target, link_sheet, source, source_sheet)?;
    let mut changed: Vec<(i32, i32)> = Vec::new();
    for &(row, column) in cells {
        if row < 1 || column < 1 {
            continue;
        }
        let value = source_value(source, source_cached, source_sheet, row, column);
        if target.link_value(link_sheet, row, column) != value {
            target.set_link_cell(link_sheet, row, column, value)?;
            changed.push((row, column));
        }
    }
    Ok(changed)
}

/// What a file keeps of one link sheet.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct LinkCache {
    /// The copy's name, as formulas spell it: `[NOW.xlsx]Model`.
    pub sheet: String,
    /// What this workbook reads there: 0-based, inclusive
    /// `[row1, column1, row2, column2]`.
    pub areas: Vec<[u32; 4]>,
    /// The cells inside `areas` that hold a value.
    pub cells: Vec<LinkCacheCell>,
}

#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct LinkCacheCell {
    /// 0-based.
    pub r: u32,
    pub c: u32,
    /// `{"t":"n","n":1}`, `{"t":"s","s":"x"}`, `{"t":"b","b":true}` or
    /// `{"t":"e","e":"#DIV/0!"}`: a value of the save patch.
    pub v: Value,
}

/// One reference of a formula to a link sheet. A part that is not absolute
/// is an offset from the formula's cell, as the parser stores it.
struct LinkRef {
    sheet: u32,
    row1: (bool, i32),
    column1: (bool, i32),
    row2: (bool, i32),
    column2: (bool, i32),
}

/// Calls `visit` on `node` and on every node under it.
fn walk<'a>(node: &'a Node, visit: &mut dyn FnMut(&'a Node)) {
    use Node::*;
    visit(node);
    match node {
        OpRangeKind { left, right }
        | OpConcatenateKind { left, right }
        | OpSumKind { left, right, .. }
        | OpProductKind { left, right, .. }
        | OpPowerKind { left, right }
        | CompareKind { left, right, .. } => {
            walk(left, visit);
            walk(right, visit);
        }
        UnaryKind { right, .. } => walk(right, visit),
        FunctionKind { args, .. } | NamedFunctionKind { args, .. } => {
            for arg in args {
                walk(arg, visit);
            }
        }
        LambdaDefKind { body, .. } => walk(body, visit),
        LambdaCallKind { lambda, args } => {
            walk(lambda, visit);
            for arg in args {
                walk(arg, visit);
            }
        }
        ImplicitIntersection { child, .. } | SpillRangeOperator { child } => walk(child, visit),
        _ => {}
    }
}

fn link_refs(node: &Node, is_link: &[bool]) -> Vec<LinkRef> {
    let mut out = Vec::new();
    let linked = |sheet: u32| is_link.get(sheet as usize).copied().unwrap_or(false);
    walk(node, &mut |n| match n {
        Node::ReferenceKind { sheet_index, absolute_row, absolute_column, row, column, .. }
            if linked(*sheet_index) =>
        {
            out.push(LinkRef {
                sheet: *sheet_index,
                row1: (*absolute_row, *row),
                column1: (*absolute_column, *column),
                row2: (*absolute_row, *row),
                column2: (*absolute_column, *column),
            });
        }
        Node::RangeKind {
            sheet_index,
            absolute_row1,
            absolute_column1,
            row1,
            column1,
            absolute_row2,
            absolute_column2,
            row2,
            column2,
            ..
        } if linked(*sheet_index) => {
            out.push(LinkRef {
                sheet: *sheet_index,
                row1: (*absolute_row1, *row1),
                column1: (*absolute_column1, *column1),
                row2: (*absolute_row2, *row2),
                column2: (*absolute_column2, *column2),
            });
        }
        _ => {}
    });
    out
}

/// The 1-based area `[row1, column1, row2, column2]` a reference covers when
/// its formula sits at (`row`, `column`); `None` when it falls off the sheet.
fn resolve(r: &LinkRef, row: i32, column: i32) -> Option<[i32; 4]> {
    let at = |(absolute, value): (bool, i32), base: i32| if absolute { value } else { value + base };
    let (r1, r2) = (at(r.row1, row), at(r.row2, row));
    let (c1, c2) = (at(r.column1, column), at(r.column2, column));
    let area = [r1.min(r2), c1.min(c2), r1.max(r2), c1.max(c2)];
    (area[0] >= 1 && area[1] >= 1 && area[2] <= LAST_ROW && area[3] <= LAST_COLUMN).then_some(area)
}

fn cache_value(value: LinkValue) -> Option<Value> {
    match value {
        LinkValue::Empty => None,
        LinkValue::Number(n) if n.is_finite() => Some(json!({"t": "n", "n": n})),
        LinkValue::Number(_) => None,
        LinkValue::Text(s) => Some(json!({"t": "s", "s": s})),
        LinkValue::Boolean(b) => Some(json!({"t": "b", "b": b})),
        LinkValue::Error(e) => {
            // A file only holds the errors Excel knows.
            let text = match e {
                Error::NIMPL | Error::ERROR | Error::CIRC => Error::VALUE.to_string(),
                other => other.to_string(),
            };
            Some(json!({"t": "e", "e": text}))
        }
    }
}

/// For every link sheet this workbook reads: where it reads, and the values
/// there. Formulas and defined names both count. A copy nothing reads is left
/// out; one that is read only where it is empty is listed with no cells.
pub fn link_cache(m: &Model) -> Vec<LinkCache> {
    let worksheets = &m.workbook.worksheets;
    let is_link: Vec<bool> = worksheets.iter().map(|ws| is_link_sheet_name(&ws.name)).collect();
    if !is_link.iter().any(|&l| l) {
        return Vec::new();
    }
    let mut areas: BTreeMap<u32, HashSet<[i32; 4]>> = BTreeMap::new();
    for (sheet, ws) in worksheets.iter().enumerate() {
        if is_link[sheet] {
            continue;
        }
        let Some(formulas) = m.parsed_formulas.get(sheet) else { continue };
        // Formulas are shared between cells: find the ones that read a link
        // once, then place them at each of their cells.
        let reading: HashMap<usize, Vec<LinkRef>> = formulas
            .iter()
            .enumerate()
            .filter_map(|(f, (node, _))| {
                let refs = link_refs(node, &is_link);
                (!refs.is_empty()).then_some((f, refs))
            })
            .collect();
        if reading.is_empty() {
            continue;
        }
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                let Some(refs) = cell.get_formula().and_then(|f| reading.get(&(f as usize))) else {
                    continue;
                };
                for r in refs {
                    if let Some(area) = resolve(r, *row, *column) {
                        areas.entry(r.sheet).or_default().insert(area);
                    }
                }
            }
        }
    }
    for (sheet, row1, column1, row2, column2) in m.defined_name_areas() {
        if is_link.get(sheet as usize).copied().unwrap_or(false)
            && row1 >= 1
            && column1 >= 1
            && row2 <= LAST_ROW
            && column2 <= LAST_COLUMN
        {
            areas.entry(sheet).or_default().insert([row1, column1, row2, column2]);
        }
    }

    areas
        .into_iter()
        .map(|(sheet, areas)| {
            let ws = &worksheets[sheet as usize];
            let mut areas: Vec<[i32; 4]> = areas.into_iter().collect();
            areas.sort_unstable();
            let mut cells: BTreeMap<(i32, i32), Value> = BTreeMap::new();
            let mut take = |row: i32, column: i32| {
                if let std::collections::btree_map::Entry::Vacant(slot) = cells.entry((row, column)) {
                    if let Some(v) = cache_value(m.link_value(sheet, row, column)) {
                        slot.insert(v);
                    }
                }
            };
            for &[r1, c1, r2, c2] in &areas {
                if r1 == r2 && c1 == c2 {
                    take(r1, c1);
                    continue;
                }
                // Walk whichever is smaller: the area, or what the copy has.
                let in_columns = |columns: &HashMap<i32, ironcalc_base::types::Cell>, row: i32, take: &mut dyn FnMut(i32, i32)| {
                    if ((c2 - c1 + 1) as usize) <= columns.len() {
                        for column in c1..=c2 {
                            if columns.contains_key(&column) {
                                take(row, column);
                            }
                        }
                    } else {
                        for &column in columns.keys().filter(|c| **c >= c1 && **c <= c2) {
                            take(row, column);
                        }
                    }
                };
                if ((r2 - r1 + 1) as usize) <= ws.sheet_data.len() {
                    for row in r1..=r2 {
                        if let Some(columns) = ws.sheet_data.get(&row) {
                            in_columns(columns, row, &mut take);
                        }
                    }
                } else {
                    for (&row, columns) in ws.sheet_data.iter().filter(|(r, _)| **r >= r1 && **r <= r2) {
                        in_columns(columns, row, &mut take);
                    }
                }
            }
            LinkCache {
                sheet: ws.name.clone(),
                areas: areas
                    .iter()
                    .map(|a| [(a[0] - 1) as u32, (a[1] - 1) as u32, (a[2] - 1) as u32, (a[3] - 1) as u32])
                    .collect(),
                cells: cells
                    .into_iter()
                    .map(|((row, column), v)| LinkCacheCell { r: (row - 1) as u32, c: (column - 1) as u32, v })
                    .collect(),
            }
        })
        .collect()
}

/// Every cell whose formula names a sheet of another workbook, with a copy
/// behind it or not: the cells whose formula text changes when a link is
/// respelled, and the ones a save has to write with the file's spelling.
/// `(sheet, row, column)`, sorted.
pub fn link_formula_cells(m: &Model) -> Vec<(u32, i32, i32)> {
    let worksheets = &m.workbook.worksheets;
    let is_link: Vec<bool> = worksheets.iter().map(|ws| is_link_sheet_name(&ws.name)).collect();
    let linked = |sheet: u32| is_link.get(sheet as usize).copied().unwrap_or(false);
    let mut out: Vec<(u32, i32, i32)> = Vec::new();
    for (sheet, ws) in worksheets.iter().enumerate() {
        if is_link[sheet] {
            continue;
        }
        let Some(formulas) = m.parsed_formulas.get(sheet) else { continue };
        let reading: HashSet<usize> = formulas
            .iter()
            .enumerate()
            .filter(|(_, (node, _))| {
                let mut found = false;
                walk(node, &mut |n| match n {
                    Node::ReferenceKind { sheet_index, .. } | Node::RangeKind { sheet_index, .. } if linked(*sheet_index) => {
                        found = true;
                    }
                    Node::WrongReferenceKind { sheet_name: Some(name), .. }
                    | Node::WrongRangeKind { sheet_name: Some(name), .. }
                        if is_link_sheet_name(name) =>
                    {
                        found = true;
                    }
                    _ => {}
                });
                found
            })
            .map(|(f, _)| f)
            .collect();
        if reading.is_empty() {
            continue;
        }
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_some_and(|f| reading.contains(&(f as usize))) {
                    out.push((sheet as u32, *row, *column));
                }
            }
        }
    }
    out.sort_unstable();
    out
}

/// The value a cell shows, as a save patch holds it (the shape of
/// `LinkCacheCell::v`); `None` for a cell without one.
pub fn patch_value(m: &Model, sheet: u32, row: i32, column: i32) -> Option<Value> {
    cache_value(m.link_value(sheet, row, column))
}

/// The sheets of other workbooks that formulas name and no copy stands
/// behind: `[1]Model` in a file as it was saved, before its links are
/// attached. Sorted; one spelling per name.
pub fn unresolved_links(m: &Model) -> Vec<String> {
    let mut found: BTreeMap<String, String> = BTreeMap::new();
    for formulas in &m.parsed_formulas {
        for (node, _) in formulas {
            walk(node, &mut |n| {
                if let Node::WrongReferenceKind { sheet_name: Some(name), .. }
                | Node::WrongRangeKind { sheet_name: Some(name), .. } = n
                {
                    if is_link_sheet_name(name) {
                        found.entry(name.to_uppercase()).or_insert_with(|| name.clone());
                    }
                }
            });
        }
    }
    found.into_values().collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::build_cached;
    use ironcalc_base::cell::CellValue;

    fn source() -> Model<'static> {
        let mut m = Model::new_empty("NOW", "en", "UTC", "en").unwrap();
        m.rename_sheet_by_index(0, "Model").unwrap();
        m.set_user_input(0, 5, 1, "Period".into()).unwrap();
        m.set_user_input(0, 5, 2, "FY24".into()).unwrap();
        m.set_user_input(0, 5, 3, "FY25".into()).unwrap();
        m.set_user_input(0, 23, 1, "Revenue".into()).unwrap();
        m.set_user_input(0, 23, 2, "100".into()).unwrap();
        m.set_user_input(0, 23, 3, "=B23*2.5".into()).unwrap();
        m.set_user_input(0, 24, 1, "Flag".into()).unwrap();
        m.set_user_input(0, 24, 2, "TRUE".into()).unwrap();
        m.set_user_input(0, 24, 3, "=1/0".into()).unwrap();
        m.evaluate();
        m
    }

    fn target() -> (Model<'static>, u32) {
        let mut m = Model::new_empty("ANT", "en", "UTC", "en").unwrap();
        let link = m.add_link_sheet("[NOW.xlsx]Model").unwrap();
        m.set_user_input(0, 1, 1, "=[NOW.xlsx]Model!C23".into()).unwrap();
        m.set_user_input(0, 2, 1, "=INDEX([NOW.xlsx]Model!A23:C23,MATCH(\"FY25\",[NOW.xlsx]Model!A5:C5,0))".into()).unwrap();
        m.set_user_input(0, 3, 1, "=[NOW.xlsx]Model!C24".into()).unwrap();
        m.set_user_input(0, 4, 1, "=IF([NOW.xlsx]Model!B24,\"yes\",\"no\")".into()).unwrap();
        (m, link)
    }

    fn text(m: &Model, row: i32, column: i32) -> String {
        m.get_formatted_cell_value(0, row, column).unwrap()
    }

    #[test]
    fn a_first_sync_copies_every_value_and_the_links_evaluate() {
        let src = source();
        let (mut dst, link) = target();
        let changed = sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        assert_eq!(changed.len(), 9);
        dst.evaluate();
        assert_eq!(text(&dst, 1, 1), "250");
        assert_eq!(text(&dst, 2, 1), "250");
        assert_eq!(text(&dst, 3, 1), "#DIV/0!");
        assert_eq!(text(&dst, 4, 1), "yes");
        // Nothing differs any more.
        assert!(sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap().is_empty());
    }

    #[test]
    fn a_source_edit_is_pushed_and_recalculated_incrementally() {
        let mut src = source();
        let (mut dst, link) = target();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.evaluate();

        src.set_user_input(0, 23, 2, "400".into()).unwrap();
        let moved = src.evaluate_edited(&[(0, 23, 2)]);
        let mut cells: Vec<(i32, i32)> = moved.iter().map(|&(_, r, c)| (r, c)).collect();
        cells.push((23, 2));
        // A cell that did not change is not reported.
        cells.push((5, 1));
        let mut changed = push_link_cells(&mut dst, link, &src, &Cached::new(), 0, &cells).unwrap();
        changed.sort_unstable();
        assert_eq!(changed, vec![(23, 2), (23, 3)]);
        let edited: Vec<(u32, i32, i32)> = changed.iter().map(|&(r, c)| (link, r, c)).collect();
        dst.evaluate_edited(&edited);
        assert_eq!(text(&dst, 1, 1), "1000");
        assert_eq!(text(&dst, 2, 1), "1000");

        // The same state as a full sync and a full evaluation.
        let (mut again, link2) = target();
        sync_link_sheet(&mut again, link2, &src, &Cached::new(), 0).unwrap();
        again.evaluate();
        for row in 1..=4 {
            assert_eq!(text(&dst, row, 1), text(&again, row, 1));
        }
    }

    #[test]
    fn a_cleared_source_cell_is_removed_from_the_copy() {
        let mut src = source();
        let (mut dst, link) = target();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.evaluate();
        src.range_clear_all(&ironcalc_base::expressions::types::Area { sheet: 0, row: 24, column: 2, width: 1, height: 1 })
            .unwrap();
        src.evaluate();
        let changed = sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        assert_eq!(changed, vec![(24, 2)]);
        assert_eq!(dst.link_value(link, 24, 2), LinkValue::Empty);
        dst.evaluate();
        assert_eq!(text(&dst, 4, 1), "no");
    }

    #[test]
    fn a_source_cell_without_a_live_value_gives_the_value_its_file_was_saved_with() {
        // An add-in call the engine cannot compute, saved by Excel as 42.
        let mut src = Model::new_empty("NOW", "en", "UTC", "en").unwrap();
        src.rename_sheet_by_index(0, "Model").unwrap();
        src.set_user_input(0, 1, 1, "=_xll.FDS(\"NOW\",\"PRICE\")".into()).unwrap();
        src.evaluate();
        assert!(matches!(src.link_value(0, 1, 1), LinkValue::Error(_)));
        let mut saved = Cached::new();
        saved.insert((0, 1, 1), serde_json::json!(42.0));
        assert_eq!(source_value(&src, &saved, 0, 1, 1), LinkValue::Number(42.0));
        // With nothing saved the error crosses as an error.
        assert!(matches!(source_value(&src, &Cached::new(), 0, 1, 1), LinkValue::Error(_)));

        let mut dst = Model::new_empty("ANT", "en", "UTC", "en").unwrap();
        let link = dst.add_link_sheet("[NOW.xlsx]Model").unwrap();
        dst.set_user_input(0, 1, 1, "=[NOW.xlsx]Model!A1*2".into()).unwrap();
        sync_link_sheet(&mut dst, link, &src, &saved, 0).unwrap();
        dst.evaluate();
        assert_eq!(dst.get_cell_value_by_index(0, 1, 1).unwrap(), CellValue::Number(84.0));
    }

    #[test]
    fn the_copy_is_not_shown_as_a_sheet_of_the_workbook() {
        let src = source();
        let (mut dst, link) = target();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.evaluate();
        // Snapshot: the link sheet keeps its position and carries no cells.
        let snapshot: Value = serde_json::from_str(&crate::snapshot_json(&dst, &Cached::new()).unwrap()).unwrap();
        let sheets = snapshot["sheets"].as_array().unwrap();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0]["link"], Value::Null);
        assert_eq!(sheets[0]["cells"].as_array().unwrap().len(), 4);
        assert_eq!(sheets[link as usize]["link"], Value::Bool(true));
        assert_eq!(sheets[link as usize]["name"], "[NOW.xlsx]Model");
        assert!(sheets[link as usize]["cells"].as_array().unwrap().is_empty());
        // The index (what describe_workbook reads) leaves it out.
        let index = crate::index::build_workbook_index(&dst).unwrap();
        assert_eq!(index.sheets.len(), 1);
        let (json, _) = crate::index::IndexCache::default().workbook_index_json(&dst).unwrap();
        let cached_index: Value = serde_json::from_str(&json).unwrap();
        assert_eq!(cached_index["sheets"].as_array().unwrap().len(), 1);
        // The layout lists it, flagged.
        let layout: Value = serde_json::from_str(&crate::window::workbook_layout_json(&dst).unwrap()).unwrap();
        assert_eq!(layout["sheets"][link as usize]["link"], Value::Bool(true));
        assert_eq!(layout["sheets"][0]["link"], Value::Bool(false));
    }

    #[test]
    fn the_cache_is_what_the_formulas_read_and_nothing_else() {
        let src = source();
        let (mut dst, link) = target();
        // The same formula in two cells: each reads its own row.
        dst.set_user_input(0, 6, 2, "=[NOW.xlsx]Model!B5".into()).unwrap();
        dst.set_user_input(0, 7, 2, "=[NOW.xlsx]Model!B6".into()).unwrap();
        // A whole column, an absolute cell, and a cell that is empty there.
        dst.set_user_input(0, 8, 1, "=SUM([NOW.xlsx]Model!B:B)+[NOW.xlsx]Model!$A$23".into()).unwrap();
        dst.set_user_input(0, 9, 1, "=[NOW.xlsx]Model!Z99".into()).unwrap();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.evaluate();

        let cache = link_cache(&dst);
        assert_eq!(cache.len(), 1);
        assert_eq!(cache[0].sheet, "[NOW.xlsx]Model");
        assert_eq!(
            cache[0].areas,
            vec![
                [0, 1, 1_048_575, 1], // B:B
                [4, 0, 4, 2],         // A5:C5
                [4, 1, 4, 1],         // B5
                [5, 1, 5, 1],         // B6
                [22, 0, 22, 0],       // $A$23
                [22, 0, 22, 2],       // A23:C23
                [22, 2, 22, 2],       // C23
                [23, 1, 23, 1],       // B24
                [23, 2, 23, 2],       // C24
                [98, 25, 98, 25],     // Z99
            ]
        );
        let cells: Vec<(u32, u32, Value)> = cache[0].cells.iter().map(|c| (c.r, c.c, c.v.clone())).collect();
        assert_eq!(
            cells,
            vec![
                (4, 0, json!({"t": "s", "s": "Period"})),
                (4, 1, json!({"t": "s", "s": "FY24"})),
                (4, 2, json!({"t": "s", "s": "FY25"})),
                (22, 0, json!({"t": "s", "s": "Revenue"})),
                (22, 1, json!({"t": "n", "n": 100.0})),
                (22, 2, json!({"t": "n", "n": 250.0})),
                (23, 1, json!({"t": "b", "b": true})),
                (23, 2, json!({"t": "e", "e": "#DIV/0!"})),
            ]
        );
        // A24 ("Flag") is in the copy and nothing reads it.
        assert!(!cells.iter().any(|(r, c, _)| (*r, *c) == (23, 0)));
        // The shape the save patch takes.
        assert_eq!(
            serde_json::to_value(&cache[0].cells[4]).unwrap(),
            json!({"r": 22, "c": 1, "v": {"t": "n", "n": 100.0}})
        );

        // A workbook that reads nothing has nothing to keep.
        assert!(link_cache(&src).is_empty());
        let mut idle = Model::new_empty("ANT", "en", "UTC", "en").unwrap();
        idle.add_link_sheet("[NOW.xlsx]Model").unwrap();
        assert!(link_cache(&idle).is_empty());
    }

    #[test]
    fn a_defined_name_counts_as_a_reader() {
        let src = source();
        let mut dst = Model::new_empty("ANT", "en", "UTC", "en").unwrap();
        let link = dst.add_link_sheet("[NOW.xlsx]Model").unwrap();
        // (The engine takes a name's formula without the leading "=".)
        dst.new_defined_name("Revenue", None, "'[NOW.xlsx]Model'!$B$23:$C$23").unwrap();
        dst.set_user_input(0, 1, 1, "=SUM(Revenue)".into()).unwrap();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.evaluate();
        assert_eq!(text(&dst, 1, 1), "350");
        let cache = link_cache(&dst);
        assert_eq!(cache.len(), 1);
        assert_eq!(cache[0].areas, vec![[22, 1, 22, 2]]);
        assert_eq!(cache[0].cells.len(), 2);
    }

    #[test]
    fn links_of_a_saved_file_are_listed_until_a_copy_stands_behind_them() {
        let mut dst = Model::new_empty("ANT", "en", "UTC", "en").unwrap();
        dst.set_user_input(0, 1, 1, "=[1]Model!C23*2".into()).unwrap();
        dst.set_user_input(0, 2, 1, "=SUM('[1]Income Statement'!A1:B2)+[1]Model!A1".into()).unwrap();
        dst.set_user_input(0, 3, 1, "=[2]Other!A1+Missing!A1".into()).unwrap();
        assert_eq!(
            unresolved_links(&dst),
            vec!["[1]Income Statement".to_string(), "[1]Model".into(), "[2]Other".into()]
        );
        assert!(link_cache(&dst).is_empty());

        // Attached under the number, then given the file's name: formulas
        // follow, and the cache is listed under the new name.
        let src = source();
        let link = dst.add_link_sheet("[1]Model").unwrap();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.rename_link_sheet(link, "[NOW.xlsx]Model").unwrap();
        dst.evaluate();
        assert_eq!(text(&dst, 1, 1), "500");
        assert_eq!(dst.get_cell_formula(0, 1, 1).unwrap().as_deref(), Some("=[NOW.xlsx]Model!C23*2"));
        // The other sheets of that workbook, and the other workbook, are
        // still what they were.
        assert_eq!(
            dst.get_cell_formula(0, 2, 1).unwrap().as_deref(),
            Some("=SUM('[1]Income Statement'!A1:B2)+[NOW.xlsx]Model!A1")
        );
        assert_eq!(unresolved_links(&dst), vec!["[1]Income Statement".to_string(), "[2]Other".into()]);
        let cache = link_cache(&dst);
        assert_eq!(cache.len(), 1);
        assert_eq!(cache[0].sheet, "[NOW.xlsx]Model");
        assert_eq!(cache[0].areas, vec![[0, 0, 0, 0], [22, 2, 22, 2]]);
        assert_eq!(cache[0].cells.len(), 1);
    }

    #[test]
    fn the_cells_that_name_another_workbook_are_listed() {
        let src = source();
        let (mut dst, link) = target();
        sync_link_sheet(&mut dst, link, &src, &Cached::new(), 0).unwrap();
        dst.set_user_input(0, 6, 2, "=A1+1".into()).unwrap();
        dst.set_user_input(0, 7, 2, "=[2]Other!A1".into()).unwrap();
        dst.set_user_input(0, 8, 2, "=SUM('[Gone.xlsx]Sheet 1'!A1:B2)".into()).unwrap();
        dst.set_user_input(0, 9, 2, "=Missing!A1".into()).unwrap();
        dst.evaluate();
        assert_eq!(
            link_formula_cells(&dst),
            vec![(0, 1, 1), (0, 2, 1), (0, 3, 1), (0, 4, 1), (0, 7, 2), (0, 8, 2)]
        );
        assert_eq!(patch_value(&dst, 0, 1, 1), Some(json!({"t": "n", "n": 250.0})));
        assert_eq!(patch_value(&dst, 0, 4, 1), Some(json!({"t": "s", "s": "yes"})));
        assert_eq!(patch_value(&dst, 0, 3, 1), Some(json!({"t": "e", "e": "#DIV/0!"})));
        assert_eq!(patch_value(&dst, 0, 30, 30), None);
        assert!(link_formula_cells(&src).is_empty());
    }

    #[test]
    fn only_link_sheets_are_written() {
        let src = source();
        let (mut dst, link) = target();
        assert!(sync_link_sheet(&mut dst, 0, &src, &Cached::new(), 0).is_err());
        assert!(push_link_cells(&mut dst, 0, &src, &Cached::new(), 0, &[(1, 1)]).is_err());
        assert!(sync_link_sheet(&mut dst, link, &src, &Cached::new(), 7).is_err());
        assert_eq!(link_sheet_index(&dst, "[now.XLSX]model"), Some(link));
        assert_eq!(link_sheet_index(&dst, "Sheet1"), None);
        assert!(is_link_sheet(&dst, link) && !is_link_sheet(&dst, 0));
        let _ = build_cached(&dst);
    }
}
