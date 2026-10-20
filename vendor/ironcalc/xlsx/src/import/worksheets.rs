#![allow(clippy::unwrap_used)]

use ironcalc_base::types::{AutoFilter, FilterColumn, FilterCriterion};
use ironcalc_base::expressions::parser::{
    new_parser_english, static_analysis::add_implicit_intersection, Parser,
};
use std::{collections::HashMap, io::Read, num::ParseIntError};

use ironcalc_base::{
    expressions::{
        parser::{stringify::to_rc_format, DefinedNameS},
        token::{get_error_by_english_name, Error},
        types::CellReferenceRC,
        utils::{column_to_number, parse_reference_a1},
    },
    types::{
        ArrayKind, Cell, Col, Color, Comment, DefinedName, Dxf, FormulaValue, Row, SheetData,
        SheetState, SpillValue, Table, Theme, Worksheet, WorksheetView,
    },
};
use roxmltree::Node;
use thiserror::Error;

use crate::error::XlsxError;
use ironcalc_base::cf_types::CfRule;

use super::{
    conditional_formatting::load_conditional_formatting,
    shared_strings::decode_xlsx_escapes,
    tables::load_table,
    util::{get_attribute, get_bool_false, get_color, get_number},
};

pub(crate) struct Sheet {
    pub(crate) name: String,
    pub(crate) sheet_id: u32,
    pub(crate) id: String,
    pub(crate) state: SheetState,
}

pub(crate) struct WorkbookXML {
    pub(crate) worksheets: Vec<Sheet>,
    pub(crate) defined_names: Vec<DefinedName>,
    /// `<calcPr iterate="1" iterateCount=… iterateDelta=…>`
    pub(crate) iteration: Option<ironcalc_base::types::IterationSettings>,
}

pub(crate) struct Relationship {
    pub(crate) target: String,
    pub(crate) rel_type: String,
}

impl WorkbookXML {
    fn get_defined_names_with_scope(&self) -> Vec<DefinedNameS> {
        let sheet_id_index: Vec<u32> = self.worksheets.iter().map(|s| s.sheet_id).collect();

        let defined_names = self
            .defined_names
            .iter()
            .map(|dn| {
                let index = dn
                    .sheet_id
                    .and_then(|sheet_id| {
                        // returns an Option<usize>
                        sheet_id_index.iter().position(|&x| x == sheet_id)
                    })
                    // convert Option<usize> to Option<u32>
                    .map(|pos| pos as u32);
                (dn.name.clone(), index, dn.formula.clone())
            })
            .collect::<Vec<_>>();
        defined_names
    }
}

fn parse_cell_reference(cell: &str) -> Result<(i32, i32), String> {
    if let Some(r) = parse_reference_a1(cell) {
        Ok((r.row, r.column))
    } else {
        Err(format!("Invalid cell reference: '{cell}'"))
    }
}

fn parse_range(range: &str) -> Result<(i32, i32, i32, i32), String> {
    let parts: Vec<&str> = range.split(':').collect();
    if parts.len() == 1 {
        if let Some(r) = parse_reference_a1(parts[0]) {
            Ok((r.row, r.column, r.row, r.column))
        } else {
            Err(format!("Invalid range: '{range}'"))
        }
    } else if parts.len() == 2 {
        match (parse_reference_a1(parts[0]), parse_reference_a1(parts[1])) {
            (Some(left), Some(right)) => Ok((left.row, left.column, right.row, right.column)),
            _ => Err(format!("Invalid range: '{range}'")),
        }
    } else {
        Err(format!("Invalid range: '{range}'"))
    }
}

#[cfg(test)]
mod test {
    use crate::import::worksheets::parse_range;

    #[test]
    fn test_parse_range() {
        assert!(parse_range("3Aw").is_err());
        assert_eq!(parse_range("A1"), Ok((1, 1, 1, 1)));
        assert_eq!(parse_range("B5:C6"), Ok((5, 2, 6, 3)));
        assert!(parse_range("A1:A2:A3").is_err());
        assert!(parse_range("A1:34").is_err());
        assert!(parse_range("A").is_err());
        assert!(parse_range("12").is_err());
    }
}

fn load_dimension(ws: Node) -> String {
    // <dimension ref="A1:O18"/>
    let application_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("dimension"))
        .collect::<Vec<Node>>();
    if application_nodes.len() == 1 {
        application_nodes[0]
            .attribute("ref")
            .unwrap_or("A1")
            .to_string()
    } else {
        "A1".to_string()
    }
}

fn load_columns(ws: Node, col_outline: &mut HashMap<i32, u8>) -> Result<Vec<Col>, XlsxError> {
    // cols
    // <cols>
    //     <col min="5" max="5" width="38.26953125" customWidth="1"/>
    //     <col min="6" max="6" width="9.1796875" style="1"/>
    //     <col min="8" max="8" width="4" customWidth="1"/>
    // </cols>
    let mut cols = Vec::new();
    let columns = ws
        .children()
        .filter(|n| n.has_tag_name("cols"))
        .collect::<Vec<Node>>();
    if columns.len() == 1 {
        for col in columns[0].children().filter(|n| n.has_tag_name("col")) {
            let min = get_attribute(&col, "min")?;
            let min = min.parse::<i32>()?;
            let max = get_attribute(&col, "max")?;
            let max = max.parse::<i32>()?;
            // `width` is optional (ECMA-376 18.3.1.13): a <col> may exist only
            // to carry style/hidden. Excel always writes it, but files saved
            // by other tools (including GridPath's surgical patcher) may not.
            let width = match col.attribute("width") {
                Some(w) => w.parse::<f64>()?,
                None => 10.0, // IronCalc's default column: 90px / COLUMN_WIDTH_FACTOR(9)
            };
            let custom_width = get_bool_false(col, "customWidth");
            let hidden = get_bool_false(col, "hidden");
            let style = col
                .attribute("style")
                .map(|s| s.parse::<i32>().unwrap_or(0));
            if let Some(level) = col.attribute("outlineLevel").and_then(|v| v.parse::<u8>().ok()) {
                if level > 0 {
                    for c in min..=max {
                        col_outline.insert(c, level);
                    }
                }
            }
            cols.push(Col {
                min,
                max,
                width,
                custom_width,
                style,
                hidden,
            })
        }
    }
    Ok(cols)
}

/// 18.3.1.2 autoFilter: the range plus per-column criteria we can model
/// (value lists and custom comparisons); anything else is flagged `other`.
fn load_auto_filter(ws: Node) -> Result<Option<AutoFilter>, XlsxError> {
    let Some(af) = ws.children().find(|n| n.has_tag_name("autoFilter")) else {
        return Ok(None);
    };
    let Some(reference) = af.attribute("ref") else {
        return Ok(None);
    };
    let (r0, c0, r1, c1) = match parse_range(reference) {
        Ok(r) => r,
        Err(_) => return Ok(None),
    };
    let mut columns = Vec::new();
    for fc in af.children().filter(|n| n.has_tag_name("filterColumn")) {
        let col_id = fc
            .attribute("colId")
            .and_then(|v| v.parse::<i32>().ok())
            .unwrap_or(0);
        let mut column = FilterColumn {
            col_id,
            values: None,
            blank: false,
            custom: None,
            other: false,
        };
        for child in fc.children().filter(|n| n.is_element()) {
            match child.tag_name().name() {
                "filters" => {
                    column.blank = matches!(child.attribute("blank"), Some("1") | Some("true"));
                    let vals: Vec<String> = child
                        .children()
                        .filter(|n| n.has_tag_name("filter"))
                        .filter_map(|n| n.attribute("val").map(|v| v.to_string()))
                        .collect();
                    // <dateGroupItem> entries are not modelled.
                    if child.children().any(|n| n.has_tag_name("dateGroupItem")) {
                        column.other = true;
                    }
                    column.values = Some(vals);
                }
                "customFilters" => {
                    let and = matches!(child.attribute("and"), Some("1") | Some("true"));
                    let items: Vec<FilterCriterion> = child
                        .children()
                        .filter(|n| n.has_tag_name("customFilter"))
                        .map(|n| FilterCriterion {
                            operator: n.attribute("operator").unwrap_or("equal").to_string(),
                            value: n.attribute("val").unwrap_or("").to_string(),
                        })
                        .collect();
                    column.custom = Some((and, items));
                }
                _ => column.other = true,
            }
        }
        columns.push(column);
    }
    Ok(Some(AutoFilter {
        range: [r0, c0, r1, c1],
        columns,
    }))
}

fn load_merge_cells(ws: Node) -> Result<Vec<String>, XlsxError> {
    // 18.3.1.55 Merge Cells
    // <mergeCells count="1">
    //    <mergeCell ref="K7:L10"/>
    // </mergeCells>
    let mut merge_cells = Vec::new();
    let merge_cells_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("mergeCells"))
        .collect::<Vec<Node>>();
    if merge_cells_nodes.len() == 1 {
        for merge_cell in merge_cells_nodes[0].children() {
            let reference = get_attribute(&merge_cell, "ref")?.to_string();
            merge_cells.push(reference);
        }
    }
    Ok(merge_cells)
}

fn load_sheet_color(ws: Node, theme: &Theme) -> Result<Color, XlsxError> {
    // <sheetPr>
    //     <tabColor theme="5" tint="-0.249977111117893"/>
    // </sheetPr>
    let mut color = Color::None;
    let sheet_pr = ws
        .children()
        .filter(|n| n.has_tag_name("sheetPr"))
        .collect::<Vec<Node>>();
    if sheet_pr.len() == 1 {
        let tabs = sheet_pr[0]
            .children()
            .filter(|n| n.has_tag_name("tabColor"))
            .collect::<Vec<Node>>();
        if tabs.len() == 1 {
            color = get_color(tabs[0], theme)?;
        }
    }
    Ok(color)
}

fn load_comments<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    path: &str,
) -> Result<Vec<Comment>, XlsxError> {
    let mut comments = Vec::new();
    let mut file = archive.by_name(path)?;
    let mut text = String::new();
    file.read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;
    let ws = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?;
    // <authors><author>Name</author>…</authors>; comments point into it by index.
    let authors = ws
        .children()
        .find(|n| n.has_tag_name("authors"))
        .map(|a| {
            a.children()
                .filter(|n| n.has_tag_name("author"))
                .map(|n| n.text().unwrap_or("").to_string())
                .collect::<Vec<String>>()
        })
        .unwrap_or_default();
    let comment_list = ws
        .children()
        .filter(|n| n.has_tag_name("commentList"))
        .collect::<Vec<Node>>();
    if comment_list.len() == 1 {
        // Element children only: pretty-printed parts have whitespace text nodes.
        for comment in comment_list[0]
            .children()
            .filter(|n| n.has_tag_name("comment"))
        {
            let text = comment
                .descendants()
                .filter(|n| n.has_tag_name("t"))
                .map(|n| n.text().unwrap_or("").to_string())
                .collect::<Vec<String>>()
                .join("");
            let cell_ref = get_attribute(&comment, "ref")?.to_string();
            let author_name = comment
                .attribute("authorId")
                .and_then(|id| id.parse::<usize>().ok())
                .and_then(|id| authors.get(id))
                .cloned()
                .unwrap_or_default();
            comments.push(Comment {
                text,
                author_name,
                author_id: None,
                cell_ref,
            });
        }
    }

    Ok(comments)
}

#[derive(Error, Debug, PartialEq)]
enum ParseReferenceError {
    #[error("RowError: {0}")]
    RowError(ParseIntError),
    #[error("ColumnError: {0}")]
    ColumnError(String),
}

// This parses Sheet1!AS23 into sheet, column and row
// FIXME: This is buggy. Does not check that is a valid sheet name
// There is a similar named function in ironcalc_base. We probably should fix both at the same time.
// NB: Maybe use regexes for this?
fn parse_reference(s: &str) -> Result<CellReferenceRC, ParseReferenceError> {
    let mut sheet_name = "".to_string();
    let mut column = "".to_string();
    let mut row = "".to_string();
    let mut state = "sheet"; // "sheet", "col", "row"
    for ch in s.chars() {
        match state {
            "sheet" => {
                if ch == '!' {
                    state = "col"
                } else {
                    sheet_name.push(ch);
                }
            }
            "col" => {
                if ch.is_ascii_alphabetic() {
                    column.push(ch);
                } else {
                    state = "row";
                    row.push(ch);
                }
            }
            _ => {
                row.push(ch);
            }
        }
    }
    Ok(CellReferenceRC {
        sheet: sheet_name,
        row: row.parse::<i32>().map_err(ParseReferenceError::RowError)?,
        column: column_to_number(&column).map_err(ParseReferenceError::ColumnError)?,
    })
}

fn from_a1_to_rc(
    formula: String,
    parser: &mut Parser,
    context: String,
    is_array_formula: bool,
) -> Result<String, XlsxError> {
    let cell_reference =
        parse_reference(&context).map_err(|error| XlsxError::Xml(error.to_string()))?;
    let mut t = parser.parse(&formula, &cell_reference);
    if !is_array_formula {
        add_implicit_intersection(&mut t, true);
    }

    Ok(to_rc_format(&t))
}

enum CellArrayKind {
    None,
    DynamicArray(i32, i32),
    ArrayFormula(i32, i32),
}

// FIXME
#[allow(clippy::too_many_arguments)]
fn get_cell_from_excel(
    cell_value: Option<&str>,
    value_metadata: Option<&str>,
    cell_type: &str,
    cell_style: i32,
    formula_index: i32,
    sheet_name: &str,
    cell_ref: &str,
    strings: &mut StringTable<'_>,
    rich_text_inline: Option<String>,
    anchor_cell: Option<(i32, i32)>,
    array_kind: CellArrayKind,
) -> Cell {
    // Possible cell types:
    // 18.18.11 ST_CellType (Cell Type)
    //   b (Boolean)
    //   d (Date)
    //   e (Error)
    //   inlineStr (Inline String)
    //   n (Number)
    //   s (Shared String)
    //   str (String)

    if formula_index == -1 {
        match cell_type {
            "b" => {
                if let Some(anchor) = anchor_cell {
                    Cell::SpillCell {
                        v: SpillValue::Boolean(cell_value == Some("1")),
                        s: cell_style,
                        a: anchor,
                    }
                } else {
                    Cell::BooleanCell {
                        v: cell_value == Some("1"),
                        s: cell_style,
                    }
                }
            }
            "n" => {
                if let Some(anchor) = anchor_cell {
                    Cell::SpillCell {
                        v: SpillValue::Number(
                            cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                        ),
                        s: cell_style,
                        a: anchor,
                    }
                } else {
                    Cell::NumberCell {
                        v: cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
                        s: cell_style,
                    }
                }
            }
            "e" => {
                // For compatibility reasons Excel does not put the value #SPILL! but adds it as a metadata
                // Older engines would just import #VALUE!
                let mut error_name = cell_value.unwrap_or("#ERROR!");
                if error_name == "#VALUE!" && value_metadata.is_some() {
                    error_name = match value_metadata {
                        Some("1") => "#CALC!",
                        Some("2") => "#SPILL!",
                        _ => error_name,
                    }
                }
                if let Some(anchor) = anchor_cell {
                    Cell::SpillCell {
                        v: SpillValue::Error(
                            get_error_by_english_name(error_name).unwrap_or(Error::ERROR),
                        ),
                        s: cell_style,
                        a: anchor,
                    }
                } else {
                    Cell::ErrorCell {
                        ei: get_error_by_english_name(error_name).unwrap_or(Error::ERROR),
                        s: cell_style,
                    }
                }
            }
            "s" => Cell::SharedString {
                si: cell_value.unwrap_or("0").parse::<i32>().unwrap_or(0),
                s: cell_style,
            },
            "str" => {
                let s = decode_xlsx_escapes(cell_value.unwrap_or(""));
                let si = strings.intern(&s);

                if let Some(anchor) = anchor_cell {
                    Cell::SpillCell {
                        v: SpillValue::Text(s),
                        s: cell_style,
                        a: anchor,
                    }
                } else {
                    Cell::SharedString { si, s: cell_style }
                }
            }
            "d" => {
                // Not implemented
                println!("Invalid type (d) in {sheet_name}!{cell_ref}");
                Cell::ErrorCell {
                    ei: Error::NIMPL,
                    s: cell_style,
                }
            }
            "inlineStr" => {
                let s = rich_text_inline.unwrap_or_default();
                let si = strings.intern(&s);

                Cell::SharedString { si, s: cell_style }
            }
            "empty" => Cell::EmptyCell { s: cell_style },
            _ => {
                // error
                if std::env::var("IC_DEBUG").is_ok() {
                    eprintln!("[ic] unexpected type ({cell_type}) in {sheet_name}!{cell_ref}");
                }
                Cell::ErrorCell {
                    ei: Error::ERROR,
                    s: cell_style,
                }
            }
        }
    } else {
        let make_cell = |fv: FormulaValue| match array_kind {
            CellArrayKind::None => Cell::CellFormula {
                f: formula_index,
                s: cell_style,
                v: fv,
            },
            CellArrayKind::DynamicArray(width, height) => Cell::ArrayFormula {
                f: formula_index,
                s: cell_style,
                r: (width, height),
                kind: ArrayKind::Dynamic,
                v: fv,
            },
            CellArrayKind::ArrayFormula(width, height) => Cell::ArrayFormula {
                f: formula_index,
                s: cell_style,
                r: (width, height),
                kind: ArrayKind::Cse,
                v: fv,
            },
        };
        match cell_type {
            "b" => make_cell(FormulaValue::Boolean(cell_value == Some("1"))),
            "n" => make_cell(FormulaValue::Number(
                cell_value.unwrap_or("0").parse::<f64>().unwrap_or(0.0),
            )),
            "e" => {
                // For compatibility reasons Excel does not put the value #SPILL! but adds it as a metadata
                // Older engines would just import #VALUE!
                let mut error_name = cell_value.unwrap_or("#ERROR!");
                if error_name == "#VALUE!" && value_metadata.is_some() {
                    error_name = match value_metadata {
                        Some("1") => "#CALC!",
                        Some("2") => "#SPILL!",
                        _ => error_name,
                    }
                }
                make_cell(FormulaValue::Error {
                    ei: get_error_by_english_name(error_name).unwrap_or(Error::ERROR),
                    o: format!("{sheet_name}!{cell_ref}"),
                    m: cell_value.unwrap_or("#ERROR!").to_string(),
                })
            }
            "s" => {
                // Not implemented
                println!("Invalid type (s) in {sheet_name}!{cell_ref}");
                make_cell(FormulaValue::Error {
                    ei: Error::NIMPL,
                    o: format!("{sheet_name}!{cell_ref}"),
                    m: Error::NIMPL.to_string(),
                })
            }
            "str" => {
                // In Excel and in IronCalc all strings in cells result of a formula are *not* shared strings.
                make_cell(FormulaValue::Text(decode_xlsx_escapes(
                    cell_value.unwrap_or(""),
                )))
            }
            "d" => {
                // Not implemented
                println!("Invalid type (d) in {sheet_name}!{cell_ref}");
                make_cell(FormulaValue::Error {
                    ei: Error::NIMPL,
                    o: format!("{sheet_name}!{cell_ref}"),
                    m: Error::NIMPL.to_string(),
                })
            }
            "inlineStr" => {
                // NB: This is untested, I don't know of any engine that uses inline strings in formulas
                make_cell(FormulaValue::Text(
                    rich_text_inline.unwrap_or("".to_string()),
                ))
            }
            _ => {
                // error
                if std::env::var("IC_DEBUG").is_ok() {
                    eprintln!("[ic] unexpected type ({cell_type}) in {sheet_name}!{cell_ref}");
                }
                make_cell(FormulaValue::Error {
                    ei: Error::ERROR,
                    o: format!("{sheet_name}!{cell_ref}"),
                    m: Error::ERROR.to_string(),
                })
            }
        }
    }
}

/// Resolves a relationship target against the directory of the part that
/// owns the relationship ("xl/worksheets/"). Absolute targets start at the
/// package root.
fn resolve_rel_target(base_dir: &str, target: &str) -> String {
    if let Some(absolute) = target.strip_prefix('/') {
        return absolute.to_string();
    }
    let mut segments: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            s => segments.push(s),
        }
    }
    segments.join("/")
}

fn load_sheet_rels<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    path: &str,
    tables: &mut HashMap<String, Table>,
    sheet_name: &str,
) -> Result<Vec<Comment>, XlsxError> {
    // ...xl/worksheets/sheet6.xml -> xl/worksheets/_rels/sheet6.xml.rels
    let mut comments = Vec::new();
    let (base_dir, file_name) = match path.rfind('/') {
        Some(i) => (&path[..=i], &path[i + 1..]),
        None => ("", path),
    };
    let rels_path = format!("{base_dir}_rels/{file_name}.rels");
    let file = archive.by_name(&rels_path);
    if file.is_err() {
        return Ok(comments);
    }
    let mut text = String::new();
    file.unwrap().read_to_string(&mut text)?;
    let doc = roxmltree::Document::parse(&text)?;

    let rels = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?
        .children()
        .filter(|n| n.has_tag_name("Relationship"))
        .collect::<Vec<Node>>();
    for rel in rels {
        let t = get_attribute(&rel, "Type")?.to_string();
        if t.ends_with("comments") {
            // Target="../comments1.xml", or absolute ("/xl/comments/comment1.xml", openpyxl)
            let target = resolve_rel_target(base_dir, get_attribute(&rel, "Target")?);
            comments = load_comments(archive, &target)?;
        } else if t.ends_with("table") {
            // Target="../tables/table1.xml" or "/xl/tables/table1.xml"
            let path = resolve_rel_target(base_dir, get_attribute(&rel, "Target")?);
            let table = load_table(archive, &path, sheet_name)?;
            tables.insert(table.name.clone(), table);
        }
    }
    Ok(comments)
}

struct SheetView {
    is_selected: bool,
    selected_row: i32,
    selected_column: i32,
    frozen_columns: i32,
    frozen_rows: i32,
    range: [i32; 4],
    show_grid_lines: bool,
    zoom: i32,
    /// First row/column of the scrolling pane (`topLeftCell`), 1-based.
    top_row: i32,
    left_column: i32,
}

impl Default for SheetView {
    fn default() -> Self {
        Self {
            is_selected: false,
            selected_row: 1,
            selected_column: 1,
            frozen_rows: 0,
            frozen_columns: 0,
            range: [1, 1, 1, 1],
            show_grid_lines: true,
            zoom: 100,
            top_row: 1,
            left_column: 1,
        }
    }
}

fn get_sheet_view(ws: Node) -> SheetView {
    // <sheetViews>
    //   <sheetView workbookViewId="0">
    //     <selection activeCell="E10" sqref="E10"/>
    //   </sheetView>
    // </sheetViews>
    // <sheetFormatPr defaultRowHeight="14.5" x14ac:dyDescent="0.35"/>

    // If we have frozen rows and columns:

    // <sheetView tabSelected="1" workbookViewId="0">
    //   <pane xSplit="3" ySplit="2" topLeftCell="D3" activePane="bottomRight" state="frozen"/>
    //   <selection pane="topRight" activeCell="D1" sqref="D1"/>
    //   <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
    //   <selection pane="bottomRight" activeCell="K16" sqref="K16"/>
    // </sheetView>

    // 18.18.52 ST_Pane (Pane Types)
    // bottomLeft, bottomRight, topLeft, topRight

    // NB: bottomLeft is used when only rows are frozen, etc
    // IronCalc ignores all those.

    let mut frozen_rows = 0;
    let mut frozen_columns = 0;

    // In IronCalc there can only be one sheetView
    let sheet_views = ws
        .children()
        .filter(|n| n.has_tag_name("sheetViews"))
        .collect::<Vec<Node>>();

    // We are only expecting one `sheetViews` element. Otherwise return a default
    if sheet_views.len() != 1 {
        return SheetView::default();
    }

    let sheet_view = sheet_views[0]
        .children()
        .filter(|n| n.has_tag_name("sheetView"))
        .collect::<Vec<Node>>();

    // We are only expecting one `sheetView` element. Otherwise return a default
    if sheet_view.len() != 1 {
        return SheetView::default();
    }

    let sheet_view = sheet_view[0];
    let is_selected = sheet_view.attribute("tabSelected").unwrap_or("0") == "1";
    let show_grid_lines = sheet_view.attribute("showGridLines").unwrap_or("1") == "1";
    let zoom = sheet_view
        .attribute("zoomScale")
        .and_then(|z| z.parse::<i32>().ok())
        .filter(|z| (10..=400).contains(z))
        .unwrap_or(100);

    let pane = sheet_view
        .children()
        .filter(|n| n.has_tag_name("pane"))
        .collect::<Vec<Node>>();

    // 18.18.53 ST_PaneState (Pane State)
    // frozen, frozenSplit, split. Excel shows frozenSplit as frozen panes
    // (a split that was frozen); only a plain split is not frozen.
    // The scrolling pane's top-left cell is the pane's topLeftCell when
    // frozen, else the sheetView's own.
    let mut top_left = sheet_view.attribute("topLeftCell");
    if pane.len() == 1 {
        if let Some("frozen") | Some("frozenSplit") = pane[0].attribute("state") {
            frozen_columns = get_number(pane[0], "xSplit");
            frozen_rows = get_number(pane[0], "ySplit");
            top_left = pane[0].attribute("topLeftCell").or(top_left);
        }
    }
    let (top_row, left_column) = match top_left.map(parse_cell_reference) {
        Some(Ok((r, c))) if r >= 1 && c >= 1 => (r, c),
        _ => (1, 1),
    };
    let selections = sheet_view
        .children()
        .filter(|n| n.has_tag_name("selection"))
        .collect::<Vec<Node>>();

    if let Some(selection) = selections.last() {
        let active_cell = match selection.attribute("activeCell").map(parse_cell_reference) {
            Some(Ok(s)) => Some(s),
            _ => None,
        };
        let sqref = match selection.attribute("sqref").map(parse_range) {
            Some(Ok(s)) => Some(s),
            _ => None,
        };

        let (selected_row, selected_column, row1, column1, row2, column2) =
            match (active_cell, sqref) {
                (Some(cell), Some(range)) => (cell.0, cell.1, range.0, range.1, range.2, range.3),
                (Some(cell), None) => (cell.0, cell.1, cell.0, cell.1, cell.0, cell.1),
                (None, Some(range)) => (range.0, range.1, range.0, range.1, range.2, range.3),
                _ => (1, 1, 1, 1, 1, 1),
            };

        SheetView {
            frozen_rows,
            frozen_columns,
            selected_row,
            selected_column,
            is_selected,
            show_grid_lines,
            range: [row1, column1, row2, column2],
            zoom,
            top_row,
            left_column,
        }
    } else {
        SheetView {
            frozen_rows,
            frozen_columns,
            is_selected,
            show_grid_lines,
            top_row,
            left_column,
            ..Default::default()
        }
    }
}

pub(super) struct SheetSettings {
    pub id: u32,
    pub name: String,
    pub state: SheetState,
    pub comments: Vec<Comment>,
}

#[allow(clippy::too_many_arguments)]
/// Shared-string interning for one sheet's parse. `shared` is the
/// workbook table as loaded (read-only, so sheets can parse in parallel);
/// strings new to this sheet go to `local` and get ids from `base`
/// upward. After a parallel parse `load_sheets` shifts each sheet's local
/// ids by the sizes of the sheets merged before it.
pub(super) struct StringTable<'a> {
    shared: &'a HashMap<String, i32>,
    base: usize,
    local: Vec<String>,
    local_index: HashMap<String, i32>,
}

impl<'a> StringTable<'a> {
    fn new(shared: &'a HashMap<String, i32>, base: usize) -> Self {
        StringTable { shared, base, local: Vec::new(), local_index: HashMap::new() }
    }
    fn intern(&mut self, s: &str) -> i32 {
        if let Some(i) = self.shared.get(s) {
            return *i;
        }
        if let Some(i) = self.local_index.get(s) {
            return *i;
        }
        let i = (self.base + self.local.len()) as i32;
        self.local.push(s.to_string());
        self.local_index.insert(s.to_string(), i);
        i
    }
}

/// One sheet parsed from its XML text, with the strings and dxfs it added
/// (ids from `base` upward, to be shifted when merged).
pub(super) struct ParsedSheet {
    pub worksheet: Worksheet,
    pub is_selected: bool,
    pub new_strings: Vec<String>,
    pub new_dxfs: Vec<Dxf>,
}

/// Parses one worksheet part. Pure over its inputs: the archive is read
/// beforehand, the shared tables are read-only, additions come back in the
/// result — so `load_sheets` can run this on several sheets at once.
#[allow(clippy::too_many_arguments)]
pub(super) fn parse_sheet_xml(
    text: &str,
    settings: SheetSettings,
    worksheets: &[String],
    tables: &HashMap<String, Table>,
    shared_index: &HashMap<String, i32>,
    string_base: usize,
    defined_names: Vec<DefinedNameS>,
    theme: &Theme,
    dxf_base: usize,
) -> Result<ParsedSheet, XlsxError> {
    let mut strings_table = StringTable::new(shared_index, string_base);
    let strings = &mut strings_table;
    let mut local_dxfs: Vec<Dxf> = Vec::new();
    let dxfs = &mut local_dxfs;
    let sheet_name = &settings.name;
    let sheet_id = settings.id;
    let state = &settings.state;

    let doc = roxmltree::Document::parse(text)?;
    let ws = doc
        .root()
        .first_child()
        .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?;
    let mut shared_formulas = Vec::new();
    // First index in `shared_formulas` for each formula string. Replaces a
    // linear scan per formula cell that was quadratic on large sheets.
    let mut formula_index_map: HashMap<String, i32> = HashMap::new();
    // One parser for the whole sheet: building one per formula cell cloned the
    // sheet list, tables and defined names each time (seconds on large files).
    let mut parser = new_parser_english(
        worksheets.to_owned(),
        defined_names.clone(),
        tables.clone(),
    );

    let dimension = load_dimension(ws);

    let sheet_view = get_sheet_view(ws);

    let mut col_outline: HashMap<i32, u8> = HashMap::new();
    let cols = load_columns(ws, &mut col_outline)?;
    let color = load_sheet_color(ws, theme)?;
    // <sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>
    let (outline_summary_below, outline_summary_right) = ws
        .children()
        .find(|n| n.has_tag_name("sheetPr"))
        .and_then(|pr| pr.children().find(|n| n.has_tag_name("outlinePr")))
        .map(|o| (o.attribute("summaryBelow") != Some("0") && o.attribute("summaryBelow") != Some("false"), o.attribute("summaryRight") != Some("0") && o.attribute("summaryRight") != Some("false")))
        .unwrap_or((true, true));
    let mut row_outline: HashMap<i32, u8> = HashMap::new();
    let auto_filter = load_auto_filter(ws)?;

    // sheetData
    // <row r="1" spans="1:15" x14ac:dyDescent="0.35">
    //     <c r="A1" t="s">
    //         <v>0</v>
    //     </c>
    //     <c r="D1">
    //         <f>C1+1</f>
    //     </c>
    // </row>

    // holds the row heights
    let mut rows = Vec::new();
    let mut sheet_data = SheetData::new();
    let sheet_data_nodes = ws
        .children()
        .filter(|n| n.has_tag_name("sheetData"))
        .collect::<Vec<Node>>()[0];

    let default_row_height = 14.5;

    // Map from the formula index in Excel to the index in IronCalc
    let mut index_map = HashMap::new();

    // Cells part of an array formula
    let mut array_cell = HashMap::new();

    for row in sheet_data_nodes.children() {
        // This is the row number 1-indexed
        let mut row_index = match get_attribute(&row, "r") {
            Ok(s) => Some(s.parse::<i32>()?),
            Err(_) => None,
        };
        // `spans` is not used in IronCalc at the moment (it's an optimization)
        // let spans = row.attribute("spans");
        // This is the height of the row
        let has_height_attribute;
        let height = match row.attribute("ht") {
            Some(s) => {
                has_height_attribute = true;
                s.parse::<f64>().unwrap_or(default_row_height)
            }
            None => {
                has_height_attribute = false;
                default_row_height
            }
        };
        let custom_height = get_bool_false(row, "customHeight");
        // The height of the row is always the visible height of the row
        // If custom_height is false that means the height was calculated automatically:
        // for example because a cell has many lines or a larger font

        let row_style = match row.attribute("s") {
            Some(s) => s.parse::<i32>().unwrap_or(0),
            None => 0,
        };
        let custom_format = get_bool_false(row, "customFormat");
        let hidden = get_bool_false(row, "hidden");
        if let (Some(ri), Some(level)) = (row_index, row.attribute("outlineLevel").and_then(|v| v.parse::<u8>().ok())) {
            if level > 0 {
                row_outline.insert(ri, level);
            }
        }

        if let Some(row_index) = row_index {
            if custom_height || custom_format || row_style != 0 || has_height_attribute || hidden {
                rows.push(Row {
                    r: row_index,
                    height,
                    s: row_style,
                    custom_height,
                    custom_format,
                    hidden,
                });
            }
        }

        // Unused attributes:
        // * thickBot, thickTop, ph, collapsed

        let mut data_row = HashMap::new();

        // 18.3.1.4 c (Cell)
        // Child Elements:
        // * v: Cell value
        // * is: Rich Text Inline
        // * f: Formula
        // Attributes:
        // r: reference. A1 style
        // s: style index
        // t: cell type
        // cm: cell metadata (used for dynamic arrays)
        // vm: value metadata (used for #SPILL! and #CALC! errors)
        // ph: Show Phonetic, unused
        for cell in row.children() {
            let cell_ref = get_attribute(&cell, "r")?;
            let (r_index, column_index) = parse_cell_reference(cell_ref).map_err(XlsxError::Xml)?;
            // Update the row_index if it was not set before
            if row_index.is_none() {
                row_index = Some(r_index);
            }

            let value_metadata = cell.attribute("vm");

            // We check the value "v" child.
            let vs: Vec<Node> = cell.children().filter(|n| n.has_tag_name("v")).collect();
            let cell_value = if vs.len() == 1 {
                Some(vs[0].text().unwrap_or(""))
            } else {
                None
            };

            // <c r="A1" t="inlineStr">
            //   <is>
            //     <t>Hello, World!</t>
            //   </is>
            // </c>
            let cell_rich_text_nodes: Vec<Node> =
                cell.children().filter(|n| n.has_tag_name("is")).collect();
            let cell_rich_text = if cell_rich_text_nodes.is_empty() {
                None
            } else {
                let texts: Vec<String> = cell_rich_text_nodes[0]
                    .descendants()
                    .filter(|n| n.has_tag_name("t"))
                    .filter_map(|n| n.text())
                    .map(|s| s.to_string())
                    .collect();

                Some(texts.join(""))
            };

            let cell_metadata = cell.attribute("cm");
            let is_dynamic_array = cell_metadata == Some("1");

            // type, the default type being "n" for number
            // If the cell does not have a value is an empty cell
            let cell_type = match cell.attribute("t") {
                Some(t) => t,
                None => {
                    if cell_value.is_none() {
                        "empty"
                    } else {
                        "n"
                    }
                }
            };

            // style index, the default style is 0
            let cell_style = match cell.attribute("s") {
                Some(s) => s.parse::<i32>().unwrap_or(0),
                None => 0,
            };

            // Check for formula
            // In Excel some formulas are shared and some are not, but in IronCalc all formulas are shared
            // A cell with a "non-shared" formula is like:
            // <c r="E3">
            //   <f>C2+1</f>
            //   <v>3</v>
            // </c>
            // A cell with a shared formula will be either an "anchor" cell:
            // <c r="D2">
            //   <f t="shared" ref="D2:D3" si="0">C2+1</f>
            //   <v>3</v>
            // </c>
            // Or a child cell:
            // <c r="D3">
            //   <f t="shared" si="0"/>
            //   <v>4</v>
            // </c>
            // In IronCalc two cells have the same formula iff the R1C1 representation is the same
            // TODO: This algorithm could end up with "repeated" shared formulas
            //       We could solve that with a second transversal.

            // In Excel a volatile spill formula might have an f element in the spilled cells.
            // But it is not a shared formula. For example:
            // <c r="A19" s="3" cm="1">
            //   <f t="array" aca="1" ref="A19:C21" ca="1">_xlfn.RANDARRAY(3,3, 0, 100,TRUE)</f>
            //   <v>52</v>
            // </c>
            // <c r="B19" s="3">
            //   <f ca="1"/>
            //   <v>20</v>
            // </c>
            // <c r="C19" s="3">
            //   <f ca="1"/>
            //   <v>41</v>
            // </c>
            // aca: Always Calculate Array
            // ca: Calculate Always
            // Those are hints Excel uses to always calculate volatiles
            // We do not use those in IronCalc
            let fs: Vec<Node> = cell.children().filter(|n| n.has_tag_name("f")).collect();
            let mut formula_index = -1;
            let mut array_kind = CellArrayKind::None;
            if fs.len() == 1 {
                // formula types:
                // 18.18.6 ST_CellFormulaType (Formula Type)
                // array (Array Formula) Formula is an array formula.
                // dataTable (Table Formula) Formula is a data table formula.
                // normal (Normal) Formula is a regular cell formula. (Default)
                // shared (Shared Formula) Formula is part of a shared formula.
                let formula_node = fs[0];
                let mut formula_type = formula_node.attribute("t").unwrap_or("normal");
                let formula_ref = formula_node.attribute("ref");
                if formula_type == "normal"
                    && formula_node.attribute("ca") == Some("1")
                    && formula_node.text().is_none()
                    && !formula_node.children().any(|n| n.is_element())
                {
                    // A daughter cell of a shared formula (<f t="shared" ca="1" si="1"/>) is
                    // also empty and may carry ca="1"; only an untyped <f ca="1"/> is a
                    // volatile spill placeholder.
                    // This is a volatile formula that needs to be recalculated at each calculation.
                    // exit the if statement
                    // <f ca="1"/>
                    formula_type = "hint-volatile";
                }
                match formula_type {
                    "shared" => {
                        // We have a shared formula
                        let si = get_attribute(&formula_node, "si")?;
                        let si = si.parse::<i32>()?;
                        match formula_ref {
                            Some(_) => {
                                // It's the anchor cell. We do not use the ref attribute in IronCalc
                                let formula = formula_node.text().unwrap_or("").to_string();
                                let context = format!("{sheet_name}!{cell_ref}");
                                let formula =
                                    from_a1_to_rc(formula, &mut parser, context, false)?;
                                match index_map.get(&si) {
                                    Some(index) => {
                                        // The index for that formula already exists meaning we
                                        // bumped into a daughter cell first: replace the
                                        // placeholder in place. (Vec::insert here would shift
                                        // every later index, corrupting already-assigned cells.)
                                        formula_index = *index;
                                        shared_formulas[formula_index as usize] = formula.clone();
                                        formula_index_map.entry(formula).or_insert(formula_index);
                                    }
                                    None => {
                                        // We haven't met any of the daughter cells
                                        match formula_index_map.get(&formula) {
                                            // The formula is already present, use that index
                                            Some(index) => {
                                                formula_index = *index;
                                            }
                                            None => {
                                                shared_formulas.push(formula.clone());
                                                formula_index = shared_formulas.len() as i32 - 1;
                                                formula_index_map.insert(formula, formula_index);
                                            }
                                        };
                                        index_map.insert(si, formula_index);
                                    }
                                }
                            }
                            None => {
                                // It's a daughter cell
                                match index_map.get(&si) {
                                    Some(index) => {
                                        formula_index = *index;
                                    }
                                    None => {
                                        // Haven't bumped into the anchor cell yet. We insert a placeholder.
                                        // Note that it is perfectly possible that the formula of the anchor cell
                                        // is already in the set of array formulas. This will lead to the above mention duplicity.
                                        // This is not a problem
                                        let placeholder = "".to_string();
                                        shared_formulas.push(placeholder);
                                        formula_index = shared_formulas.len() as i32 - 1;
                                        index_map.insert(si, formula_index);
                                    }
                                }
                            }
                        }
                    }
                    "dataTable" => {
                        return Err(XlsxError::NotImplemented("data table formulas".to_string()));
                    }
                    "array" => {
                        let range = match formula_ref {
                            Some(r) => r,
                            None => {
                                return Err(XlsxError::Xml(
                                    "Array formulas must have a ref attribute".to_string(),
                                ))
                            }
                        };
                        // The reference is the set of cell it spills into.
                        let (row1, column1, row2, column2) = parse_range(range)
                            .map_err(|_| XlsxError::Xml(format!("Invalid range: {}", range)))?;
                        // (row1, colum1) has to be this cell. We need to mark all the other ones as part of the array formula
                        if row1 != r_index || column1 != column_index {
                            return Err(XlsxError::Xml(
                                "The first cell of the range of an array formula must be the anchor cell".to_string(),
                            ));
                        }
                        for r in row1..=row2 {
                            for c in column1..=column2 {
                                if r == row1 && c == column1 {
                                    // skip the anchor cell
                                    continue;
                                }
                                array_cell.insert((r, c), (r_index, column_index));
                            }
                        }
                        if is_dynamic_array {
                            array_kind =
                                CellArrayKind::DynamicArray(column2 - column1 + 1, row2 - row1 + 1);
                        } else {
                            array_kind =
                                CellArrayKind::ArrayFormula(column2 - column1 + 1, row2 - row1 + 1);
                        }
                        let formula = formula_node.text().unwrap_or("").to_string();
                        let context = format!("{sheet_name}!{cell_ref}");
                        let formula = from_a1_to_rc(formula, &mut parser, context, true)?;

                        match formula_index_map.get(&formula) {
                            Some(index) => formula_index = *index,
                            None => {
                                shared_formulas.push(formula.clone());
                                formula_index = shared_formulas.len() as i32 - 1;
                                formula_index_map.insert(formula, formula_index);
                            }
                        }
                    }
                    "normal" => {
                        // Its a cell with a simple formula
                        let formula = formula_node.text().unwrap_or("").to_string();
                        let context = format!("{sheet_name}!{cell_ref}");
                        let formula = from_a1_to_rc(formula, &mut parser, context, false)?;

                        match formula_index_map.get(&formula) {
                            Some(index) => formula_index = *index,
                            None => {
                                shared_formulas.push(formula.clone());
                                formula_index = shared_formulas.len() as i32 - 1;
                                formula_index_map.insert(formula, formula_index);
                            }
                        }
                    }
                    "hint-volatile" => {}
                    _ => {
                        return Err(XlsxError::Xml(format!(
                            "Invalid formula type {formula_type:?}.",
                        )));
                    }
                }
            }
            let anchor_cell = array_cell.get(&(r_index, column_index)).cloned();
            let cell = get_cell_from_excel(
                cell_value,
                value_metadata,
                cell_type,
                cell_style,
                formula_index,
                sheet_name,
                cell_ref,
                strings,
                cell_rich_text,
                anchor_cell,
                array_kind,
            );
            data_row.insert(column_index, cell);
        }
        if let Some(row_index) = row_index {
            sheet_data.insert(row_index, data_row);
        } else {
            return Err(XlsxError::Xml(
                "Row without a row index (r attribute)".to_string(),
            ));
        }
    }

    let merge_cells = load_merge_cells(ws)?;

    let conditional_formatting = load_conditional_formatting(ws, theme, dxfs, dxf_base)?;
    // pageSetup
    // <pageSetup orientation="portrait" r:id="rId1"/>

    let mut views = HashMap::new();
    views.insert(
        0,
        WorksheetView {
            row: sheet_view.selected_row,
            column: sheet_view.selected_column,
            range: sheet_view.range,
            top_row: sheet_view.top_row,
            left_column: sheet_view.left_column,
            zoom: sheet_view.zoom,
        },
    );

    Ok(ParsedSheet {
        worksheet: Worksheet {
            dimension,
            cols,
            rows,
            shared_formulas,
            sheet_data,
            name: sheet_name.to_string(),
            sheet_id,
            state: state.to_owned(),
            color,
            merge_cells,
            comments: settings.comments,
            frozen_rows: sheet_view.frozen_rows,
            frozen_columns: sheet_view.frozen_columns,
            show_grid_lines: sheet_view.show_grid_lines,
            views,
            conditional_formatting,
            row_outline,
            col_outline,
            outline_summary_below,
            outline_summary_right,
            auto_filter,
        },
        is_selected: sheet_view.is_selected,
        new_strings: strings_table.local,
        new_dxfs: local_dxfs,
    })
}

pub(super) fn load_sheets<R: Read + std::io::Seek>(
    archive: &mut zip::read::ZipArchive<R>,
    rels: &HashMap<String, Relationship>,
    workbook: &WorkbookXML,
    tables: &mut HashMap<String, Table>,
    shared_strings: &mut Vec<String>,
    theme: &Theme,
    dxfs: &mut Vec<Dxf>,
) -> Result<(Vec<Worksheet>, u32), XlsxError> {
    // load comments and tables
    let mut comments = HashMap::new();
    for sheet in &workbook.worksheets {
        let rel = &rels[&sheet.id];
        if rel.rel_type.ends_with("worksheet") {
            let path = &rel.target;
            let path = if let Some(p) = path.strip_prefix('/') {
                p.to_string()
            } else {
                format!("xl/{path}")
            };
            comments.insert(
                &sheet.id,
                load_sheet_rels(archive, &path, tables, &sheet.name)?,
            );
        }
    }

    // load all sheets
    let worksheets: &Vec<String> = &workbook.worksheets.iter().map(|s| s.name.clone()).collect();
    let defined_names = workbook.get_defined_names_with_scope();

    // GRIDPATH: the worksheet parts are read sequentially (one zip reader),
    // then parsed in parallel — each sheet's XML parse, cell loop and
    // formula conversion are independent. Strings and dxfs a sheet adds
    // are returned with it and merged in sheet order afterwards, their ids
    // shifted past the additions of the sheets before it.
    let mut pending: Vec<(SheetSettings, String)> = Vec::new();
    for sheet in &workbook.worksheets {
        let rel_id = &sheet.id;
        let rel = &rels[rel_id];
        if rel.rel_type.ends_with("worksheet") {
            let path = &rel.target;
            let path = if let Some(p) = path.strip_prefix('/') {
                p.to_string()
            } else {
                format!("xl/{path}")
            };
            let settings = SheetSettings {
                name: sheet.name.to_string(),
                id: sheet.sheet_id,
                state: sheet.state.clone(),
                comments: comments
                    .get(rel_id)
                    .ok_or_else(|| XlsxError::Xml("Corrupt XML structure".to_string()))?
                    .to_vec(),
            };
            let mut file = archive.by_name(&path)?;
            let mut text = String::new();
            file.read_to_string(&mut text)?;
            pending.push((settings, text));
        }
    }

    let shared_index: HashMap<String, i32> = {
        let mut m = HashMap::with_capacity(shared_strings.len());
        for (i, s) in shared_strings.iter().enumerate() {
            m.entry(s.clone()).or_insert(i as i32);
        }
        m
    };
    let string_base = shared_strings.len();
    let dxf_base = dxfs.len();
    let tables_ref: &HashMap<String, Table> = tables;

    let parse_one = |settings: SheetSettings, text: &str| -> Result<ParsedSheet, XlsxError> {
        parse_sheet_xml(text, settings, worksheets, tables_ref, &shared_index, string_base, defined_names.clone(), theme, dxf_base)
    };

    let results: Vec<Result<ParsedSheet, XlsxError>> = parse_sheets_parallel(pending, &parse_one);

    let mut sheets = Vec::with_capacity(results.len());
    let mut selected_sheet = 0;
    let mut string_offset = 0usize;
    let mut dxf_offset = 0usize;
    for (index, result) in results.into_iter().enumerate() {
        let parsed = result?;
        let mut ws = parsed.worksheet;
        if string_offset > 0 {
            for row in ws.sheet_data.values_mut() {
                for cell in row.values_mut() {
                    if let Cell::SharedString { si, .. } = cell {
                        if *si as usize >= string_base {
                            *si += string_offset as i32;
                        }
                    }
                }
            }
        }
        if dxf_offset > 0 {
            for cf in &mut ws.conditional_formatting {
                shift_dxf_id(&mut cf.cf_rule, dxf_base as u32, dxf_offset as u32);
            }
        }
        string_offset += parsed.new_strings.len();
        shared_strings.extend(parsed.new_strings);
        dxf_offset += parsed.new_dxfs.len();
        dxfs.extend(parsed.new_dxfs);
        if parsed.is_selected {
            selected_sheet = index as u32;
        }
        sheets.push(ws);
    }
    Ok((sheets, selected_sheet))
}

/// Shifts a rule's dxf id when it points at a dxf this sheet added.
fn shift_dxf_id(rule: &mut CfRule, base: u32, offset: u32) {
    use CfRule::*;
    let id = match rule {
        CellIs { dxf_id, .. }
        | Formula { dxf_id, .. }
        | Text { dxf_id, .. }
        | TimePeriod { dxf_id, .. }
        | DuplicateValues { dxf_id, .. }
        | UniqueValues { dxf_id, .. }
        | Blanks { dxf_id, .. }
        | NotBlanks { dxf_id, .. }
        | Errors { dxf_id, .. }
        | NoErrors { dxf_id, .. }
        | AboveAverage { dxf_id, .. }
        | BelowAverage { dxf_id, .. } => dxf_id,
        _ => return,
    };
    if *id >= base {
        *id += offset;
    }
}

/// Runs `parse_one` over every pending sheet, on as many threads as the
/// machine has cores (sequentially on wasm32, which has no threads).
fn parse_sheets_parallel<F>(pending: Vec<(SheetSettings, String)>, parse_one: &F) -> Vec<Result<ParsedSheet, XlsxError>>
where
    F: Fn(SheetSettings, &str) -> Result<ParsedSheet, XlsxError> + Sync,
{
    #[cfg(target_arch = "wasm32")]
    {
        pending.into_iter().map(|(settings, text)| parse_one(settings, &text)).collect()
    }
    #[cfg(not(target_arch = "wasm32"))]
    {
        let threads = std::thread::available_parallelism().map(|n| n.get()).unwrap_or(1).min(pending.len()).max(1);
        // IC_SEQUENTIAL_PARSE=1 forces the single-threaded path (A/B checks).
        if threads <= 1 || std::env::var("IC_SEQUENTIAL_PARSE").is_ok() {
            return pending.into_iter().map(|(settings, text)| parse_one(settings, &text)).collect();
        }
        // Largest parts first, round-robin over the workers, results kept in sheet order.
        let mut order: Vec<usize> = (0..pending.len()).collect();
        order.sort_by_key(|&i| std::cmp::Reverse(pending[i].1.len()));
        let mut buckets: Vec<Vec<usize>> = vec![Vec::new(); threads];
        for (k, i) in order.into_iter().enumerate() {
            buckets[k % threads].push(i);
        }
        let items: Vec<Option<(SheetSettings, String)>> = pending.into_iter().map(Some).collect();
        let slots: Vec<std::sync::Mutex<Option<Result<ParsedSheet, XlsxError>>>> = (0..items.len()).map(|_| std::sync::Mutex::new(None)).collect();
        let items = std::sync::Mutex::new(items);
        std::thread::scope(|scope| {
            for bucket in buckets {
                let items = &items;
                let slots = &slots;
                scope.spawn(move || {
                    for i in bucket {
                        let taken = items.lock().unwrap_or_else(|e| e.into_inner())[i].take();
                        if let Some((settings, text)) = taken {
                            let r = parse_one(settings, &text);
                            *slots[i].lock().unwrap_or_else(|e| e.into_inner()) = Some(r);
                        }
                    }
                });
            }
        });
        slots
            .into_iter()
            .map(|m| m.into_inner().unwrap_or_else(|e| e.into_inner()).unwrap_or_else(|| Err(XlsxError::Xml("sheet parse produced no result".to_string()))))
            .collect()
    }
}

#[cfg(test)]
mod tests {
    use crate::import::worksheets::parse_reference;

    #[test]
    fn parse_reference_works() {
        let cell_reference = parse_reference("📈 Overview!B2");
        assert!(cell_reference.is_ok());
        let cell_reference = cell_reference.unwrap();
        assert_eq!(cell_reference.sheet, "📈 Overview");
    }
}
