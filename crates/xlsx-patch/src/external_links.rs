//! Links to other workbooks: `xl/externalLinks/externalLinkN.xml`.
//!
//! A formula that reads another workbook is stored as `[n]Sheet!A1`. `n` is
//! the 1-based position of an `<externalReference>` in workbook.xml, and each
//! of those points at a part that
//!   - names the other file, through a relationship of its own,
//!   - lists that file's sheet names, and
//!   - keeps the last value of every cell this workbook reads from it, so
//!     the numbers are there before the other file has been found.
//!
//! `read` lists what a file has. `apply` creates a link the file does not
//! have yet, or brings an existing one up to date: sheet names are added and
//! the kept values of the cells the patch names are replaced. Everything
//! else in an existing part stays byte for byte, and so does the path it was
//! saved with.

use std::collections::{BTreeMap, HashMap, HashSet};

use quick_xml::events::{BytesStart, Event};
use serde::Serialize;

use super::package::PartStore;
use super::patch::{CellValue, ExternalLinkPatch, ExternalSheetPatch};
use super::refs::{format_cell_ref, parse_cell_ref};
use super::sheet_xml::{sanitize_xlsx_text, trim_float};
use super::structure::{insert_before_container_end, max_rid};
use super::validate::normalize_path;
use super::workbook_xml::{attr_value, local_name, reader};
use super::PatchError;

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const LINK_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
const LINK_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
const LINK_PATH_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

const MAX_ROW: u32 = 1_048_576;
const MAX_COL: u32 = 16_384;

/// Children of `<workbook>` that come after `<externalReferences>`.
const AFTER_EXTERNAL_REFERENCES: &[&[u8]] = &[
    b"definedNames",
    b"calcPr",
    b"oleSize",
    b"customWorkbookViews",
    b"pivotCaches",
    b"smartTagPr",
    b"smartTagTypes",
    b"webPublishing",
    b"fileRecoveryPr",
    b"webPublishObjects",
    b"extLst",
];

/// One link a file was saved with.
#[derive(Debug, Clone, PartialEq, Serialize)]
pub struct ExternalLinkInfo {
    /// The `n` of `[n]Sheet!A1`.
    pub index: u32,
    /// The other file, as the relationship spells it. `None` for a link that
    /// is not to a workbook (DDE, OLE): it takes an index all the same.
    pub target: Option<String>,
    /// The other file's sheet names, in the order the link lists them.
    pub sheets: Vec<String>,
}

struct Entry {
    part: Option<String>,
    book: bool,
    info: ExternalLinkInfo,
}

/// The links of a workbook, in index order.
pub fn read(bytes: &[u8]) -> Result<Vec<ExternalLinkInfo>, PatchError> {
    let mut store = PartStore::open(bytes)?;
    let root_rels = store.read("_rels/.rels")?;
    let workbook_path = super::find_rel_target(&root_rels, "/officeDocument", "")
        .unwrap_or_else(|| "xl/workbook.xml".to_string());
    let (wb_dir, wb_file) = super::split_part_path(&workbook_path);
    let wb_rels_path = format!("{wb_dir}_rels/{wb_file}.rels");
    Ok(table(&mut store, &workbook_path, &wb_dir, &wb_rels_path)?
        .into_iter()
        .map(|e| e.info)
        .collect())
}

/// Creates and updates links. Runs on the store like every other step, so
/// the validity gate sees what it wrote.
pub(crate) fn apply(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
    links: &[ExternalLinkPatch],
) -> Result<(), PatchError> {
    if links.is_empty() {
        return Ok(());
    }
    let mut entries = table(store, workbook_path, wb_dir, wb_rels_path)?;
    let mut ordered: Vec<&ExternalLinkPatch> = links.iter().collect();
    ordered.sort_by_key(|l| l.index);
    let mut seen: HashSet<u32> = HashSet::new();
    for link in ordered {
        if link.index == 0 || !seen.insert(link.index) {
            return Err(PatchError::BadValue(format!("external link index {}", link.index)));
        }
        check_sheets(link)?;
        let position = (link.index - 1) as usize;
        if position < entries.len() {
            let entry = &entries[position];
            let part = match (&entry.part, entry.book) {
                (Some(part), true) => part.clone(),
                _ => {
                    return Err(PatchError::BadValue(format!(
                        "external link {} is not a link to a workbook",
                        link.index
                    )))
                }
            };
            if link.sheets.is_empty() {
                continue;
            }
            let xml = store.read(&part)?;
            let updated = update_part(&xml, &link.sheets)?;
            store.write(&part, updated);
        } else if position == entries.len() {
            let part = create(store, workbook_path, wb_dir, wb_rels_path, link)?;
            entries.push(Entry {
                part: Some(part),
                book: true,
                info: ExternalLinkInfo {
                    index: link.index,
                    target: Some(link.target.clone()),
                    sheets: link.sheets.iter().map(|s| s.name.clone()).collect(),
                },
            });
        } else {
            return Err(PatchError::BadValue(format!(
                "external link {} would leave a gap: the workbook has {}",
                link.index,
                entries.len()
            )));
        }
    }
    Ok(())
}

fn check_sheets(link: &ExternalLinkPatch) -> Result<(), PatchError> {
    let mut names: HashSet<String> = HashSet::new();
    for sheet in &link.sheets {
        if sheet.name.is_empty() || !names.insert(sheet.name.to_lowercase()) {
            return Err(PatchError::BadValue(format!(
                "external link {}: sheet {:?}",
                link.index, sheet.name
            )));
        }
        let off_sheet = sheet.cells.iter().any(|c| c.r >= MAX_ROW || c.c >= MAX_COL)
            || sheet.areas.iter().any(|a| a[0] > a[2] || a[1] > a[3] || a[2] >= MAX_ROW || a[3] >= MAX_COL);
        if off_sheet {
            return Err(PatchError::BadValue(format!(
                "external link {}: a cell of {:?} is off the sheet",
                link.index, sheet.name
            )));
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// reading
// ---------------------------------------------------------------------------

fn table(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
) -> Result<Vec<Entry>, PatchError> {
    let workbook_xml = store.read(workbook_path)?;
    let ids = reference_ids(&workbook_xml)?;
    if ids.is_empty() {
        return Ok(Vec::new());
    }
    let rels = relationships(&store.read(wb_rels_path)?)?;
    let mut out = Vec::with_capacity(ids.len());
    for (i, id) in ids.iter().enumerate() {
        let mut entry = Entry {
            part: None,
            book: false,
            info: ExternalLinkInfo { index: i as u32 + 1, target: None, sheets: Vec::new() },
        };
        let part = match rels.get(id) {
            Some((target, false)) => Some(match target.strip_prefix('/') {
                Some(absolute) => absolute.to_string(),
                None => normalize_path(&format!("{wb_dir}{target}")),
            }),
            _ => None,
        };
        if let Some(part) = part.filter(|p| store.exists(p)) {
            let scan = scan_book(&store.read(&part)?)?;
            if scan.book.is_some() {
                entry.book = true;
                entry.info.sheets = scan.names;
                let (dir, file) = super::split_part_path(&part);
                let part_rels = format!("{dir}_rels/{file}.rels");
                if let (Some(id), true) = (scan.book_rid, store.exists(&part_rels)) {
                    entry.info.target =
                        relationships(&store.read(&part_rels)?)?.remove(&id).map(|(target, _)| target);
                }
            }
            entry.part = Some(part);
        }
        out.push(entry);
    }
    Ok(out)
}

/// The `r:id` of every `<externalReference>`, in order.
fn reference_ids(workbook_xml: &[u8]) -> Result<Vec<String>, PatchError> {
    let mut out = Vec::new();
    let mut r = reader(workbook_xml);
    loop {
        match r.read_event() {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"externalReference" =>
            {
                out.push(attr_value(&e, b"id").unwrap_or_default());
            }
            Ok(Event::Eof) => return Ok(out),
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("workbook: {e}"))),
        }
    }
}

/// Relationship Id → (Target, is it outside the package).
fn relationships(rels_xml: &[u8]) -> Result<HashMap<String, (String, bool)>, PatchError> {
    let mut out = HashMap::new();
    let mut r = reader(rels_xml);
    loop {
        match r.read_event() {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                if let (Some(id), Some(target)) = (attr_value(&e, b"Id"), attr_value(&e, b"Target")) {
                    let external = attr_value(&e, b"TargetMode").as_deref() == Some("External");
                    out.insert(id, (target, external));
                }
            }
            Ok(Event::Eof) => return Ok(out),
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
    }
}

/// Where an element sits in its part, in bytes.
#[derive(Debug, Clone, Copy)]
struct Span {
    /// `<` of the start tag.
    start: usize,
    /// Just past the start tag.
    open_end: usize,
    /// `<name/>`: no content and no end tag.
    empty: bool,
    /// `<` of the end tag (`open_end` when empty).
    close: usize,
}

#[derive(Default)]
struct BookScan {
    /// Name prefix the part's elements carry: "" or e.g. "x:".
    prefix: String,
    book: Option<Span>,
    book_rid: Option<String>,
    names: Vec<String>,
    names_el: Option<Span>,
    /// Start of the first child that has to come after `<sheetNames>`.
    after_names: Option<usize>,
    data_set: Option<Span>,
    /// Each `<sheetData>`: its sheetId and the bytes it takes, whole.
    data: Vec<(u32, usize, usize)>,
}

fn bom_len(xml: &[u8]) -> usize {
    // quick-xml leaves a UTF-8 BOM out of buffer_position.
    if xml.starts_with(&[0xef, 0xbb, 0xbf]) {
        3
    } else {
        0
    }
}

fn prefix_of(qname: &[u8]) -> String {
    match qname.iter().rposition(|&b| b == b':') {
        Some(i) => String::from_utf8_lossy(&qname[..=i]).into_owned(),
        None => String::new(),
    }
}

fn scan_book(xml: &[u8]) -> Result<BookScan, PatchError> {
    let bom = bom_len(xml);
    let mut r = reader(&xml[bom..]);
    let mut out = BookScan::default();
    let mut depth = 0usize;
    // Depth at which the children of <externalBook> start.
    let mut children: Option<usize> = None;
    let mut in_names = false;
    let mut in_set = false;
    let mut open_data: Option<(u32, usize)> = None;
    loop {
        let start = r.buffer_position() as usize + bom;
        let event = r
            .read_event()
            .map_err(|e| PatchError::Xml(format!("external link: {e}")))?;
        let end = r.buffer_position() as usize + bom;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let qname = e.name();
                let local = local_name(qname.as_ref());
                let span = Span { start, open_end: end, empty, close: end };
                if out.book.is_none() && local == b"externalBook" {
                    out.prefix = prefix_of(qname.as_ref());
                    out.book = Some(span);
                    out.book_rid = attr_value(&e, b"id");
                    if !empty {
                        children = Some(depth + 1);
                    }
                } else if children == Some(depth) {
                    match local {
                        b"sheetNames" if out.names_el.is_none() => {
                            out.names_el = Some(span);
                            in_names = !empty;
                        }
                        b"definedNames" | b"sheetDataSet" => {
                            out.after_names.get_or_insert(start);
                            if local == b"sheetDataSet" && out.data_set.is_none() {
                                out.data_set = Some(span);
                                in_set = !empty;
                            }
                        }
                        _ => {}
                    }
                } else if children.map(|d| d + 1) == Some(depth) {
                    if in_names && local == b"sheetName" {
                        out.names.push(attr_value(&e, b"val").unwrap_or_default());
                    } else if in_set && local == b"sheetData" {
                        match attr_value(&e, b"sheetId").and_then(|v| v.parse::<u32>().ok()) {
                            Some(id) if empty => out.data.push((id, start, end)),
                            Some(id) => open_data = Some((id, start)),
                            None => {}
                        }
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if children == Some(depth) {
                    // A child of the book closes.
                    if in_names {
                        in_names = false;
                        if let Some(el) = out.names_el.as_mut() {
                            el.close = start;
                        }
                    } else if in_set {
                        in_set = false;
                        if let Some(el) = out.data_set.as_mut() {
                            el.close = start;
                        }
                    }
                } else if children.map(|d| d + 1) == Some(depth) {
                    if let Some((id, from)) = open_data.take() {
                        out.data.push((id, from, end));
                    }
                } else if children == Some(depth + 1) {
                    // The book itself closes.
                    if let Some(book) = out.book.as_mut() {
                        book.close = start;
                    }
                    children = None;
                }
            }
            _ => {}
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------
// writing
// ---------------------------------------------------------------------------

fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

fn escape_attr(s: &str) -> String {
    escape_text(s).replace('"', "&quot;")
}

fn sheet_name_el(out: &mut Vec<u8>, p: &str, name: &str) {
    out.extend_from_slice(format!("<{p}sheetName val=\"{}\"/>", escape_attr(name)).as_bytes());
}

/// One kept value. `None` for a number a file cannot hold.
fn cell_xml(p: &str, r: u32, c: u32, v: &CellValue) -> Option<Vec<u8>> {
    let (kind, text) = match v {
        CellValue::Number { n } if n.is_finite() => ("", trim_float(*n)),
        CellValue::Number { .. } => return None,
        CellValue::Str { s } => (" t=\"str\"", escape_text(&sanitize_xlsx_text(s))),
        CellValue::Bool { b } => (" t=\"b\"", if *b { "1" } else { "0" }.to_string()),
        CellValue::Error { e } => (" t=\"e\"", escape_text(e)),
    };
    Some(
        format!("<{p}cell r=\"{}\"{kind}><{p}v>{text}</{p}v></{p}cell>", format_cell_ref(r, c))
            .into_bytes(),
    )
}

/// `cells` (by row, then column) as `<row>` elements.
fn rows_xml(out: &mut Vec<u8>, p: &str, cells: &BTreeMap<(u32, u32), Vec<u8>>) {
    let mut open: Option<u32> = None;
    for ((row, _), bytes) in cells {
        if open != Some(*row) {
            if open.is_some() {
                out.extend_from_slice(format!("</{p}row>").as_bytes());
            }
            out.extend_from_slice(format!("<{p}row r=\"{}\">", row + 1).as_bytes());
            open = Some(*row);
        }
        out.extend_from_slice(bytes);
    }
    if open.is_some() {
        out.extend_from_slice(format!("</{p}row>").as_bytes());
    }
}

/// The patch's cells over what the file kept (`kept`), minus what the
/// patch's areas cover.
fn merged_cells(
    mut kept: BTreeMap<(u32, u32), Vec<u8>>,
    p: &str,
    sheet: &ExternalSheetPatch,
) -> BTreeMap<(u32, u32), Vec<u8>> {
    if !kept.is_empty() {
        let mut singles: HashSet<(u32, u32)> = HashSet::new();
        let mut ranges: Vec<[u32; 4]> = Vec::new();
        for a in &sheet.areas {
            if a[0] == a[2] && a[1] == a[3] {
                singles.insert((a[0], a[1]));
            } else {
                ranges.push(*a);
            }
        }
        kept.retain(|&(r, c), _| {
            !(singles.contains(&(r, c))
                || ranges.iter().any(|a| r >= a[0] && r <= a[2] && c >= a[1] && c <= a[3]))
        });
    }
    for cell in &sheet.cells {
        match cell_xml(p, cell.r, cell.c, &cell.v) {
            Some(bytes) => {
                kept.insert((cell.r, cell.c), bytes);
            }
            None => {
                kept.remove(&(cell.r, cell.c));
            }
        }
    }
    kept
}

/// A new `<sheetData>` for `sheet`.
fn sheet_data_xml(out: &mut Vec<u8>, p: &str, id: u32, sheet: &ExternalSheetPatch) {
    let cells = merged_cells(BTreeMap::new(), p, sheet);
    if cells.is_empty() {
        out.extend_from_slice(format!("<{p}sheetData sheetId=\"{id}\"/>").as_bytes());
        return;
    }
    out.extend_from_slice(format!("<{p}sheetData sheetId=\"{id}\">").as_bytes());
    rows_xml(out, p, &cells);
    out.extend_from_slice(format!("</{p}sheetData>").as_bytes());
}

/// An existing `<sheetData>` (`xml` is the whole element) with the patch's
/// cells merged in. Cells the patch does not cover keep their bytes.
fn merge_sheet_data(xml: &[u8], p: &str, sheet: &ExternalSheetPatch) -> Result<Vec<u8>, PatchError> {
    let bad = |what: &str| PatchError::Unsupported(format!("external link values: {what}"));
    let mut r = reader(xml);
    let mut kept: BTreeMap<(u32, u32), Vec<u8>> = BTreeMap::new();
    let mut open_tag: Vec<u8> = Vec::new();
    let mut tag_name = String::new();
    let mut depth = 0usize;
    let mut in_row = false;
    let mut open_cell: Option<((u32, u32), usize)> = None;
    loop {
        let start = r.buffer_position() as usize;
        let event = r
            .read_event()
            .map_err(|e| PatchError::Xml(format!("external link: {e}")))?;
        let end = r.buffer_position() as usize;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let qname = e.name();
                let local = local_name(qname.as_ref());
                if depth == 0 {
                    tag_name = String::from_utf8_lossy(qname.as_ref()).into_owned();
                    open_tag = xml[start..end].to_vec();
                    if empty {
                        // `<sheetData sheetId="0"/>` → `<sheetData sheetId="0">`
                        open_tag.truncate(open_tag.len().saturating_sub(2));
                        open_tag.push(b'>');
                    }
                } else if depth == 1 && local == b"row" {
                    in_row = !empty;
                } else if depth == 2 && in_row && local == b"cell" {
                    let key = attr_value(&e, b"r")
                        .and_then(|v| parse_cell_ref(&v))
                        .ok_or_else(|| bad("a kept cell has no usable reference"))?;
                    if empty {
                        kept.insert(key, xml[start..end].to_vec());
                    } else {
                        open_cell = Some((key, start));
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 2 {
                    if let Some((key, from)) = open_cell.take() {
                        kept.insert(key, xml[from..end].to_vec());
                    }
                } else if depth == 1 {
                    in_row = false;
                }
            }
            _ => {}
        }
    }
    if open_tag.is_empty() {
        return Err(bad("unreadable sheet"));
    }
    let cells = merged_cells(kept, p, sheet);
    let mut out = open_tag;
    rows_xml(&mut out, p, &cells);
    out.extend_from_slice(format!("</{tag_name}>").as_bytes());
    Ok(out)
}

struct Edit {
    at: usize,
    until: usize,
    with: Vec<u8>,
}

/// `xml` with every edit applied. Edits do not overlap; two at the same
/// place go in the order they were pushed.
fn splice(xml: &[u8], mut edits: Vec<Edit>) -> Vec<u8> {
    edits.sort_by_key(|e| e.at);
    let mut out = Vec::with_capacity(xml.len() + edits.iter().map(|e| e.with.len()).sum::<usize>());
    let mut position = 0usize;
    for edit in edits {
        out.extend_from_slice(&xml[position..edit.at]);
        out.extend_from_slice(&edit.with);
        position = edit.until;
    }
    out.extend_from_slice(&xml[position..]);
    out
}

/// An existing link part with `sheets` brought up to date.
pub(crate) fn update_part(xml: &[u8], sheets: &[ExternalSheetPatch]) -> Result<Vec<u8>, PatchError> {
    let scan = scan_book(xml)?;
    let Some(book) = scan.book else {
        return Err(PatchError::BadValue("external link is not a link to a workbook".into()));
    };
    let p = scan.prefix.as_str();

    // Where each sheet of the patch sits in the link's list of names.
    let mut names = scan.names.clone();
    let mut new_names: Vec<&str> = Vec::new();
    let mut ids: Vec<u32> = Vec::with_capacity(sheets.len());
    for sheet in sheets {
        let wanted = sheet.name.to_lowercase();
        let id = match names.iter().position(|n| n.to_lowercase() == wanted) {
            Some(i) => i,
            None => {
                names.push(sheet.name.clone());
                new_names.push(&sheet.name);
                names.len() - 1
            }
        };
        ids.push(id as u32);
    }

    let mut name_els: Vec<u8> = Vec::new();
    for name in &new_names {
        sheet_name_el(&mut name_els, p, name);
    }
    let wrap = |tag: &str, inner: &[u8]| -> Vec<u8> {
        let mut out = format!("<{p}{tag}>").into_bytes();
        out.extend_from_slice(inner);
        out.extend_from_slice(format!("</{p}{tag}>").as_bytes());
        out
    };

    let mut edits: Vec<Edit> = Vec::new();
    // Values of sheets the link holds none for yet, in sheet order.
    let mut fresh: BTreeMap<u32, Vec<u8>> = BTreeMap::new();
    for (sheet, id) in sheets.iter().zip(&ids) {
        match scan.data.iter().find(|(have, _, _)| have == id) {
            Some(&(_, from, to)) => edits.push(Edit {
                at: from,
                until: to,
                with: merge_sheet_data(&xml[from..to], p, sheet)?,
            }),
            None => {
                let mut el = Vec::new();
                sheet_data_xml(&mut el, p, *id, sheet);
                fresh.insert(*id, el);
            }
        }
    }
    let fresh: Vec<u8> = fresh.into_values().flatten().collect();

    if book.empty {
        // `<externalBook r:id="rId1"/>`: give it content.
        let mut with = xml[book.start..book.open_end].to_vec();
        with.truncate(with.len().saturating_sub(2));
        with.push(b'>');
        with.extend_from_slice(&wrap("sheetNames", &name_els));
        with.extend_from_slice(&wrap("sheetDataSet", &fresh));
        with.extend_from_slice(format!("</{p}externalBook>").as_bytes());
        return Ok(splice(xml, vec![Edit { at: book.start, until: book.open_end, with }]));
    }

    // Names first: when both go in at the end of the book, the names lead.
    if !name_els.is_empty() {
        edits.push(match scan.names_el {
            Some(el) if el.empty => {
                Edit { at: el.start, until: el.open_end, with: wrap("sheetNames", &name_els) }
            }
            Some(el) => Edit { at: el.close, until: el.close, with: name_els },
            None => {
                let at = scan.after_names.unwrap_or(book.close);
                Edit { at, until: at, with: wrap("sheetNames", &name_els) }
            }
        });
    }
    if !fresh.is_empty() {
        edits.push(match scan.data_set {
            Some(el) if el.empty => {
                Edit { at: el.start, until: el.open_end, with: wrap("sheetDataSet", &fresh) }
            }
            Some(el) => Edit { at: el.close, until: el.close, with: fresh },
            None => Edit { at: book.close, until: book.close, with: wrap("sheetDataSet", &fresh) },
        });
    }
    Ok(splice(xml, edits))
}

/// A new link: its part, the relationship that names the other file, and
/// its entries in the workbook, the workbook's relationships and the
/// content types. Returns the part's name.
fn create(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
    link: &ExternalLinkPatch,
) -> Result<String, PatchError> {
    if link.target.trim().is_empty() {
        return Err(PatchError::BadValue(format!("external link {} names no file", link.index)));
    }
    let mut n = 1u32;
    let (part, part_rels) = loop {
        let part = format!("{wb_dir}externalLinks/externalLink{n}.xml");
        let part_rels = format!("{wb_dir}externalLinks/_rels/externalLink{n}.xml.rels");
        if !store.exists(&part) && !store.exists(&part_rels) {
            break (part, part_rels);
        }
        n += 1;
    };

    let mut xml = XML_DECL.as_bytes().to_vec();
    xml.extend_from_slice(
        format!("<externalLink xmlns=\"{MAIN_NS}\"><externalBook xmlns:r=\"{REL_NS}\" r:id=\"rId1\"><sheetNames>")
            .as_bytes(),
    );
    for sheet in &link.sheets {
        sheet_name_el(&mut xml, "", &sheet.name);
    }
    xml.extend_from_slice(b"</sheetNames><sheetDataSet>");
    for (id, sheet) in link.sheets.iter().enumerate() {
        sheet_data_xml(&mut xml, "", id as u32, sheet);
    }
    xml.extend_from_slice(b"</sheetDataSet></externalBook></externalLink>");
    store.write(&part, xml);

    store.write(
        &part_rels,
        format!(
            "{XML_DECL}<Relationships xmlns=\"{PACKAGE_REL_NS}\"><Relationship Id=\"rId1\" Type=\"{LINK_PATH_REL_TYPE}\" Target=\"{}\" TargetMode=\"External\"/></Relationships>",
            escape_attr(&link.target)
        )
        .into_bytes(),
    );

    let rels_xml = store.read(wb_rels_path)?;
    let rid = format!("rId{}", max_rid(&rels_xml)? + 1);
    let target = part.strip_prefix(wb_dir).unwrap_or(&part).to_string();
    let mut rel = BytesStart::new("Relationship");
    rel.push_attribute(("Id", rid.as_str()));
    rel.push_attribute(("Type", LINK_REL_TYPE));
    rel.push_attribute(("Target", target.as_str()));
    store.write(wb_rels_path, insert_before_container_end(&rels_xml, b"Relationships", rel)?);

    let content_types = store.read("[Content_Types].xml")?;
    let mut over = BytesStart::new("Override");
    over.push_attribute(("PartName", format!("/{part}").as_str()));
    over.push_attribute(("ContentType", LINK_CT));
    store.write(
        "[Content_Types].xml",
        insert_before_container_end(&content_types, b"Types", over)?,
    );

    let workbook_xml = store.read(workbook_path)?;
    store.write(workbook_path, add_external_reference(&workbook_xml, &rid)?);
    Ok(part)
}

/// workbook.xml with one more `<externalReference>`, at the end of the list
/// (so the links before it keep their numbers). The list is created, where
/// the schema wants it, when the workbook has none.
pub(crate) fn add_external_reference(xml: &[u8], rid: &str) -> Result<Vec<u8>, PatchError> {
    let bom = bom_len(xml);
    let mut r = reader(&xml[bom..]);
    let mut depth = 0usize;
    let mut prefix = String::new();
    // The prefix the workbook binds to the relationships namespace.
    let mut rel_prefix: Option<String> = None;
    let mut list: Option<Span> = None;
    let mut list_open = false;
    let mut before: Option<usize> = None;
    let mut root_close: Option<usize> = None;
    loop {
        let start = r.buffer_position() as usize + bom;
        let event = r
            .read_event()
            .map_err(|e| PatchError::Xml(format!("workbook: {e}")))?;
        let end = r.buffer_position() as usize + bom;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let qname = e.name();
                let local = local_name(qname.as_ref());
                if depth == 0 {
                    prefix = prefix_of(qname.as_ref());
                    for attr in e.attributes().with_checks(false).flatten() {
                        if let Some(name) = attr.key.as_ref().strip_prefix(b"xmlns:") {
                            if attr.value.as_ref() == REL_NS.as_bytes() {
                                rel_prefix = Some(String::from_utf8_lossy(name).into_owned());
                            }
                        }
                    }
                } else if depth == 1 {
                    if local == b"externalReferences" && list.is_none() {
                        list = Some(Span { start, open_end: end, empty, close: end });
                        list_open = !empty;
                    } else if before.is_none() && AFTER_EXTERNAL_REFERENCES.contains(&local) {
                        before = Some(start);
                    }
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1 {
                    if list_open {
                        list_open = false;
                        if let Some(el) = list.as_mut() {
                            el.close = start;
                        }
                    }
                } else if depth == 0 {
                    root_close = Some(start);
                }
            }
            _ => {}
        }
    }

    let reference = match &rel_prefix {
        Some(rp) => format!("<{prefix}externalReference {rp}:id=\"{}\"/>", escape_attr(rid)),
        None => format!(
            "<{prefix}externalReference xmlns:r=\"{REL_NS}\" r:id=\"{}\"/>",
            escape_attr(rid)
        ),
    }
    .into_bytes();
    let wrapped = || {
        let mut out = format!("<{prefix}externalReferences>").into_bytes();
        out.extend_from_slice(&reference);
        out.extend_from_slice(format!("</{prefix}externalReferences>").as_bytes());
        out
    };
    let edit = match list {
        Some(el) if el.empty => Edit { at: el.start, until: el.open_end, with: wrapped() },
        Some(el) => Edit { at: el.close, until: el.close, with: reference.clone() },
        None => {
            let at = before
                .or(root_close)
                .ok_or_else(|| PatchError::Xml("workbook: no place for externalReferences".into()))?;
            Edit { at, until: at, with: wrapped() }
        }
    };
    Ok(splice(xml, vec![edit]))
}
