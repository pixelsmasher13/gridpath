//! Sheet-level structure operations: create, rename, delete.
//!
//! Create is additive and safe with any package content. Rename and delete
//! rewrite sheet-name references package-wide and are guarded: when the
//! package contains parts whose references we can't rewrite (pivots, tables,
//! slicers…) the operation fails closed and the caller falls back visibly.

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, Event};
use quick_xml::Writer;

use super::package::PartStore;
use super::refs;
use super::workbook_xml::{attr_value, local_name, reader, text_minimal, with_attr, xml_err};
use super::PatchError;

const WORKSHEET_CT: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKSHEET_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Parts this engine does not rewrite. Each holds positions on a sheet, or
/// names a sheet: a table's range, a pivot table's place and its cache's
/// source range, a slicer's anchor.
const UNSUPPORTED_PART_PREFIXES: &[&str] = &[
    "xl/pivotTables/",
    "xl/pivotCache/",
    "xl/slicers/",
    "xl/slicerCaches/",
    "xl/tables/",
    "xl/timelines/",
    "xl/timelineCaches/",
    "xl/queryTables/",
];

/// Of those, the ones that sit ON a sheet (the sheet's relationships point
/// at them, or at the table they belong to). The rest hang off the
/// workbook and only matter when they name the sheet.
const SHEET_HOSTED_PREFIXES: &[&str] =
    &["xl/pivotTables/", "xl/slicers/", "xl/tables/", "xl/timelines/", "xl/queryTables/"];

/// What a structural edit does to the sheet it is about.
pub enum StructuralEdit<'a> {
    /// Rows or columns move on the sheet stored at `sheet_part`: every
    /// position on that sheet changes. `sheet_paths` is (name, part) for
    /// every sheet of the workbook.
    Shift { sheet_part: &'a str, sheet_paths: &'a [(String, String)] },
    /// The sheet is renamed or deleted: what names it changes.
    Name,
}

/// Refuse a structural edit of sheet `sheet_name` when it would leave one
/// of the parts above pointing at the wrong place: a part that sits on the
/// sheet whose rows or columns move, or one that names the sheet. A
/// workbook that merely CONTAINS a table or a pivot table somewhere else is
/// edited like any other.
pub fn preflight_structure(store: &mut PartStore, sheet_name: &str, edit: StructuralEdit) -> Result<(), PatchError> {
    let parts: Vec<String> = store
        .names()
        .into_iter()
        .filter(|n| !n.contains("/_rels/") && UNSUPPORTED_PART_PREFIXES.iter().any(|p| n.starts_with(*p)))
        .collect();
    if parts.is_empty() {
        return Ok(());
    }
    let kind_of = |name: &str| -> String {
        let prefix = UNSUPPORTED_PART_PREFIXES.iter().find(|p| name.starts_with(**p)).copied().unwrap_or("");
        prefix.trim_start_matches("xl/").trim_end_matches('/').to_string()
    };

    if let StructuralEdit::Shift { sheet_part, sheet_paths } = edit {
        // Which sheet each hosted part sits on.
        let mut host: std::collections::HashMap<String, String> = Default::default();
        for (_, part) in sheet_paths {
            for target in related_parts(store, part)? {
                if SHEET_HOSTED_PREFIXES.iter().any(|p| target.starts_with(*p)) {
                    host.insert(target, part.clone());
                }
            }
        }
        // A query table can belong to a table instead of to the sheet.
        for (part, sheet) in host.clone() {
            if part.starts_with("xl/tables/") {
                for target in related_parts(store, &part)? {
                    if target.starts_with("xl/queryTables/") {
                        host.entry(target).or_insert_with(|| sheet.clone());
                    }
                }
            }
        }
        for name in parts.iter().filter(|n| SHEET_HOSTED_PREFIXES.iter().any(|p| n.starts_with(*p))) {
            match host.get(name) {
                Some(on) if on != sheet_part => {}
                Some(_) => {
                    return Err(PatchError::Unsupported(format!(
                        "sheet {sheet_name:?} has {} on it ({name})",
                        kind_of(name)
                    )))
                }
                // Not reachable from any sheet: no telling where it sits.
                None => {
                    return Err(PatchError::Unsupported(format!(
                        "workbook contains {} ({name}) that no sheet refers to",
                        kind_of(name)
                    )))
                }
            }
        }
    }

    for name in &parts {
        // A pivot cache's records are the cached rows: data, no references.
        if name.contains("pivotCacheRecords") {
            continue;
        }
        let bytes = store.read(name)?;
        if names_sheet(&bytes, sheet_name) {
            return Err(PatchError::Unsupported(format!(
                "{} ({name}) refers to sheet {sheet_name:?}",
                kind_of(name)
            )));
        }
    }
    Ok(())
}

/// Does a table / pivot / slicer part name sheet `name`? In these parts a
/// sheet is named in two ways: as the qualifier of a reference in a
/// formula (`Data!B2`, in a table's calculated column), and as a `sheet`
/// attribute (a pivot cache's `<worksheetSource sheet="Data">`). The bare
/// name is not looked for: a sheet called Data would be "named" by every
/// `firstDataRow` and `dataDxfId` in the workbook.
fn names_sheet(bytes: &[u8], name: &str) -> bool {
    if !name.is_ascii() {
        return mentions_name(bytes, name);
    }
    if mentions_qualifier(bytes, name) {
        return true;
    }
    xml_spellings(name).iter().any(|spelled| {
        [format!("sheet=\"{spelled}\""), format!("sheet='{spelled}'")]
            .iter()
            .any(|attr| contains_ignore_ascii_case(bytes, attr.as_bytes()))
    })
}

/// The parts `part`'s own relationships point at (external targets left out).
fn related_parts(store: &mut PartStore, part: &str) -> Result<Vec<String>, PatchError> {
    let (dir, file) = match part.rfind('/') {
        Some(i) => (&part[..=i], &part[i + 1..]),
        None => ("", part),
    };
    let rels_path = format!("{dir}_rels/{file}.rels");
    if !store.exists(&rels_path) {
        return Ok(Vec::new());
    }
    let rels = store.read(&rels_path)?;
    Ok(super::row_col::list_relationships(&rels)?
        .into_iter()
        .filter(|(_, target)| !target.contains("://"))
        .map(|(_, target)| super::row_col::resolve_rel_target(dir, &target))
        .collect())
}

// ---------------------------------------------------------------------------
// shared workbook.xml helpers
// ---------------------------------------------------------------------------

/// (name, sheetId, rId) for every sheet, in workbook order.
pub fn list_sheets(workbook_xml: &[u8]) -> Result<Vec<(String, u32, String)>, PatchError> {
    let mut out = Vec::new();
    let mut r = reader(workbook_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"sheet" =>
            {
                let name = attr_value(&e, b"name");
                let sheet_id = attr_value(&e, b"sheetId").and_then(|v| v.parse().ok());
                let rid = attr_value(&e, b"id");
                if let (Some(name), Some(sheet_id), Some(rid)) = (name, sheet_id, rid) {
                    out.push((name, sheet_id, rid));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("workbook sheets: {e}"))),
        }
        buf.clear();
    }
    Ok(out)
}

fn sheet_names_eq(a: &str, b: &str) -> bool {
    a.len() == b.len()
        && a.chars()
            .zip(b.chars())
            .all(|(x, y)| x.eq_ignore_ascii_case(&y))
}

/// Max rId number in a rels part ("rId7" → 7).
pub(crate) fn max_rid(rels_xml: &[u8]) -> Result<u32, PatchError> {
    let mut max = 0u32;
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                if let Some(id) = attr_value(&e, b"Id") {
                    if let Some(n) = id.strip_prefix("rId").and_then(|s| s.parse::<u32>().ok()) {
                        max = max.max(n);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
        buf.clear();
    }
    Ok(max)
}

/// The prefix a part's root element binds to the relationships namespace
/// (`r` in `r:id`), if it binds one.
pub(crate) fn relationships_prefix(xml: &[u8]) -> Result<Option<String>, PatchError> {
    let mut r = reader(xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                for attr in e.attributes().with_checks(false).flatten() {
                    if let Some(prefix) = attr.key.as_ref().strip_prefix(b"xmlns:") {
                        if attr.value.as_ref() == REL_NS.as_bytes() {
                            return Ok(Some(String::from_utf8_lossy(prefix).into_owned()));
                        }
                    }
                }
                return Ok(None);
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("root: {e}"))),
        }
        buf.clear();
    }
}

/// Insert `element` (serialized empty element) right before `</container>`.
pub(crate) fn insert_before_container_end(
    xml: &[u8],
    container: &[u8],
    element: BytesStart<'static>,
) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut inserted = false;
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("insert: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::End(e) if !inserted && local_name(e.name().as_ref()) == container => {
                w.write_event(Event::Empty(element.clone())).map_err(xml_err)?;
                inserted = true;
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            // A self-closing container (<sheets/>) can't happen for the
            // containers we use (sheets/Relationships/Types always have
            // children), so Start/Empty pass through.
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    if !inserted {
        return Err(PatchError::Xml(format!(
            "container {} not found",
            String::from_utf8_lossy(container)
        )));
    }
    Ok(w.into_inner())
}

// ---------------------------------------------------------------------------
// create
// ---------------------------------------------------------------------------

pub fn create_sheet(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
    name: &str,
    tab_color: Option<&str>,
) -> Result<(), PatchError> {
    if name.is_empty() || name.len() > 31 * 4 {
        return Err(PatchError::BadValue(format!("bad sheet name {name:?}")));
    }
    let workbook_xml = store.read(workbook_path)?;
    let sheets = list_sheets(&workbook_xml)?;
    // Idempotent: re-sending a create for an existing sheet is a no-op
    // (patches may be re-applied after a save that didn't refresh state).
    if sheets.iter().any(|(n, _, _)| sheet_names_eq(n, name)) {
        return Ok(());
    }

    let sheet_id = sheets.iter().map(|(_, id, _)| *id).max().unwrap_or(0) + 1;
    let rels_xml = store.read(wb_rels_path)?;
    let rid = format!("rId{}", max_rid(&rels_xml)? + 1);

    // First unused worksheets/sheetN.xml part name.
    let mut n = sheets.len() as u32 + 1;
    let part = loop {
        let candidate = format!("{wb_dir}worksheets/sheet{n}.xml");
        if !store.exists(&candidate) {
            break candidate;
        }
        n += 1;
    };

    // Minimal worksheet part.
    let mut w = Writer::new(Vec::new());
    w.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))
        .map_err(xml_err)?;
    let mut ws = BytesStart::new("worksheet");
    ws.push_attribute(("xmlns", MAIN_NS));
    ws.push_attribute(("xmlns:r", REL_NS));
    w.write_event(Event::Start(ws)).map_err(xml_err)?;
    if let Some(color) = tab_color {
        let rgb = color.trim_start_matches('#');
        if rgb.len() == 6 && rgb.bytes().all(|b| b.is_ascii_hexdigit()) {
            w.write_event(Event::Start(BytesStart::new("sheetPr")))
                .map_err(xml_err)?;
            let mut tc = BytesStart::new("tabColor");
            tc.push_attribute(("rgb", format!("FF{}", rgb.to_ascii_uppercase()).as_str()));
            w.write_event(Event::Empty(tc)).map_err(xml_err)?;
            w.write_event(Event::End(BytesEnd::new("sheetPr")))
                .map_err(xml_err)?;
        }
    }
    w.write_event(Event::Empty(BytesStart::new("sheetData")))
        .map_err(xml_err)?;
    w.write_event(Event::End(BytesEnd::new("worksheet")))
        .map_err(xml_err)?;
    store.write(&part, w.into_inner());

    // workbook.xml <sheets> entry.
    let mut sheet_el = BytesStart::new("sheet");
    sheet_el.push_attribute(("name", name));
    sheet_el.push_attribute(("sheetId", sheet_id.to_string().as_str()));
    // The relationship id under the prefix the workbook binds to that
    // namespace. A workbook that binds none on its root (openpyxl declares
    // it on each `<sheet>` instead) gets the declaration on this element:
    // an undeclared prefix is a file that does not open at all.
    match relationships_prefix(&workbook_xml)? {
        Some(prefix) => sheet_el.push_attribute((format!("{prefix}:id").as_str(), rid.as_str())),
        None => {
            sheet_el.push_attribute(("xmlns:r", REL_NS));
            sheet_el.push_attribute(("r:id", rid.as_str()));
        }
    }
    let new_wb = insert_before_container_end(&workbook_xml, b"sheets", sheet_el)?;
    store.write(workbook_path, new_wb);

    // workbook.xml.rels entry. Target is relative to the workbook dir.
    let target = part
        .strip_prefix(wb_dir)
        .unwrap_or(&part)
        .to_string();
    let mut rel = BytesStart::new("Relationship");
    rel.push_attribute(("Id", rid.as_str()));
    rel.push_attribute(("Type", WORKSHEET_REL_TYPE));
    rel.push_attribute(("Target", target.as_str()));
    let new_rels = insert_before_container_end(&rels_xml, b"Relationships", rel)?;
    store.write(wb_rels_path, new_rels);

    // [Content_Types].xml override.
    let ct = store.read("[Content_Types].xml")?;
    let mut ov = BytesStart::new("Override");
    ov.push_attribute(("PartName", format!("/{part}").as_str()));
    ov.push_attribute(("ContentType", WORKSHEET_CT));
    let new_ct = insert_before_container_end(&ct, b"Types", ov)?;
    store.write("[Content_Types].xml", new_ct);

    Ok(())
}

// ---------------------------------------------------------------------------
// rename
// ---------------------------------------------------------------------------

pub fn rename_sheet(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    old_name: &str,
    new_name: &str,
) -> Result<(), PatchError> {
    preflight_structure(store, old_name, StructuralEdit::Name)?;
    let workbook_xml = store.read(workbook_path)?;
    let sheets = list_sheets(&workbook_xml)?;
    if !sheets.iter().any(|(n, _, _)| sheet_names_eq(n, old_name)) {
        // Idempotent: already renamed.
        if sheets.iter().any(|(n, _, _)| sheet_names_eq(n, new_name)) {
            return Ok(());
        }
        return Err(PatchError::SheetNotFound(old_name.to_string()));
    }
    if sheets.iter().any(|(n, _, _)| sheet_names_eq(n, new_name)) {
        return Err(PatchError::BadValue(format!(
            "sheet {new_name:?} already exists"
        )));
    }

    // workbook.xml: the name attribute plus defined-name refs.
    let renamed = rewrite_sheet_name_attr(&workbook_xml, old_name, new_name)?;
    let renamed = rewrite_formula_texts(&renamed, &[b"definedName"], &mut |f| {
        refs::rename_sheet_in_formula(f, old_name, new_name)
    })?;
    store.write(workbook_path, renamed);

    // Every worksheet: formulas (f), CF/DV formulas, ext formulas.
    // Every chart part: series formulas (c:f).
    let chart_parts = crate::charts::chart_parts(store)?;
    for part in store.names() {
        let is_sheet = part.starts_with(&format!("{wb_dir}worksheets/")) && part.ends_with(".xml");
        let is_chart = chart_parts.contains(&part);
        if !is_sheet && !is_chart {
            continue;
        }
        let bytes = store.read(&part)?;
        if !mentions_name(&bytes, old_name) {
            continue;
        }
        let mut dirty = false;
        let out = rewrite_formula_texts(&bytes, &[b"f", b"formula", b"formula1", b"formula2"], &mut |f| {
            let new = refs::rename_sheet_in_formula(f, old_name, new_name)?;
            if new != f {
                dirty = true;
            }
            Ok(new)
        })?;
        if dirty {
            store.write(&part, out);
        }
    }
    Ok(())
}

/// The ways `text` can be spelled in a part's bytes: as it is, and as XML
/// writes it. A sheet called `P&L` is `P&amp;L` in every part that names
/// it, so a search for the plain name alone finds nothing.
fn xml_spellings(text: &str) -> Vec<String> {
    let minimal = text.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;");
    let full = minimal.replace('"', "&quot;").replace('\'', "&apos;");
    let mut out = vec![text.to_string()];
    for s in [minimal, full] {
        if !out.contains(&s) {
            out.push(s);
        }
    }
    out
}

fn contains_ignore_ascii_case(bytes: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty()
        && bytes.len() >= needle.len()
        && bytes
            .windows(needle.len())
            .any(|w| w.iter().zip(needle).all(|(a, b)| a.eq_ignore_ascii_case(b)))
}

/// Cheap pre-filter: could this part possibly reference `name`? ASCII
/// case-insensitive substring test, in each spelling XML gives the name.
/// A miss here means the part is left alone, so it errs on the side of
/// passing: non-ASCII names always pass, and so do names with an
/// apostrophe (formulas double it).
pub(super) fn mentions_name(bytes: &[u8], name: &str) -> bool {
    if !name.is_ascii() || name.contains('\'') {
        return true;
    }
    xml_spellings(name).iter().any(|s| contains_ignore_ascii_case(bytes, s.as_bytes()))
}

/// Sets or clears a sheet's tab colour: `<sheetPr><tabColor rgb="FFRRGGBB"/>`
/// in the worksheet part (sheetPr is the first child of <worksheet>).
pub fn set_tab_color(
    store: &mut PartStore,
    workbook_path: &str,
    wb_rels_path: &str,
    wb_dir: &str,
    name: &str,
    color: Option<&str>,
) -> Result<(), PatchError> {
    let workbook_xml = store.read(workbook_path)?;
    let rels_xml = store.read(wb_rels_path)?;
    let parts = crate::workbook_xml::sheet_part_paths(&workbook_xml, &rels_xml)?;
    let part = parts
        .iter()
        .find(|(n, _)| sheet_names_eq(n, name))
        .map(|(_, p)| p.clone())
        .ok_or_else(|| PatchError::SheetNotFound(name.to_string()))?;
    let part = if part.starts_with('/') { part[1..].to_string() } else if part.starts_with("xl/") { part } else { format!("{wb_dir}{part}") };
    let rgb = match color {
        Some(c) => {
            let hex = c.trim_start_matches('#');
            if hex.len() != 6 || !hex.bytes().all(|b| b.is_ascii_hexdigit()) {
                return Err(PatchError::BadValue(format!("bad tab colour {c:?}")));
            }
            Some(format!("FF{}", hex.to_ascii_uppercase()))
        }
        None => None,
    };
    let xml = store.read(&part)?;
    let mut r = reader(&xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut in_sheet_pr = false;
    let mut had_sheet_pr = false;
    let write_tab_color = |w: &mut Writer<Vec<u8>>, rgb: &Option<String>| -> Result<(), PatchError> {
        if let Some(rgb) = rgb {
            let mut tc = BytesStart::new("tabColor");
            tc.push_attribute(("rgb", rgb.as_str()));
            w.write_event(Event::Empty(tc)).map_err(xml_err)?;
        }
        Ok(())
    };
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("tab colour: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) if local_name(e.name().as_ref()) == b"worksheet" => {
                w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                in_sheet_pr = true;
                had_sheet_pr = true;
                w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
                // tabColor is the first child of sheetPr.
                write_tab_color(&mut w, &rgb)?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                had_sheet_pr = true;
                if rgb.is_some() {
                    let start = BytesStart::from(e.to_owned());
                    w.write_event(Event::Start(start)).map_err(xml_err)?;
                    write_tab_color(&mut w, &rgb)?;
                    w.write_event(Event::End(BytesEnd::new("sheetPr"))).map_err(xml_err)?;
                } else {
                    w.write_event(Event::Empty(e.to_owned())).map_err(xml_err)?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                in_sheet_pr = false;
                w.write_event(Event::End(e.to_owned())).map_err(xml_err)?;
            }
            Event::Empty(e) | Event::Start(e) if in_sheet_pr && local_name(e.name().as_ref()) == b"tabColor" => {
                // old colour dropped (the new one, if any, was written at the start)
            }
            Event::End(e) if in_sheet_pr && local_name(e.name().as_ref()) == b"tabColor" => {}
            Event::Start(e) if !had_sheet_pr && is_first_worksheet_child(local_name(e.name().as_ref())) => {
                // No sheetPr yet: it must precede dimension / sheetViews / sheetFormatPr / cols / sheetData.
                had_sheet_pr = true;
                if rgb.is_some() {
                    w.write_event(Event::Start(BytesStart::new("sheetPr"))).map_err(xml_err)?;
                    write_tab_color(&mut w, &rgb)?;
                    w.write_event(Event::End(BytesEnd::new("sheetPr"))).map_err(xml_err)?;
                }
                w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
            }
            Event::Empty(e) if !had_sheet_pr && is_first_worksheet_child(local_name(e.name().as_ref())) => {
                had_sheet_pr = true;
                if rgb.is_some() {
                    w.write_event(Event::Start(BytesStart::new("sheetPr"))).map_err(xml_err)?;
                    write_tab_color(&mut w, &rgb)?;
                    w.write_event(Event::End(BytesEnd::new("sheetPr"))).map_err(xml_err)?;
                }
                w.write_event(Event::Empty(e.to_owned())).map_err(xml_err)?;
            }
            other => w.write_event(other.into_owned()).map_err(xml_err)?,
        }
        buf.clear();
    }
    store.write(&part, w.into_inner());
    Ok(())
}

fn is_first_worksheet_child(tag: &[u8]) -> bool {
    matches!(tag, b"dimension" | b"sheetViews" | b"sheetFormatPr" | b"cols" | b"sheetData")
}

/// Hides or shows a sheet: sets or removes `state="hidden"` on its
/// `<sheet>` element in workbook.xml. Refuses to hide the last visible sheet.
pub fn set_sheet_visibility(
    store: &mut PartStore,
    workbook_path: &str,
    name: &str,
    hidden: bool,
) -> Result<(), PatchError> {
    let workbook_xml = store.read(workbook_path)?;
    let sheets = list_sheets(&workbook_xml)?;
    if !sheets.iter().any(|(n, _, _)| sheet_names_eq(n, name)) {
        return Err(PatchError::SheetNotFound(name.to_string()));
    }
    let mut r = reader(&workbook_xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut visible_after = 0usize;
    let mut out_events: Vec<Event<'static>> = Vec::new();
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("visibility: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Empty(e) | Event::Start(e) if local_name(e.name().as_ref()) == b"sheet" => {
                let is_target = attr_value(&e, b"name").as_deref().is_some_and(|n| sheet_names_eq(n, name));
                let mut el = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
                let mut now_hidden = false;
                for a in e.attributes().flatten() {
                    let key = a.key.as_ref().to_vec();
                    if key == b"state" {
                        if is_target {
                            continue; // rewritten below
                        }
                        now_hidden = a.value.as_ref() != b"visible";
                    }
                    el.push_attribute((key.as_slice(), a.value.as_ref()));
                }
                if is_target {
                    if hidden {
                        el.push_attribute(("state", "hidden"));
                    }
                    now_hidden = hidden;
                }
                if !now_hidden {
                    visible_after += 1;
                }
                out_events.push(Event::Empty(el.into_owned()));
            }
            other => out_events.push(other.into_owned()),
        }
        buf.clear();
    }
    if visible_after == 0 {
        return Err(PatchError::BadValue("cannot hide the last visible sheet".to_string()));
    }
    for ev in out_events {
        w.write_event(ev).map_err(xml_err)?;
    }
    store.write(workbook_path, w.into_inner());
    Ok(())
}

fn rewrite_sheet_name_attr(
    workbook_xml: &[u8],
    old_name: &str,
    new_name: &str,
) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(workbook_xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("rename: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"sheet"
                    && attr_value(&e, b"name").as_deref().is_some_and(|n| {
                        sheet_names_eq(n, old_name)
                    }) =>
            {
                w.write_event(Event::Empty(with_attr(&e, "name", new_name)))
                    .map_err(xml_err)?;
            }
            Event::Start(e)
                if local_name(e.name().as_ref()) == b"sheet"
                    && attr_value(&e, b"name").as_deref().is_some_and(|n| {
                        sheet_names_eq(n, old_name)
                    }) =>
            {
                w.write_event(Event::Start(with_attr(&e, "name", new_name)))
                    .map_err(xml_err)?;
            }
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}

/// Stream-rewrite `xml`, transforming the text content of every element
/// whose local name is in `tags` with `f`. Elements whose text doesn't
/// change are still re-serialized, but only when the file passed the
/// `mentions_name` pre-filter (so untouched parts stay byte-identical).
pub fn rewrite_formula_texts(
    xml: &[u8],
    tags: &[&[u8]],
    f: &mut dyn FnMut(&str) -> Result<String, PatchError>,
) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut in_tag: Option<Vec<u8>> = None;
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("formula rewrite: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                if tags.contains(&name.as_slice()) {
                    in_tag = Some(name);
                }
                w.write_event(Event::Start(e)).map_err(xml_err)?;
            }
            Event::Text(t) if in_tag.is_some() => {
                let text = t
                    .unescape()
                    .map_err(|e| PatchError::Xml(format!("formula text: {e}")))?;
                let new = f(&text)?;
                w.write_event(Event::Text(text_minimal(&new))).map_err(xml_err)?;
            }
            Event::End(e) => {
                if in_tag.as_deref() == Some(local_name(e.name().as_ref())) {
                    in_tag = None;
                }
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}

// ---------------------------------------------------------------------------
// delete
// ---------------------------------------------------------------------------

pub fn delete_sheet(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
    name: &str,
) -> Result<(), PatchError> {
    // (A sheet that has a table or a pivot table ON it is refused further
    // down, with every other sheet that has parts of its own.)
    preflight_structure(store, name, StructuralEdit::Name)?;
    let workbook_xml = store.read(workbook_path)?;
    let sheets = list_sheets(&workbook_xml)?;
    let Some(pos) = sheets.iter().position(|(n, _, _)| sheet_names_eq(n, name)) else {
        return Ok(()); // idempotent: already gone
    };
    if sheets.len() <= 1 {
        return Err(PatchError::Unsupported(
            "cannot delete the only sheet".into(),
        ));
    }
    let rid = sheets[pos].2.clone();

    // Resolve the sheet's part path via rels.
    let rels_xml = store.read(wb_rels_path)?;
    let target = rel_target_for_id(&rels_xml, &rid)?
        .ok_or_else(|| PatchError::MissingPart(format!("rel {rid}")))?;
    let part = if let Some(stripped) = target.strip_prefix('/') {
        stripped.to_string()
    } else {
        format!("{wb_dir}{target}")
    };

    // Guard: nothing else may reference this sheet by name. Formulas that
    // reference a deleted sheet need #REF! rewrites we don't do yet.
    let chart_parts = crate::charts::chart_parts(store)?;
    for other in store.names() {
        let is_sheet = other.starts_with(&format!("{wb_dir}worksheets/"))
            && other.ends_with(".xml")
            && other != part;
        let is_chart = chart_parts.contains(&other);
        if !is_sheet && !is_chart {
            continue;
        }
        let bytes = store.read(&other)?;
        if mentions_qualifier(&bytes, name) {
            return Err(PatchError::Unsupported(format!(
                "sheet {name:?} is referenced by {other}"
            )));
        }
    }
    // Defined names referencing it (global scope) block the delete too.
    if defined_names_mention(&workbook_xml, name)? {
        return Err(PatchError::Unsupported(format!(
            "sheet {name:?} is referenced by defined names"
        )));
    }

    // Guard: a sheet with its own rels (drawings, comments, tables…) would
    // leave orphaned parts behind. Fail closed for now.
    let (part_dir, part_file) = match part.rfind('/') {
        Some(i) => (&part[..=i], &part[i + 1..]),
        None => ("", part.as_str()),
    };
    let sheet_rels = format!("{part_dir}_rels/{part_file}.rels");
    if store.exists(&sheet_rels) && rels_has_relationships(&store.read(&sheet_rels)?)? {
        return Err(PatchError::Unsupported(format!(
            "sheet {name:?} has attached parts (drawings, comments…)"
        )));
    }

    // Drop: part, empty rels, workbook entry, workbook rel, CT override.
    store.remove(&part);
    if store.exists(&sheet_rels) {
        store.remove(&sheet_rels);
    }

    let wb = drop_sheet_entry(&workbook_xml, name, pos as u32)?;
    store.write(workbook_path, wb);

    let rels = drop_relationship(&rels_xml, &rid)?;
    store.write(wb_rels_path, rels);

    let ct = store.read("[Content_Types].xml")?;
    let ct = super::workbook_xml::strip_content_type_override(&ct, &format!("/{part}"))?;
    store.write("[Content_Types].xml", ct);

    Ok(())
}

fn rel_target_for_id(rels_xml: &[u8], rid: &str) -> Result<Option<String>, PatchError> {
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                if attr_value(&e, b"Id").as_deref() == Some(rid) {
                    return Ok(attr_value(&e, b"Target"));
                }
            }
            Ok(Event::Eof) => return Ok(None),
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
        buf.clear();
    }
}

fn rels_has_relationships(rels_xml: &[u8]) -> Result<bool, PatchError> {
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                return Ok(true)
            }
            Ok(Event::Eof) => return Ok(false),
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
        buf.clear();
    }
}

/// Does any worksheet/chart formula in `bytes` reference sheet `name` as a
/// qualifier? Byte-level heuristic: the name (quoted or not) followed by
/// `!`. Conservative: false positives just block a delete.
fn mentions_qualifier(bytes: &[u8], name: &str) -> bool {
    if !name.is_ascii() {
        // Can't do a cheap scan; be conservative.
        return mentions_name(bytes, name);
    }
    // `bytes` is a part as it is on disk (or text already unescaped): the
    // qualifier is looked for in each spelling XML gives it.
    let plain = format!("{name}!");
    let quoted = format!("'{}'!", name.replace('\'', "''"));
    [plain, quoted]
        .iter()
        .flat_map(|q| xml_spellings(q))
        .any(|q| contains_ignore_ascii_case(bytes, q.as_bytes()))
}

fn defined_names_mention(workbook_xml: &[u8], name: &str) -> Result<bool, PatchError> {
    let mut r = reader(workbook_xml);
    let mut buf = Vec::new();
    let mut in_dn = false;
    let mut found = false;
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if local_name(e.name().as_ref()) == b"definedName" => {
                in_dn = true;
            }
            Ok(Event::Text(t)) if in_dn => {
                let text = t.unescape().unwrap_or_default();
                if mentions_qualifier(text.as_bytes(), name) {
                    found = true;
                }
            }
            Ok(Event::End(e)) if local_name(e.name().as_ref()) == b"definedName" => {
                in_dn = false;
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("workbook: {e}"))),
        }
        buf.clear();
    }
    Ok(found)
}

/// Remove the `<sheet>` entry, fix `localSheetId` on defined names (drop
/// names scoped to the deleted sheet, decrement later ones) and clamp
/// `activeTab`.
fn drop_sheet_entry(
    workbook_xml: &[u8],
    name: &str,
    deleted_index: u32,
) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(workbook_xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut skipping_dn = false; // inside a definedName being dropped
    let mut skipping_sheet = false; // inside the <sheet> being dropped
    let is_deleted_sheet = |e: &BytesStart| {
        local_name(e.name().as_ref()) == b"sheet"
            && attr_value(e, b"name")
                .as_deref()
                .is_some_and(|n| sheet_names_eq(n, name))
    };
    let clamp_active = |e: &BytesStart| -> BytesStart<'static> {
        let active = attr_value(e, b"activeTab")
            .and_then(|v| v.parse::<u32>().ok())
            .unwrap_or(0);
        let new_active = if active > deleted_index { active - 1 } else { 0 };
        with_attr(e, "activeTab", &new_active.to_string())
    };
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("delete: {e}")))?;
        match ev {
            Event::Eof => break,
            _ if skipping_dn || skipping_sheet => match ev {
                Event::End(e) if skipping_dn && local_name(e.name().as_ref()) == b"definedName" => {
                    skipping_dn = false;
                }
                Event::End(e) if skipping_sheet && local_name(e.name().as_ref()) == b"sheet" => {
                    skipping_sheet = false;
                }
                _ => {}
            },
            Event::Empty(ref e) if is_deleted_sheet(e) => {}
            Event::Start(ref e) if is_deleted_sheet(e) => {
                skipping_sheet = true;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"definedName" => {
                match attr_value(&e, b"localSheetId").and_then(|v| v.parse::<u32>().ok()) {
                    Some(id) if id == deleted_index => {
                        skipping_dn = true;
                    }
                    Some(id) if id > deleted_index => {
                        w.write_event(Event::Start(with_attr(
                            &e,
                            "localSheetId",
                            &(id - 1).to_string(),
                        )))
                        .map_err(xml_err)?;
                    }
                    _ => w.write_event(Event::Start(e)).map_err(xml_err)?,
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"workbookView" => {
                w.write_event(Event::Empty(clamp_active(&e))).map_err(xml_err)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"workbookView" => {
                w.write_event(Event::Start(clamp_active(&e))).map_err(xml_err)?;
            }
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}

fn drop_relationship(rels_xml: &[u8], rid: &str) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(rels_xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("rels: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"Relationship"
                    && attr_value(&e, b"Id").as_deref() == Some(rid) => {}
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}
