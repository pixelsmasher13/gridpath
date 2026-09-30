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

/// Part-name prefixes whose presence means the structural engine cannot
/// guarantee reference integrity. Fail closed rather than corrupt.
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

/// Refuse structural edits (rename/delete/row-col shifts) when the package
/// carries parts with cell/sheet references we don't rewrite.
pub fn preflight_structure(store: &PartStore) -> Result<(), PatchError> {
    for name in store.names() {
        if let Some(prefix) = UNSUPPORTED_PART_PREFIXES
            .iter()
            .find(|p| name.starts_with(**p))
        {
            let kind = prefix.trim_start_matches("xl/").trim_end_matches('/');
            return Err(PatchError::Unsupported(format!(
                "workbook contains {kind} ({name})"
            )));
        }
    }
    Ok(())
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
fn max_rid(rels_xml: &[u8]) -> Result<u32, PatchError> {
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

/// Insert `element` (serialized empty element) right before `</container>`.
fn insert_before_container_end(
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
    sheet_el.push_attribute(("r:id", rid.as_str()));
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
    preflight_structure(store)?;
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
    for part in store.names() {
        let is_sheet = part.starts_with(&format!("{wb_dir}worksheets/")) && part.ends_with(".xml");
        let is_chart = part.starts_with(&format!("{wb_dir}charts/")) && part.ends_with(".xml");
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

/// Cheap pre-filter: could this part possibly reference `name`? ASCII
/// case-insensitive substring test; non-ASCII names always pass.
pub(super) fn mentions_name(bytes: &[u8], name: &str) -> bool {
    if !name.is_ascii() {
        return true;
    }
    let needle = name.to_ascii_lowercase();
    let needle = needle.as_bytes();
    if needle.is_empty() || bytes.len() < needle.len() {
        return false;
    }
    bytes
        .windows(needle.len())
        .any(|w| w.iter().zip(needle).all(|(a, b)| a.eq_ignore_ascii_case(b)))
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
    preflight_structure(store)?;
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
    for other in store.names() {
        let is_sheet = other.starts_with(&format!("{wb_dir}worksheets/"))
            && other.ends_with(".xml")
            && other != part;
        let is_chart = other.starts_with(&format!("{wb_dir}charts/")) && other.ends_with(".xml");
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
    let plain = format!("{name}!").to_ascii_lowercase();
    let quoted = format!("'{}'!", name.replace('\'', "''")).to_ascii_lowercase();
    let hay: Vec<u8> = bytes.iter().map(|b| b.to_ascii_lowercase()).collect();
    let find = |needle: &[u8]| -> bool {
        !needle.is_empty()
            && hay.len() >= needle.len()
            && hay.windows(needle.len()).any(|w| w == needle)
    };
    find(plain.as_bytes()) || find(quoted.as_bytes())
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
