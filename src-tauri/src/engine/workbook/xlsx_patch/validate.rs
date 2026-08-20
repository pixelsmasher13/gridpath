//! Save-time validity gate.
//!
//! Excel enforces OOXML rules that lenient tooling (openpyxl, ExcelJS,
//! browsers) silently accepts — child-element order inside a worksheet,
//! singleton containers, monotonically increasing row/cell references,
//! non-empty defined names, content-type/rels referential integrity. Every
//! bug we've shipped that made Excel offer to "repair" a file was one of
//! these. This module re-checks every part a patch rewrote, plus the
//! package-level registries, BEFORE any bytes reach disk: a surgical save
//! either produces a file Excel opens cleanly or refuses to write at all.
//!
//! This is not full XSD validation — it encodes the specific invariant
//! classes our generator can break. Parts we copy raw are correct by
//! construction (they're the original bytes) and are not re-checked.

use std::collections::{HashMap, HashSet};

use quick_xml::events::Event;
use quick_xml::Reader;

use super::package::PartStore;
use super::refs::parse_cell_ref;
use super::sheet_xml::WORKSHEET_CHILD_ORDER;
use super::workbook_xml::{attr_value, local_name};
use super::PatchError;

/// CT_Workbook child sequence (ECMA-376).
const WORKBOOK_CHILD_ORDER: &[&[u8]] = &[
    b"fileVersion",
    b"fileSharing",
    b"workbookPr",
    b"workbookProtection",
    b"bookViews",
    b"sheets",
    b"functionGroups",
    b"externalReferences",
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

/// Top-level elements that may legally repeat.
const WORKSHEET_REPEATABLE: &[&[u8]] = &[b"cols", b"conditionalFormatting"];
const WORKBOOK_REPEATABLE: &[&[u8]] = &[b"fileRecoveryPr"];

fn err(part: &str, what: impl std::fmt::Display) -> PatchError {
    PatchError::Validation(format!("{part}: {what}"))
}

fn strict_reader(xml: &[u8]) -> Reader<&[u8]> {
    let mut r = Reader::from_reader(xml);
    r.config_mut().trim_text(false);
    // Unlike the rewrite readers, validation wants mismatched end tags to
    // be errors.
    r.config_mut().check_end_names = true;
    r
}

/// Validate every part this patch wrote plus package-level registries.
/// Called after all ops are applied, before the zip is rebuilt.
pub fn validate_store(
    store: &mut PartStore,
    workbook_path: &str,
    wb_rels_path: &str,
) -> Result<(), PatchError> {
    for name in store.dirty_names() {
        let bytes = store.read(&name)?;
        validate_part(&name, &bytes, workbook_path)?;
        // .vml and other non-XML parts: rewritten via quick-xml, no strict
        // schema to hold them to.
    }

    // Only parts this patch ADDED must prove content-type coverage:
    // pre-existing parts are raw copies of a file Excel already accepted.
    let added = store.added_names();
    validate_content_types(store, &added)?;
    validate_workbook_rels(store, workbook_path, wb_rels_path)?;
    Ok(())
}

/// Validate a whole package produced by a full re-serialization (the
/// ExcelJS fallback writer). Unlike `validate_store`, nothing here is a
/// trusted raw copy — every XML part is checked, and every part must have
/// content-type coverage.
pub fn validate_package(bytes: &[u8]) -> Result<(), PatchError> {
    let mut store = PartStore::open(bytes)?;
    let root_rels = store.read("_rels/.rels")?;
    let workbook_path = super::find_rel_target(&root_rels, "/officeDocument", "")
        .unwrap_or_else(|| "xl/workbook.xml".to_string());
    let (wb_dir, wb_file) = super::split_part_path(&workbook_path);
    let wb_rels_path = format!("{wb_dir}_rels/{wb_file}.rels");

    // Zip DIRECTORY entries ("xl/", trailing slash) are not OPC parts and
    // need no content type — Excel ignores them. ExcelJS emits them on
    // every write (its internal JSZip creates folder objects), so without
    // this filter every full export fails validation with
    // "new part xl/ has no content type".
    let names: Vec<String> = store
        .names()
        .into_iter()
        .filter(|n| !n.ends_with('/'))
        .collect();
    for name in &names {
        let bytes = store.read(name)?;
        validate_part(name, &bytes, &workbook_path)?;
    }
    validate_content_types(&mut store, &names)?;
    validate_workbook_rels(&mut store, &workbook_path, &wb_rels_path)?;
    Ok(())
}

fn validate_part(name: &str, bytes: &[u8], workbook_path: &str) -> Result<(), PatchError> {
    if name == workbook_path {
        validate_workbook(name, bytes)
    } else if is_worksheet_part(name) {
        validate_worksheet(name, bytes)
    } else if name.ends_with(".rels") {
        validate_rels(name, bytes)
    } else if name.ends_with(".xml") {
        // Includes [Content_Types].xml — its package-level integrity is
        // checked separately with full context.
        well_formed(name, bytes)
    } else {
        Ok(())
    }
}

fn is_worksheet_part(name: &str) -> bool {
    name.contains("worksheets/") && name.ends_with(".xml") && !name.contains("/_rels/")
}

/// Full strict parse — catches unbalanced tags and malformed markup.
fn well_formed(part: &str, xml: &[u8]) -> Result<(), PatchError> {
    let mut r = strict_reader(xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => {}
            Err(e) => return Err(err(part, format!("malformed XML: {e}"))),
        }
        buf.clear();
    }
}

/// Track schema-sequence position and singleton constraints for the
/// top-level children of a root element.
struct SequenceCheck<'a> {
    part: &'a str,
    order: &'static [&'static [u8]],
    repeatable: &'static [&'static [u8]],
    last: Option<(usize, Vec<u8>)>,
    seen: HashSet<Vec<u8>>,
}

impl<'a> SequenceCheck<'a> {
    fn new(
        part: &'a str,
        order: &'static [&'static [u8]],
        repeatable: &'static [&'static [u8]],
    ) -> Self {
        Self { part, order, repeatable, last: None, seen: HashSet::new() }
    }

    fn child(&mut self, name: &[u8]) -> Result<(), PatchError> {
        let Some(idx) = self.order.iter().position(|n| *n == name) else {
            // Unknown to our table (extensions, mc:AlternateContent) — no
            // ordering claim.
            return Ok(());
        };
        if let Some((last_idx, ref last_name)) = self.last {
            if idx < last_idx {
                return Err(err(
                    self.part,
                    format!(
                        "<{}> must come before <{}> (schema sequence)",
                        String::from_utf8_lossy(name),
                        String::from_utf8_lossy(last_name),
                    ),
                ));
            }
        }
        if !self.repeatable.contains(&name) && !self.seen.insert(name.to_vec()) {
            return Err(err(
                self.part,
                format!("duplicate <{}> element", String::from_utf8_lossy(name)),
            ));
        }
        self.last = Some((idx, name.to_vec()));
        Ok(())
    }
}

/// cfRule types that must NOT carry an `operator` attribute (Excel repairs
/// files where they do — the ExcelJS containsErrors bug class).
const CFRULE_NO_OPERATOR: &[&str] =
    &["containsBlanks", "notContainsBlanks", "containsErrors", "notContainsErrors"];

pub(super) fn validate_worksheet(part: &str, xml: &[u8]) -> Result<(), PatchError> {
    let mut r = strict_reader(xml);
    let mut buf = Vec::new();
    let mut seq = SequenceCheck::new(part, WORKSHEET_CHILD_ORDER, WORKSHEET_REPEATABLE);
    let mut depth = 0usize;
    let mut in_sheet_data = false;
    // Monotonic row/cell reference tracking (Excel requires ascending order).
    let mut next_row: u32 = 0; // implicit index if r is absent
    let mut last_row_attr: Option<u32> = None;
    let mut cur_row: Option<u32> = None;
    let mut last_col: Option<u32> = None;

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| err(part, format!("malformed XML: {e}")))?;
        let is_start = matches!(ev, Event::Start(_));
        match ev {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                if depth == 1 {
                    seq.child(&name)?;
                    if name == b"sheetData" && is_start {
                        in_sheet_data = true;
                        next_row = 0;
                        last_row_attr = None;
                    }
                }
                if in_sheet_data && depth == 2 && name == b"row" {
                    let row = match attr_value(e, b"r") {
                        Some(v) => {
                            let n = v.parse::<u32>().map_err(|_| {
                                err(part, format!("row with non-numeric r=\"{v}\""))
                            })?;
                            if n == 0 {
                                return Err(err(part, "row with r=\"0\" (rows are 1-based)"));
                            }
                            n - 1
                        }
                        None => next_row,
                    };
                    if let Some(prev) = last_row_attr {
                        if row <= prev {
                            return Err(err(
                                part,
                                format!("row {} out of order after row {}", row + 1, prev + 1),
                            ));
                        }
                    }
                    last_row_attr = Some(row);
                    next_row = row + 1;
                    cur_row = Some(row);
                    last_col = None;
                }
                if in_sheet_data && depth == 3 && name == b"c" {
                    if let Some(rc) = attr_value(e, b"r") {
                        let Some((rr, cc)) = parse_cell_ref(&rc) else {
                            return Err(err(part, format!("cell with bad r=\"{rc}\"")));
                        };
                        if let Some(row) = cur_row {
                            if rr != row {
                                return Err(err(
                                    part,
                                    format!("cell {rc} inside row {}", row + 1),
                                ));
                            }
                        }
                        if let Some(prev) = last_col {
                            if cc <= prev {
                                return Err(err(
                                    part,
                                    format!("cell {rc} out of order within its row"),
                                ));
                            }
                        }
                        last_col = Some(cc);
                    }
                }
                if name == b"cfRule" {
                    if let (Some(ty), Some(_)) =
                        (attr_value(e, b"type"), attr_value(e, b"operator"))
                    {
                        if CFRULE_NO_OPERATOR.contains(&ty.as_str()) {
                            return Err(err(
                                part,
                                format!("cfRule type=\"{ty}\" must not carry an operator"),
                            ));
                        }
                    }
                }
                if is_start {
                    depth += 1;
                }
            }
            Event::End(ref e) => {
                depth = depth.saturating_sub(1);
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name == b"sheetData" && depth == 1 {
                    in_sheet_data = false;
                }
                if name == b"row" && depth == 2 {
                    cur_row = None;
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

pub(super) fn validate_workbook(part: &str, xml: &[u8]) -> Result<(), PatchError> {
    let mut r = strict_reader(xml);
    let mut buf = Vec::new();
    let mut seq = SequenceCheck::new(part, WORKBOOK_CHILD_ORDER, WORKBOOK_REPEATABLE);
    let mut depth = 0usize;
    let mut sheet_names: HashSet<String> = HashSet::new();
    let mut sheet_ids: HashSet<String> = HashSet::new();
    let mut sheet_count = 0usize;
    // Non-empty check for <definedName>: text must appear before the end tag.
    let mut defined_name: Option<(String, bool)> = None;

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| err(part, format!("malformed XML: {e}")))?;
        let is_start = matches!(ev, Event::Start(_));
        match ev {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                if depth == 1 {
                    seq.child(&name)?;
                }
                if name == b"sheet" {
                    sheet_count += 1;
                    match attr_value(e, b"name") {
                        Some(n) if !n.is_empty() => {
                            if !sheet_names.insert(n.to_lowercase()) {
                                return Err(err(part, format!("duplicate sheet name {n:?}")));
                            }
                        }
                        _ => return Err(err(part, "sheet without a name")),
                    }
                    match attr_value(e, b"sheetId") {
                        Some(id) => {
                            if !sheet_ids.insert(id.clone()) {
                                return Err(err(part, format!("duplicate sheetId {id}")));
                            }
                        }
                        None => return Err(err(part, "sheet without sheetId")),
                    }
                    if attr_value(e, b"id").map_or(true, |v| v.is_empty()) {
                        return Err(err(part, "sheet without r:id"));
                    }
                }
                if name == b"definedName" {
                    let dn = attr_value(e, b"name").unwrap_or_default();
                    if !is_start {
                        // Self-closing definedName has no content — the
                        // ExcelJS gutted-constant bug class.
                        return Err(err(part, format!("empty definedName {dn:?}")));
                    }
                    defined_name = Some((dn, false));
                }
                if is_start {
                    depth += 1;
                }
            }
            Event::Text(ref t) => {
                if let Some((_, has_text)) = defined_name.as_mut() {
                    if !t.iter().all(|b| b.is_ascii_whitespace()) {
                        *has_text = true;
                    }
                }
            }
            Event::End(ref e) => {
                depth = depth.saturating_sub(1);
                if local_name(e.name().as_ref()) == b"definedName" {
                    if let Some((dn, has_text)) = defined_name.take() {
                        if !has_text {
                            return Err(err(part, format!("empty definedName {dn:?}")));
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    if sheet_count == 0 {
        return Err(err(part, "workbook has no sheets"));
    }
    Ok(())
}

pub(super) fn validate_rels(part: &str, xml: &[u8]) -> Result<(), PatchError> {
    let mut r = strict_reader(xml);
    let mut buf = Vec::new();
    let mut ids: HashSet<String> = HashSet::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                match attr_value(e, b"Id") {
                    Some(id) if !id.is_empty() => {
                        if !ids.insert(id.clone()) {
                            return Err(err(part, format!("duplicate relationship Id {id}")));
                        }
                    }
                    _ => return Err(err(part, "relationship without Id")),
                }
                if attr_value(e, b"Target").map_or(true, |t| t.is_empty()) {
                    return Err(err(part, "relationship without Target"));
                }
            }
            Ok(Event::Eof) => return Ok(()),
            Ok(_) => {}
            Err(e) => return Err(err(part, format!("malformed XML: {e}"))),
        }
        buf.clear();
    }
}

/// Package-level: every Override points at a real part, no duplicate
/// PartNames, and every part in `must_cover` is covered by a Default or an
/// Override (worksheets need their explicit Override — Excel won't accept
/// the xml Default for a sheet). For surgical saves `must_cover` is just
/// the added parts — pre-existing parts are raw copies of a file Excel
/// already accepted, and the validator never demands more of the output
/// than the input satisfied. For full-export validation it's every part.
fn validate_content_types(
    store: &mut PartStore,
    must_cover: &[String],
) -> Result<(), PatchError> {
    let part = "[Content_Types].xml";
    let xml = store.read(part)?;
    let names: HashSet<String> = store.names().into_iter().collect();

    let mut defaults: HashSet<String> = HashSet::new();
    let mut overrides: HashSet<String> = HashSet::new();
    let mut r = strict_reader(&xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match local_name(e.name().as_ref()) {
                    b"Default" => {
                        if let Some(ext) = attr_value(e, b"Extension") {
                            defaults.insert(ext.to_lowercase());
                        }
                    }
                    b"Override" => {
                        let Some(pn) = attr_value(e, b"PartName") else {
                            return Err(err(part, "Override without PartName"));
                        };
                        if !overrides.insert(pn.clone()) {
                            return Err(err(part, format!("duplicate Override for {pn}")));
                        }
                        let target = pn.trim_start_matches('/');
                        if !names.contains(target) {
                            return Err(err(
                                part,
                                format!("Override for missing part {pn}"),
                            ));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(err(part, format!("malformed XML: {e}"))),
        }
        buf.clear();
    }

    for name in must_cover {
        if name == part || name.starts_with("_rels/") || name.contains("/_rels/") {
            continue; // rels are covered by the rels Default in any valid base
        }
        let covered = overrides.contains(&format!("/{name}"))
            || name
                .rsplit('.')
                .next()
                .map_or(false, |ext| defaults.contains(&ext.to_lowercase()));
        if !covered {
            return Err(err(part, format!("new part {name} has no content type")));
        }
        if is_worksheet_part(name) && !overrides.contains(&format!("/{name}")) {
            return Err(err(part, format!("worksheet {name} has no Override")));
        }
    }
    Ok(())
}

/// Package-level: every workbook relationship resolves to a real part, and
/// every sheet's r:id resolves to a relationship.
fn validate_workbook_rels(
    store: &mut PartStore,
    workbook_path: &str,
    wb_rels_path: &str,
) -> Result<(), PatchError> {
    let rels_xml = store.read(wb_rels_path)?;
    let (wb_dir, _) = match workbook_path.rfind('/') {
        Some(i) => (&workbook_path[..=i], &workbook_path[i + 1..]),
        None => ("", workbook_path),
    };

    let mut targets: HashMap<String, String> = HashMap::new(); // Id → resolved target
    let mut r = strict_reader(&rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                let external = attr_value(e, b"TargetMode").as_deref() == Some("External");
                if let (Some(id), Some(target)) =
                    (attr_value(e, b"Id"), attr_value(e, b"Target"))
                {
                    if !external {
                        let resolved = if let Some(s) = target.strip_prefix('/') {
                            s.to_string()
                        } else {
                            normalize_path(&format!("{wb_dir}{target}"))
                        };
                        targets.insert(id, resolved);
                    } else {
                        targets.insert(id, String::new());
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(err(wb_rels_path, format!("malformed XML: {e}"))),
        }
        buf.clear();
    }

    for (id, target) in &targets {
        if !target.is_empty() && !store.exists(target) {
            return Err(err(
                wb_rels_path,
                format!("relationship {id} points at missing part {target}"),
            ));
        }
    }

    // Every <sheet r:id> must resolve.
    let wb_xml = store.read(workbook_path)?;
    let mut r = strict_reader(&wb_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if local_name(e.name().as_ref()) == b"sheet" =>
            {
                if let Some(rid) = attr_value(e, b"id") {
                    if !targets.contains_key(&rid) {
                        let name = attr_value(e, b"name").unwrap_or_default();
                        return Err(err(
                            workbook_path,
                            format!("sheet {name:?} references missing relationship {rid}"),
                        ));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(err(workbook_path, format!("malformed XML: {e}"))),
        }
        buf.clear();
    }
    Ok(())
}

/// Collapse "xl/../foo" and "xl/./foo" segments.
fn normalize_path(path: &str) -> String {
    let mut out: Vec<&str> = Vec::new();
    for seg in path.split('/') {
        match seg {
            "." | "" => {}
            ".." => {
                out.pop();
            }
            s => out.push(s),
        }
    }
    out.join("/")
}
