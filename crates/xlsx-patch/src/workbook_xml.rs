//! xl/workbook.xml, its rels, and [Content_Types].xml edits.
//!
//! Everything here is a streaming rewrite: events are copied through
//! verbatim except for the specific elements being changed. These parts are
//! small, so re-serialization noise (attribute quoting) is confined to
//! elements we actually touch.

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use super::patch::DefinedName;
use super::PatchError;

pub(super) fn reader(xml: &[u8]) -> Reader<&[u8]> {
    let mut r = Reader::from_reader(xml);
    r.config_mut().trim_text(false);
    r.config_mut().check_end_names = false;
    r
}

pub(super) fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().rposition(|&b| b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

pub(super) fn attr_value(e: &BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().with_checks(false).flatten() {
        if attr.key.as_ref() == key || local_name(attr.key.as_ref()) == key {
            return Some(attr.unescape_value().ok()?.into_owned());
        }
    }
    None
}

pub(super) fn xml_err(e: quick_xml::Error) -> PatchError {
    PatchError::Xml(e.to_string())
}

/// Text event with minimal escaping (only `&`, `<`, `>`): formulas are full
/// of apostrophes and quotes and `BytesText::new` would turn them into
/// `&apos;`/`&quot;` noise.
pub(super) fn text_minimal(s: &str) -> BytesText<'static> {
    let escaped = s
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");
    BytesText::from_escaped(escaped)
}

/// Map sheet display names to their zip part paths, via workbook.xml
/// (name → r:id) joined with workbook.xml.rels (r:id → target).
pub fn sheet_part_paths(
    workbook_xml: &[u8],
    rels_xml: &[u8],
) -> Result<Vec<(String, String)>, PatchError> {
    let mut rid_to_target: std::collections::HashMap<String, String> = Default::default();
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                if let (Some(id), Some(target)) = (attr_value(&e, b"Id"), attr_value(&e, b"Target"))
                {
                    rid_to_target.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
        buf.clear();
    }

    let mut out = Vec::new();
    let mut r = reader(workbook_xml);
    buf.clear();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"sheet" =>
            {
                let name = attr_value(&e, b"name");
                let rid = attr_value(&e, b"id"); // r:id — matched by local name
                if let (Some(name), Some(rid)) = (name, rid) {
                    if let Some(target) = rid_to_target.get(&rid) {
                        out.push((name, resolve_target(target)));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("workbook: {e}"))),
        }
        buf.clear();
    }
    Ok(out)
}

/// Rel targets are relative to xl/ ("worksheets/sheet5.xml") or absolute
/// part names ("/xl/worksheets/sheet5.xml").
fn resolve_target(target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        stripped.to_string()
    } else {
        format!("xl/{target}")
    }
}

/// Elements that must come AFTER calcPr in CT_Workbook's sequence — used to
/// find the insertion point when the file has no calcPr/definedNames yet.
const AFTER_CALC_PR: &[&[u8]] = &[
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

struct WorkbookRewrite<'p> {
    pending_names: Vec<&'p DefinedName>,
    wrote_defined: bool,
    wrote_calc: bool,
    set_full_calc: bool,
}

impl<'p> WorkbookRewrite<'p> {
    /// Emit any not-yet-written definedNames element and calcPr. Called when
    /// we reach a spot in the document past which they may no longer appear.
    fn flush_inserts(&mut self, w: &mut Writer<Vec<u8>>) -> Result<(), PatchError> {
        if !self.wrote_defined {
            if !self.pending_names.is_empty() {
                w.write_event(Event::Start(BytesStart::new("definedNames")))
                    .map_err(xml_err)?;
                for dn in self.pending_names.drain(..) {
                    write_defined_name(w, dn)?;
                }
                w.write_event(Event::End(BytesEnd::new("definedNames")))
                    .map_err(xml_err)?;
            }
            self.wrote_defined = true;
        }
        if !self.wrote_calc {
            let mut calc = BytesStart::new("calcPr");
            calc.push_attribute(("fullCalcOnLoad", "1"));
            w.write_event(Event::Empty(calc)).map_err(xml_err)?;
            self.wrote_calc = true;
        }
        Ok(())
    }

    fn needs_insert_before(&self, name: &[u8]) -> bool {
        (!self.wrote_defined || !self.wrote_calc) && AFTER_CALC_PR.contains(&name)
    }
}

/// Rewrite workbook.xml: optionally force `calcPr/@fullCalcOnLoad="1"` and
/// create-or-replace defined names. Order per CT_Workbook: definedNames
/// immediately precedes calcPr.
pub fn patch_workbook_xml(
    xml: &[u8],
    set_full_calc: bool,
    defined_names: &[DefinedName],
) -> Result<Vec<u8>, PatchError> {
    let mut st = WorkbookRewrite {
        pending_names: defined_names.iter().collect(),
        wrote_defined: defined_names.is_empty(),
        wrote_calc: !set_full_calc,
        set_full_calc,
    };

    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut depth = 0usize;
    let mut in_defined_names = false;
    // When true we're inside a definedName whose ref we replaced: swallow its
    // original text nodes until the matching end tag.
    let mut swallowing_ref_text = false;

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("workbook: {e}")))?;
        match ev {
            Event::Eof => break,

            Event::Start(e) if depth == 1 && local_name(e.name().as_ref()) == b"definedNames" => {
                in_defined_names = true;
                st.wrote_defined = true;
                depth += 1;
                w.write_event(Event::Start(e)).map_err(xml_err)?;
            }
            Event::End(e)
                if in_defined_names && local_name(e.name().as_ref()) == b"definedNames" =>
            {
                for dn in st.pending_names.drain(..) {
                    write_defined_name(&mut w, dn)?;
                }
                in_defined_names = false;
                depth = depth.saturating_sub(1);
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            Event::Start(e)
                if in_defined_names && local_name(e.name().as_ref()) == b"definedName" =>
            {
                depth += 1;
                let name = attr_value(&e, b"name");
                let ours = name
                    .as_deref()
                    .and_then(|n| st.pending_names.iter().position(|d| d.name == n));
                w.write_event(Event::Start(e)).map_err(xml_err)?;
                if let Some(idx) = ours {
                    let dn = st.pending_names.remove(idx);
                    w.write_event(Event::Text(BytesText::new(&dn.r#ref)))
                        .map_err(xml_err)?;
                    swallowing_ref_text = true;
                }
            }
            Event::Text(_) if swallowing_ref_text => {}
            Event::End(e)
                if swallowing_ref_text && local_name(e.name().as_ref()) == b"definedName" =>
            {
                swallowing_ref_text = false;
                depth = depth.saturating_sub(1);
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }

            Event::Empty(e) if depth == 1 && local_name(e.name().as_ref()) == b"calcPr" => {
                // definedNames precedes calcPr; emit missing ones now.
                let was_calc_pending = !st.wrote_calc;
                st.wrote_calc = true; // don't let flush_inserts emit a second calcPr
                st.flush_inserts(&mut w)?;
                if st.set_full_calc && was_calc_pending {
                    w.write_event(Event::Empty(with_attr(&e, "fullCalcOnLoad", "1")))
                        .map_err(xml_err)?;
                } else {
                    w.write_event(Event::Empty(e)).map_err(xml_err)?;
                }
            }
            Event::Start(e) if depth == 1 && local_name(e.name().as_ref()) == b"calcPr" => {
                let was_calc_pending = !st.wrote_calc;
                st.wrote_calc = true;
                st.flush_inserts(&mut w)?;
                depth += 1;
                if st.set_full_calc && was_calc_pending {
                    w.write_event(Event::Start(with_attr(&e, "fullCalcOnLoad", "1")))
                        .map_err(xml_err)?;
                } else {
                    w.write_event(Event::Start(e)).map_err(xml_err)?;
                }
            }

            Event::Start(e) => {
                if depth == 1 && st.needs_insert_before(local_name(e.name().as_ref())) {
                    st.flush_inserts(&mut w)?;
                }
                depth += 1;
                w.write_event(Event::Start(e)).map_err(xml_err)?;
            }
            Event::Empty(e) => {
                if depth == 1 && st.needs_insert_before(local_name(e.name().as_ref())) {
                    st.flush_inserts(&mut w)?;
                }
                w.write_event(Event::Empty(e)).map_err(xml_err)?;
            }
            Event::End(e) => {
                if depth == 1 && local_name(e.name().as_ref()) == b"workbook" {
                    st.flush_inserts(&mut w)?;
                }
                depth = depth.saturating_sub(1);
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            other => {
                w.write_event(other).map_err(xml_err)?;
            }
        }
        buf.clear();
    }

    Ok(w.into_inner())
}

fn write_defined_name(w: &mut Writer<Vec<u8>>, dn: &DefinedName) -> Result<(), PatchError> {
    let mut e = BytesStart::new("definedName");
    e.push_attribute(("name", dn.name.as_str()));
    w.write_event(Event::Start(e)).map_err(xml_err)?;
    w.write_event(Event::Text(BytesText::new(&dn.r#ref)))
        .map_err(xml_err)?;
    w.write_event(Event::End(BytesEnd::new("definedName")))
        .map_err(xml_err)?;
    Ok(())
}

/// Copy of `e` with `key` set to `value` (replacing any existing value).
pub(super) fn with_attr(e: &BytesStart, key: &str, value: &str) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    let mut out = BytesStart::new(name);
    for attr in e.attributes().with_checks(false).flatten() {
        if attr.key.as_ref() != key.as_bytes() {
            out.push_attribute(attr);
        }
    }
    out.push_attribute((key, value));
    out
}

/// Drop the workbook.xml.rels Relationship whose Type ends with /calcChain.
pub fn strip_calc_chain_rel(rels_xml: &[u8]) -> Result<Vec<u8>, PatchError> {
    filter_elements(rels_xml, |e| {
        local_name(e.name().as_ref()) == b"Relationship"
            && attr_value(e, b"Type").is_some_and(|t| t.ends_with("/calcChain"))
    })
}

/// Drop the [Content_Types].xml Override for the given part name.
pub fn strip_content_type_override(ct_xml: &[u8], part_name: &str) -> Result<Vec<u8>, PatchError> {
    filter_elements(ct_xml, |e| {
        local_name(e.name().as_ref()) == b"Override"
            && attr_value(e, b"PartName").is_some_and(|p| p == part_name)
    })
}

/// Stream-copy `xml`, dropping every element (Start..End or Empty) matching
/// `drop`. Only sensible for elements whose children should go with them.
fn filter_elements(xml: &[u8], drop: impl Fn(&BytesStart) -> bool) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut skipping_depth: Option<usize> = None;
    let mut depth = 0usize;
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("filter: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Empty(e) => {
                if skipping_depth.is_none() && !drop(&e) {
                    w.write_event(Event::Empty(e)).map_err(xml_err)?;
                }
            }
            Event::Start(e) => {
                depth += 1;
                if skipping_depth.is_none() && drop(&e) {
                    skipping_depth = Some(depth);
                } else if skipping_depth.is_none() {
                    w.write_event(Event::Start(e)).map_err(xml_err)?;
                }
            }
            Event::End(e) => {
                if skipping_depth == Some(depth) {
                    skipping_depth = None;
                } else if skipping_depth.is_none() {
                    w.write_event(Event::End(e)).map_err(xml_err)?;
                }
                depth = depth.saturating_sub(1);
            }
            other => {
                if skipping_depth.is_none() {
                    w.write_event(other).map_err(xml_err)?;
                }
            }
        }
        buf.clear();
    }
    Ok(w.into_inner())
}
