//! Row/column insert-delete: package-wide positional reference shifting.
//!
//! One operation touches, at most:
//!   - the target worksheet (rows/cells renumbered, ranges adjusted,
//!     shared formulas materialized, formulas positionally adjusted),
//!   - other worksheets whose formulas reference the target sheet,
//!   - workbook.xml defined names,
//!   - chart parts (series `c:f` formulas),
//!   - the target sheet's comments, VML anchors, and drawing anchors.
//!
//! Everything else stays byte-identical. When the package contains carriers
//! we can't rewrite (pivots, tables, slicers, OLE objects, data tables) the
//! op fails closed and the caller falls back to the visible copy flow.

use std::collections::HashMap;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use super::package::PartStore;
use super::patch::RowColOp;
use super::refs::{
    adjust_formula, adjust_sqref_list, format_cell_ref, parse_cell_ref, shift_formula,
    RowColShift, MAX_COL, MAX_ROW,
};
use super::sheet_xml::scan_sheet;
use super::structure::{mentions_name, preflight_structure, rewrite_formula_texts};
use super::workbook_xml::{self, attr_value, local_name, reader, text_minimal, xml_err};
use super::PatchError;

pub fn op_shift(op: &RowColOp) -> (&str, RowColShift) {
    match op {
        RowColOp::InsertRows { sheet, before, count } => (
            sheet.as_str(),
            RowColShift { rows: true, start: *before, count: *count, insert: true },
        ),
        RowColOp::DeleteRows { sheet, start, count } => (
            sheet.as_str(),
            RowColShift { rows: true, start: *start, count: *count, insert: false },
        ),
        RowColOp::InsertColumns { sheet, before, count } => (
            sheet.as_str(),
            RowColShift { rows: false, start: *before, count: *count, insert: true },
        ),
        RowColOp::DeleteColumns { sheet, start, count } => (
            sheet.as_str(),
            RowColShift { rows: false, start: *start, count: *count, insert: false },
        ),
    }
}

fn names_eq(a: &str, b: &str) -> bool {
    a.len() == b.len()
        && a.chars()
            .zip(b.chars())
            .all(|(x, y)| x.eq_ignore_ascii_case(&y))
}

pub fn apply_row_col_op(
    store: &mut PartStore,
    workbook_path: &str,
    wb_dir: &str,
    wb_rels_path: &str,
    op: &RowColOp,
) -> Result<(), PatchError> {
    let (sheet_name, shift) = op_shift(op);
    if shift.count == 0 {
        return Ok(());
    }
    preflight_structure(store)?;

    let workbook_xml = store.read(workbook_path)?;
    let wb_rels = store.read(wb_rels_path)?;
    let sheet_paths = workbook_xml::sheet_part_paths(&workbook_xml, &wb_rels)?;
    let target_part = sheet_paths
        .iter()
        .find(|(n, _)| names_eq(n, sheet_name))
        .map(|(_, p)| p.clone())
        .ok_or_else(|| PatchError::SheetNotFound(sheet_name.to_string()))?;

    // --- 1. the target worksheet ---
    let bytes = store.read(&target_part)?;
    let out = shift_sheet(&bytes, &shift, sheet_name, true)?;
    store.write(&target_part, out.xml);

    // --- 2. other worksheets referencing the target sheet ---
    for (name, part) in &sheet_paths {
        if part == &target_part || names_eq(name, sheet_name) {
            continue;
        }
        let bytes = store.read(part)?;
        if !mentions_name(&bytes, sheet_name) {
            continue;
        }
        let out = shift_sheet(&bytes, &shift, sheet_name, false)?;
        if out.dirty {
            store.write(part, out.xml);
        }
    }

    // --- 3. workbook.xml defined names ---
    if mentions_name(&workbook_xml, sheet_name) {
        let mut dirty = false;
        let out = rewrite_formula_texts(&workbook_xml, &[b"definedName"], &mut |f| {
            let new = adjust_formula(f, &shift, sheet_name, false)?;
            if new != f {
                dirty = true;
            }
            Ok(new)
        })?;
        if dirty {
            store.write(workbook_path, out);
        }
    }

    // --- 4. chart parts (series formulas reference sheets by name) ---
    let chart_prefix = format!("{wb_dir}charts/");
    for part in store.names() {
        if !part.starts_with(&chart_prefix) || !part.ends_with(".xml") || part.contains("/_rels/")
        {
            continue;
        }
        let bytes = store.read(&part)?;
        if !mentions_name(&bytes, sheet_name) {
            continue;
        }
        let mut dirty = false;
        let out = rewrite_formula_texts(&bytes, &[b"f"], &mut |f| {
            let new = adjust_formula(f, &shift, sheet_name, false)?;
            if new != f {
                dirty = true;
            }
            Ok(new)
        })?;
        if dirty {
            store.write(&part, out);
        }
    }

    // --- 5. the target sheet's comments / VML / drawing anchors ---
    let (part_dir, part_file) = match target_part.rfind('/') {
        Some(i) => (&target_part[..=i], &target_part[i + 1..]),
        None => ("", target_part.as_str()),
    };
    let rels_path = format!("{part_dir}_rels/{part_file}.rels");
    if store.exists(&rels_path) {
        let rels = store.read(&rels_path)?;
        for (rtype, target) in list_relationships(&rels)? {
            let resolved = resolve_rel_target(part_dir, &target);
            if rtype.ends_with("/comments") || rtype.ends_with("/threadedComment") {
                let bytes = store.read(&resolved)?;
                store.write(&resolved, shift_comments(&bytes, &shift)?);
            } else if rtype.ends_with("/vmlDrawing") {
                let bytes = store.read(&resolved)?;
                store.write(&resolved, shift_vml(&bytes, &shift)?);
            } else if rtype.ends_with("/drawing") {
                ensure_unshared_part(store, &sheet_paths, &target_part, &resolved)?;
                let bytes = store.read(&resolved)?;
                store.write(&resolved, shift_drawing(&bytes, &shift)?);
            }
        }
    }

    Ok(())
}

// ---------------------------------------------------------------------------
// rels helpers
// ---------------------------------------------------------------------------

fn list_relationships(rels_xml: &[u8]) -> Result<Vec<(String, String)>, PatchError> {
    let mut out = Vec::new();
    let mut r = reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(e)) | Ok(Event::Start(e))
                if local_name(e.name().as_ref()) == b"Relationship" =>
            {
                if let (Some(ty), Some(target)) =
                    (attr_value(&e, b"Type"), attr_value(&e, b"Target"))
                {
                    out.push((ty, target));
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("rels: {e}"))),
        }
        buf.clear();
    }
    Ok(out)
}

/// Resolve a rel target ("../drawings/drawing1.xml") against the source
/// part's directory ("xl/worksheets/").
pub fn resolve_rel_target(base_dir: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return stripped.to_string();
    }
    let mut segments: Vec<&str> = base_dir.split('/').filter(|s| !s.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            s => segments.push(s),
        }
    }
    segments.join("/")
}

/// A drawing part referenced by more than one sheet can't be shifted for
/// just one of them. Fail closed (unseen in practice, but cheap to check).
fn ensure_unshared_part(
    store: &mut PartStore,
    sheet_paths: &[(String, String)],
    target_part: &str,
    part: &str,
) -> Result<(), PatchError> {
    for (_, sheet_part) in sheet_paths {
        if sheet_part == target_part {
            continue;
        }
        let (dir, file) = match sheet_part.rfind('/') {
            Some(i) => (&sheet_part[..=i], &sheet_part[i + 1..]),
            None => ("", sheet_part.as_str()),
        };
        let rels_path = format!("{dir}_rels/{file}.rels");
        if !store.exists(&rels_path) {
            continue;
        }
        let rels = store.read(&rels_path)?;
        for (_, target) in list_relationships(&rels)? {
            if resolve_rel_target(dir, &target) == part {
                return Err(PatchError::Unsupported(format!(
                    "drawing part {part} is shared between sheets"
                )));
            }
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// worksheet transform
// ---------------------------------------------------------------------------

pub struct ShiftOutput {
    pub xml: Vec<u8>,
    pub dirty: bool,
}

/// Counts and guards that must be known before streaming (count attributes
/// precede their children; unsupported carriers must fail before output).
#[derive(Default)]
struct PreScan {
    merge_total: usize,
    merge_dropped: usize,
    dv_total: usize,
    dv_dropped: usize,
    cols_total: usize,
    cols_dropped: usize,
}

fn prescan(
    xml: &[u8],
    shift: &RowColShift,
    existing_merges: &[String],
) -> Result<PreScan, PatchError> {
    let mut out = PreScan {
        merge_total: existing_merges.len(),
        ..Default::default()
    };
    for m in existing_merges {
        if adjust_sqref_list(m, shift)?.is_none() {
            out.merge_dropped += 1;
        }
    }
    let mut r = reader(xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                match local_name(e.name().as_ref()) {
                    b"oleObjects" | b"controls" => {
                        return Err(PatchError::Unsupported(
                            "sheet contains OLE objects or form controls".into(),
                        ));
                    }
                    b"dataValidation" => {
                        out.dv_total += 1;
                        if let Some(sqref) = attr_value(e, b"sqref") {
                            if adjust_sqref_list(&sqref, shift)?.is_none() {
                                out.dv_dropped += 1;
                            }
                        }
                    }
                    b"col" if !shift.rows => {
                        out.cols_total += 1;
                        let min = attr_value(e, b"min").and_then(|v| v.parse::<u32>().ok());
                        let max = attr_value(e, b"max").and_then(|v| v.parse::<u32>().ok());
                        if let (Some(min), Some(max)) = (min, max) {
                            if min >= 1 && shift.adjust_span(min - 1, max - 1).is_none() {
                                out.cols_dropped += 1;
                            }
                        }
                    }
                    b"f" => {
                        if attr_value(e, b"t").as_deref() == Some("dataTable") {
                            return Err(PatchError::Unsupported(
                                "sheet contains a what-if data table".into(),
                            ));
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(PatchError::Xml(format!("prescan: {e}"))),
        }
        buf.clear();
    }
    Ok(out)
}

/// Copy of `e` with attribute updates applied: `Some(v)` sets, `None`
/// removes. Untouched attributes keep their original order and value.
fn set_attrs(e: &BytesStart, updates: &[(&str, Option<&str>)]) -> BytesStart<'static> {
    let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    let mut out = BytesStart::new(name);
    let mut written: Vec<&str> = Vec::new();
    for attr in e.attributes().with_checks(false).flatten() {
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        if let Some((k, v)) = updates.iter().find(|(k, _)| k.as_bytes() == key.as_bytes()) {
            match v {
                Some(v) => {
                    out.push_attribute((*k, *v));
                    written.push(k);
                }
                None => {} // dropped
            }
        } else {
            out.push_attribute(attr);
        }
    }
    for (k, v) in updates {
        if let Some(v) = v {
            if !written.contains(k) {
                out.push_attribute((*k, *v));
            }
        }
    }
    out
}

/// Clamp-adjust a single cell ref ("B7"); unparseable input passes through.
fn clamp_cell(a1: &str, shift: &RowColShift) -> String {
    match parse_cell_ref(a1) {
        Some((r, c)) => {
            if shift.rows {
                format_cell_ref(shift.clamp_index(r), c)
            } else {
                format_cell_ref(r, shift.clamp_index(c))
            }
        }
        None => a1.to_string(),
    }
}

enum FMode {
    /// Plain/array formula: adjust its text.
    AdjustText,
    /// Materialized shared formula: replacement already written, swallow
    /// the original element's content and end tag.
    Swallow,
}

enum Open {
    /// Emit the (possibly rebuilt) tag.
    Emit(BytesStart<'static>),
    /// Emit the original event untouched.
    Raw,
    /// Drop the element (and its subtree when it's a Start).
    Drop,
}

/// Rewrite one worksheet for `shift`.
///
/// `is_target` = the sheet the rows/columns move on: rows/cells renumber
/// and local ranges adjust. On other sheets only formulas whose refs are
/// qualified with `target_sheet` change (plus shared-group materialization
/// when member adjustments diverge).
pub fn shift_sheet(
    xml: &[u8],
    shift: &RowColShift,
    target_sheet: &str,
    is_target: bool,
) -> Result<ShiftOutput, PatchError> {
    let scan = scan_sheet(xml, &Default::default())?;

    // Shared formulas: a group must be materialized when any member's
    // derived formula adjusts differently from "derive from adjusted
    // master" — and always on the target sheet, where member positions and
    // the group's ref attribute go stale. Materialized members keep their
    // cached <v>.
    let mut materialize: HashMap<(u32, u32), String> = HashMap::new();
    for (si, g) in &scan.shared_groups {
        if g.formula.is_empty() {
            return Err(PatchError::SharedFormulaMissing(*si));
        }
        let mut entries = Vec::new();
        let mut needed = is_target;
        for m in &g.members {
            let derived = if *m == g.master {
                g.formula.clone()
            } else {
                let dr = m.0 as i64 - g.master.0 as i64;
                let dc = m.1 as i64 - g.master.1 as i64;
                shift_formula(&g.formula, dr, dc)?
            };
            let adjusted = adjust_formula(&derived, shift, target_sheet, is_target)?;
            if adjusted != derived {
                needed = true;
            }
            entries.push((*m, adjusted));
        }
        if needed {
            materialize.extend(entries);
        }
    }

    let pre = if is_target {
        prescan(xml, shift, &scan.existing_merges)?
    } else {
        PreScan::default()
    };

    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut dirty = is_target;
    let mut depth = 0usize;
    // Local names of open Start elements (parallel to depth).
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut skip_until: Option<usize> = None;

    let mut in_sheet_data = false;
    let mut cur_row_old: u32 = 0; // source coordinates
    let mut cur_row_new: u32 = 0;
    let mut next_col: u32 = 0;
    let mut cur_cell_old: Option<(u32, u32)> = None;
    let mut f_mode: Option<FMode> = None;

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("shift: {e}")))?;

        if let Some(k) = skip_until {
            match ev {
                Event::Start(_) => depth += 1,
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if depth < k {
                        skip_until = None;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        let was_empty = matches!(ev, Event::Empty(_));
        match ev {
            Event::Eof => break,

            Event::Start(ref e) | Event::Empty(ref e) => {
                let name = local_name(e.name().as_ref()).to_vec();

                // --- <f> needs multi-event output; handle before Open ---
                if name == b"f" && cur_cell_old.is_some() {
                    let cell = cur_cell_old.unwrap();
                    let t = attr_value(e, b"t");
                    if t.as_deref() == Some("shared") {
                        if let Some(formula) = materialize.get(&cell) {
                            w.write_event(Event::Start(BytesStart::new("f")))
                                .map_err(xml_err)?;
                            w.write_event(Event::Text(text_minimal(formula)))
                                .map_err(xml_err)?;
                            w.write_event(Event::End(BytesEnd::new("f")))
                                .map_err(xml_err)?;
                            dirty = true;
                            if !was_empty {
                                f_mode = Some(FMode::Swallow);
                                depth += 1;
                                stack.push(name);
                            }
                            buf.clear();
                            continue;
                        }
                        // Untouched group on a non-target sheet.
                        write_open(&mut w, e, was_empty)?;
                        if !was_empty {
                            depth += 1;
                            stack.push(name);
                        }
                        buf.clear();
                        continue;
                    }
                    // Array formulas carry a range attribute of their own.
                    let mut open = e.to_owned();
                    if is_target {
                        if let Some(ref_attr) = attr_value(e, b"ref") {
                            match adjust_sqref_list(&ref_attr, shift)? {
                                Some(new_ref) => {
                                    if new_ref != ref_attr {
                                        dirty = true;
                                    }
                                    open = set_attrs(e, &[("ref", Some(&new_ref))]);
                                }
                                None => {
                                    // The anchor cell survives (we're here)
                                    // but the range died: collapse to the
                                    // cell itself.
                                    let cell_ref = format_cell_ref(
                                        if shift.rows { cur_row_new } else { cell.0 },
                                        if shift.rows {
                                            cell.1
                                        } else {
                                            shift.clamp_index(cell.1)
                                        },
                                    );
                                    open = set_attrs(e, &[("ref", Some(&cell_ref))]);
                                    dirty = true;
                                }
                            }
                        }
                    }
                    write_open(&mut w, &open, was_empty)?;
                    if !was_empty {
                        f_mode = Some(FMode::AdjustText);
                        depth += 1;
                        stack.push(name);
                    }
                    buf.clear();
                    continue;
                }

                let action = open_action(
                    e,
                    &name,
                    shift,
                    is_target,
                    &pre,
                    in_sheet_data,
                    &stack,
                    &mut cur_row_old,
                    &mut cur_row_new,
                    &mut next_col,
                    &mut cur_cell_old,
                    &mut dirty,
                )?;

                match action {
                    Open::Drop => {
                        if !was_empty {
                            skip_until = Some(depth + 1);
                            depth += 1;
                        }
                    }
                    Open::Emit(rebuilt) => {
                        write_open(&mut w, &rebuilt, was_empty)?;
                        if !was_empty {
                            depth += 1;
                            stack.push(name.clone());
                        }
                    }
                    Open::Raw => {
                        write_open(&mut w, e, was_empty)?;
                        if !was_empty {
                            depth += 1;
                            stack.push(name.clone());
                        }
                    }
                }
                if name == b"sheetData" && !was_empty {
                    in_sheet_data = true;
                }
                if name == b"c" && was_empty {
                    cur_cell_old = None;
                }
            }

            Event::Text(ref t) => {
                match f_mode {
                    Some(FMode::Swallow) => {}
                    Some(FMode::AdjustText) => {
                        let text = t
                            .unescape()
                            .map_err(|e| PatchError::Xml(format!("f text: {e}")))?;
                        let new = adjust_formula(&text, shift, target_sheet, is_target)?;
                        if new != text.as_ref() {
                            dirty = true;
                        }
                        w.write_event(Event::Text(text_minimal(&new))).map_err(xml_err)?;
                    }
                    None => {
                        let handled = adjust_text_by_context(
                            &mut w,
                            t,
                            &stack,
                            shift,
                            target_sheet,
                            is_target,
                            &mut dirty,
                        )?;
                        if !handled {
                            w.write_event(Event::Text(t.to_owned())).map_err(xml_err)?;
                        }
                    }
                }
            }

            Event::End(ref e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                depth = depth.saturating_sub(1);
                stack.pop();
                match name.as_slice() {
                    b"f" => {
                        let swallow = matches!(f_mode, Some(FMode::Swallow));
                        f_mode = None;
                        if swallow {
                            buf.clear();
                            continue; // replacement already closed
                        }
                        w.write_event(Event::End(e.to_owned())).map_err(xml_err)?;
                    }
                    b"c" => {
                        cur_cell_old = None;
                        w.write_event(Event::End(e.to_owned())).map_err(xml_err)?;
                    }
                    b"sheetData" => {
                        in_sheet_data = false;
                        w.write_event(Event::End(e.to_owned())).map_err(xml_err)?;
                    }
                    _ => {
                        w.write_event(Event::End(e.to_owned())).map_err(xml_err)?;
                    }
                }
            }

            other => {
                w.write_event(other).map_err(xml_err)?;
            }
        }
        buf.clear();
    }

    Ok(ShiftOutput {
        xml: w.into_inner(),
        dirty,
    })
}

fn write_open(
    w: &mut Writer<Vec<u8>>,
    e: &BytesStart,
    was_empty: bool,
) -> Result<(), PatchError> {
    if was_empty {
        w.write_event(Event::Empty(e.to_owned())).map_err(xml_err)?;
    } else {
        w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
    }
    Ok(())
}

/// Formula-bearing text elements outside `<f>`: CF rule formulas, DV
/// formula1/formula2, and extension-list (x14 CF, sparklines) f / sqref.
fn adjust_text_by_context(
    w: &mut Writer<Vec<u8>>,
    t: &BytesText,
    stack: &[Vec<u8>],
    shift: &RowColShift,
    target_sheet: &str,
    is_target: bool,
    dirty: &mut bool,
) -> Result<bool, PatchError> {
    let Some(top) = stack.last() else {
        return Ok(false);
    };
    let in_ext = stack.iter().any(|n| n == b"extLst");
    let is_formula_tag = top == b"formula" || top == b"formula1" || top == b"formula2";
    let is_ext_f = in_ext && top == b"f";
    let is_ext_sqref = in_ext && top == b"sqref";

    if is_formula_tag || is_ext_f {
        let text = t
            .unescape()
            .map_err(|e| PatchError::Xml(format!("formula text: {e}")))?;
        let new = adjust_formula(&text, shift, target_sheet, is_target)?;
        if new != text.as_ref() {
            *dirty = true;
        }
        w.write_event(Event::Text(text_minimal(&new))).map_err(xml_err)?;
        return Ok(true);
    }
    if is_ext_sqref && is_target {
        let text = t
            .unescape()
            .map_err(|e| PatchError::Xml(format!("sqref text: {e}")))?;
        match adjust_sqref_list(&text, shift)? {
            Some(new) => {
                if new != text.as_ref() {
                    *dirty = true;
                }
                w.write_event(Event::Text(text_minimal(&new))).map_err(xml_err)?;
                Ok(true)
            }
            None => Err(PatchError::Unsupported(
                "an extended formatting range was fully deleted".into(),
            )),
        }
    } else {
        Ok(false)
    }
}

/// Decide what to do with a non-`<f>` open tag.
#[allow(clippy::too_many_arguments)]
fn open_action(
    e: &BytesStart,
    name: &[u8],
    shift: &RowColShift,
    is_target: bool,
    pre: &PreScan,
    in_sheet_data: bool,
    stack: &[Vec<u8>],
    cur_row_old: &mut u32,
    cur_row_new: &mut u32,
    next_col: &mut u32,
    cur_cell_old: &mut Option<(u32, u32)>,
    dirty: &mut bool,
) -> Result<Open, PatchError> {
    // Cell/row tracking runs on every sheet (materialization needs coords);
    // rewriting only on the target sheet.
    match name {
        b"row" if in_sheet_data => {
            let old_r = attr_value(e, b"r")
                .and_then(|v| v.parse::<u32>().ok())
                .map(|v| v - 1)
                .unwrap_or(*cur_row_old);
            *cur_row_old = old_r + 1; // next implicit row
            *next_col = 0;
            if !is_target {
                *cur_row_new = old_r;
                return Ok(Open::Raw);
            }
            if shift.rows {
                match shift.adjust_index(old_r) {
                    None => return Ok(Open::Drop),
                    Some(nr) => {
                        if nr >= MAX_ROW {
                            return Err(PatchError::RefOutOfRange);
                        }
                        *cur_row_new = nr;
                        if nr == old_r {
                            return Ok(Open::Raw);
                        }
                        return Ok(Open::Emit(set_attrs(
                            e,
                            &[("r", Some(&(nr + 1).to_string()))],
                        )));
                    }
                }
            }
            // Column op: row number keeps, spans go stale.
            *cur_row_new = old_r;
            if attr_value(e, b"spans").is_some() {
                return Ok(Open::Emit(set_attrs(e, &[("spans", None)])));
            }
            return Ok(Open::Raw);
        }

        b"c" => {
            let rc = attr_value(e, b"r")
                .and_then(|v| parse_cell_ref(&v))
                .unwrap_or((cur_row_old.saturating_sub(1), *next_col));
            *next_col = rc.1 + 1;
            *cur_cell_old = Some(rc);
            if !is_target {
                return Ok(Open::Raw);
            }
            let new_c = if shift.rows {
                rc.1
            } else {
                match shift.adjust_index(rc.1) {
                    None => {
                        *cur_cell_old = None;
                        return Ok(Open::Drop);
                    }
                    Some(nc) => {
                        if nc >= MAX_COL {
                            return Err(PatchError::RefOutOfRange);
                        }
                        nc
                    }
                }
            };
            let new_ref = format_cell_ref(*cur_row_new, new_c);
            let cur_ref = attr_value(e, b"r");
            if cur_ref.as_deref() == Some(new_ref.as_str()) {
                return Ok(Open::Raw);
            }
            return Ok(Open::Emit(set_attrs(e, &[("r", Some(&new_ref))])));
        }

        _ => {}
    }

    if !is_target {
        return Ok(Open::Raw);
    }

    match name {
        b"dimension" => {
            if let Some(ref_attr) = attr_value(e, b"ref") {
                let new = adjust_sqref_list(&ref_attr, shift)?.unwrap_or_else(|| "A1".into());
                if new != ref_attr {
                    *dirty = true;
                    return Ok(Open::Emit(set_attrs(e, &[("ref", Some(&new))])));
                }
            }
            Ok(Open::Raw)
        }

        b"col" if !shift.rows => {
            let min = attr_value(e, b"min").and_then(|v| v.parse::<u32>().ok());
            let max = attr_value(e, b"max").and_then(|v| v.parse::<u32>().ok());
            if let (Some(min), Some(max)) = (min, max) {
                if min >= 1 {
                    return match shift.adjust_span(min - 1, max - 1) {
                        None => Ok(Open::Drop),
                        Some((a, b)) => {
                            if b >= MAX_COL {
                                return Err(PatchError::RefOutOfRange);
                            }
                            if (a, b) == (min - 1, max - 1) {
                                Ok(Open::Raw)
                            } else {
                                Ok(Open::Emit(set_attrs(
                                    e,
                                    &[
                                        ("min", Some(&(a + 1).to_string())),
                                        ("max", Some(&(b + 1).to_string())),
                                    ],
                                )))
                            }
                        }
                    };
                }
            }
            Ok(Open::Raw)
        }
        b"cols" if !shift.rows && pre.cols_total > 0 && pre.cols_total == pre.cols_dropped => {
            Ok(Open::Drop)
        }

        b"mergeCells" => {
            if pre.merge_total > 0 && pre.merge_total == pre.merge_dropped {
                return Ok(Open::Drop);
            }
            if pre.merge_dropped > 0 && attr_value(e, b"count").is_some() {
                let n = (pre.merge_total - pre.merge_dropped).to_string();
                return Ok(Open::Emit(set_attrs(e, &[("count", Some(&n))])));
            }
            Ok(Open::Raw)
        }
        b"mergeCell" => match attr_value(e, b"ref") {
            Some(ref_attr) => match adjust_sqref_list(&ref_attr, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != ref_attr => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("ref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"conditionalFormatting" => match attr_value(e, b"sqref") {
            Some(sqref) => match adjust_sqref_list(&sqref, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != sqref => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("sqref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"dataValidations" => {
            if pre.dv_total > 0 && pre.dv_total == pre.dv_dropped {
                return Ok(Open::Drop);
            }
            if pre.dv_dropped > 0 && attr_value(e, b"count").is_some() {
                let n = (pre.dv_total - pre.dv_dropped).to_string();
                return Ok(Open::Emit(set_attrs(e, &[("count", Some(&n))])));
            }
            Ok(Open::Raw)
        }
        b"dataValidation" => match attr_value(e, b"sqref") {
            Some(sqref) => match adjust_sqref_list(&sqref, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != sqref => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("sqref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"hyperlink" => match attr_value(e, b"ref") {
            Some(ref_attr) => match adjust_sqref_list(&ref_attr, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != ref_attr => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("ref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"protectedRange" | b"ignoredError" => match attr_value(e, b"sqref") {
            Some(sqref) => match adjust_sqref_list(&sqref, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != sqref => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("sqref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"cellWatch" => match attr_value(e, b"r") {
            Some(r_attr) => match adjust_sqref_list(&r_attr, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != r_attr => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("r", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"autoFilter" => {
            // Only the worksheet-level filter (tables are preflight-blocked;
            // a filter inside customSheetViews gets the same treatment).
            match attr_value(e, b"ref") {
                Some(ref_attr) => {
                    if !shift.rows {
                        // Column ops inside the filter range would shift
                        // filterColumn colId offsets — fail closed unless
                        // the range moves as a whole or is untouched.
                        if let Some((_, c1, _, c2)) = super::refs::parse_a1_range(&ref_attr) {
                            let end = shift.start + shift.count; // exclusive
                            let intersects = if shift.insert {
                                shift.start > c1 && shift.start <= c2
                            } else {
                                shift.start <= c2 && end > c1
                            };
                            if intersects {
                                return Err(PatchError::Unsupported(
                                    "column change intersects an AutoFilter range".into(),
                                ));
                            }
                        }
                    }
                    match adjust_sqref_list(&ref_attr, shift)? {
                        None => Ok(Open::Drop),
                        Some(new) if new != ref_attr => {
                            *dirty = true;
                            Ok(Open::Emit(set_attrs(e, &[("ref", Some(&new))])))
                        }
                        Some(_) => Ok(Open::Raw),
                    }
                }
                None => Ok(Open::Raw),
            }
        }
        b"sortState" | b"sortCondition" => match attr_value(e, b"ref") {
            Some(ref_attr) => match adjust_sqref_list(&ref_attr, shift)? {
                None => Ok(Open::Drop),
                Some(new) if new != ref_attr => {
                    *dirty = true;
                    Ok(Open::Emit(set_attrs(e, &[("ref", Some(&new))])))
                }
                Some(_) => Ok(Open::Raw),
            },
            None => Ok(Open::Raw),
        },

        b"brk" => {
            let parent_rows = stack.iter().rev().find_map(|n| {
                if n == b"rowBreaks" {
                    Some(true)
                } else if n == b"colBreaks" {
                    Some(false)
                } else {
                    None
                }
            });
            let Some(parent_rows) = parent_rows else {
                return Ok(Open::Raw);
            };
            let mut updates: Vec<(&str, Option<String>)> = Vec::new();
            if parent_rows == shift.rows {
                // id is the 1-based index of the row/col before the break.
                if let Some(id) = attr_value(e, b"id").and_then(|v| v.parse::<u32>().ok()) {
                    if id >= 1 {
                        let nid = shift.clamp_index(id - 1) + 1;
                        if nid != id {
                            *dirty = true;
                            updates.push(("id", Some(nid.to_string())));
                        }
                    }
                }
            } else {
                // min/max span the perpendicular axis (0-based).
                let min = attr_value(e, b"min").and_then(|v| v.parse::<u32>().ok());
                let max = attr_value(e, b"max").and_then(|v| v.parse::<u32>().ok());
                if let (Some(min), Some(max)) = (min, max) {
                    if let Some((a, b)) = shift.adjust_span(min, max) {
                        if (a, b) != (min, max) {
                            *dirty = true;
                            updates.push(("min", Some(a.to_string())));
                            updates.push(("max", Some(b.to_string())));
                        }
                    }
                }
            }
            if updates.is_empty() {
                Ok(Open::Raw)
            } else {
                let refs: Vec<(&str, Option<&str>)> = updates
                    .iter()
                    .map(|(k, v)| (*k, v.as_deref()))
                    .collect();
                Ok(Open::Emit(set_attrs(e, &refs)))
            }
        }

        b"sheetView" | b"pane" => {
            if let Some(tlc) = attr_value(e, b"topLeftCell") {
                let new = clamp_cell(&tlc, shift);
                if new != tlc {
                    *dirty = true;
                    return Ok(Open::Emit(set_attrs(e, &[("topLeftCell", Some(&new))])));
                }
            }
            Ok(Open::Raw)
        }
        b"selection" => {
            let mut updates: Vec<(String, String)> = Vec::new();
            if let Some(ac) = attr_value(e, b"activeCell") {
                let new = clamp_cell(&ac, shift);
                if new != ac {
                    updates.push(("activeCell".into(), new));
                }
            }
            if let Some(sq) = attr_value(e, b"sqref") {
                let new = adjust_sqref_list(&sq, shift)?
                    .unwrap_or_else(|| clamp_cell(sq.split(':').next().unwrap_or("A1"), shift));
                if new != sq {
                    updates.push(("sqref".into(), new));
                }
            }
            if updates.is_empty() {
                Ok(Open::Raw)
            } else {
                *dirty = true;
                let refs: Vec<(&str, Option<&str>)> = updates
                    .iter()
                    .map(|(k, v)| (k.as_str(), Some(v.as_str())))
                    .collect();
                Ok(Open::Emit(set_attrs(e, &refs)))
            }
        }

        _ => Ok(Open::Raw),
    }
}

// ---------------------------------------------------------------------------
// comments / VML / drawing anchors
// ---------------------------------------------------------------------------

/// Shift `<comment ref="…">` / `<threadedComment ref="…">`. Comments
/// anchored on deleted cells are removed along with the cells — Excel's own
/// delete-rows behavior. (Their VML note boxes are dropped by `shift_vml`,
/// keyed on the same anchor cell; threaded replies share the parent's ref
/// and vanish with it.)
pub(super) fn shift_comments(xml: &[u8], shift: &RowColShift) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut skip_buf = Vec::new();
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("comments: {e}")))?;
        let was_empty = matches!(ev, Event::Empty(_));
        match ev {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e)
                if matches!(
                    local_name(e.name().as_ref()),
                    b"comment" | b"threadedComment"
                ) =>
            {
                let rebuilt = match attr_value(e, b"ref") {
                    Some(ref_attr) => match parse_cell_ref(&ref_attr) {
                        Some((row, col)) => {
                            let (nr, nc) = if shift.rows {
                                (shift.adjust_index(row), Some(col))
                            } else {
                                (Some(row), shift.adjust_index(col))
                            };
                            match (nr, nc) {
                                (Some(nr), Some(nc)) => {
                                    let new_ref = format_cell_ref(nr, nc);
                                    if new_ref == ref_attr {
                                        e.to_owned()
                                    } else {
                                        set_attrs(e, &[("ref", Some(&new_ref))])
                                    }
                                }
                                _ => {
                                    // Anchor cell deleted → the comment goes
                                    // with it. Skip the whole subtree.
                                    if !was_empty {
                                        skip_subtree(&mut r, &mut skip_buf, "comments")?;
                                    }
                                    buf.clear();
                                    continue;
                                }
                            }
                        }
                        None => e.to_owned(),
                    },
                    None => e.to_owned(),
                };
                write_open(&mut w, &rebuilt, was_empty)?;
            }
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}

/// Consume events until the End matching an already-consumed Start.
fn skip_subtree(
    r: &mut quick_xml::Reader<&[u8]>,
    buf: &mut Vec<u8>,
    what: &str,
) -> Result<(), PatchError> {
    let mut depth = 1usize;
    loop {
        buf.clear();
        match r
            .read_event_into(buf)
            .map_err(|e| PatchError::Xml(format!("{what}: {e}")))?
        {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
            Event::Eof => {
                return Err(PatchError::Xml(format!("{what}: truncated XML")));
            }
            _ => {}
        }
    }
}

/// Shift legacy VML anchors (comment boxes): `<x:Anchor>` holds
/// "col1, dx1, row1, dy1, col2, dx2, row2, dy2" and `<x:Row>`/`<x:Column>`
/// hold the commented cell's 0-based coordinates.
///
/// Each `<v:shape>` is buffered so that when its `<x:ClientData
/// ObjectType="Note">` anchor cell turns out to be deleted, the whole shape
/// is dropped along with its comment (mirroring `shift_comments`). Non-Note
/// shapes (buttons, controls) on deleted cells are kept and clamped onto the
/// deletion point.
fn shift_vml(xml: &[u8], shift: &RowColShift) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    // Which text element we're inside, if it's one we rewrite.
    let mut mode: Option<&'static str> = None;
    // Per-shape buffering state.
    let mut shape: Option<Writer<Vec<u8>>> = None;
    let mut shape_depth = 0usize;
    let mut is_note = false;
    let mut anchor_deleted = false;
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("vml: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                let name = name.as_slice();
                if shape.is_none() && name == b"shape" {
                    let mut sw = Writer::new(Vec::new());
                    sw.write_event(Event::Start(e)).map_err(xml_err)?;
                    shape = Some(sw);
                    shape_depth = 1;
                    is_note = false;
                    anchor_deleted = false;
                    buf.clear();
                    continue;
                }
                if name == b"ClientData"
                    && attr_value(&e, b"ObjectType").as_deref() == Some("Note")
                {
                    is_note = true;
                }
                mode = match name {
                    b"Anchor" => Some("anchor"),
                    b"Row" => Some(if shift.rows { "index" } else { "keep" }),
                    b"Column" => Some(if shift.rows { "keep" } else { "index" }),
                    _ => None,
                };
                let out = match shape.as_mut() {
                    Some(sw) => {
                        shape_depth += 1;
                        sw
                    }
                    None => &mut w,
                };
                out.write_event(Event::Start(e)).map_err(xml_err)?;
            }
            Event::Text(t) if mode.is_some() => {
                let text = t
                    .unescape()
                    .map_err(|e| PatchError::Xml(format!("vml text: {e}")))?;
                let new = match mode {
                    Some("anchor") => shift_vml_anchor(&text, shift),
                    Some("index") => match text.trim().parse::<u32>().ok() {
                        Some(i) => {
                            if shift.adjust_index(i).is_none() {
                                anchor_deleted = true;
                            }
                            shift.clamp_index(i).to_string()
                        }
                        None => text.to_string(),
                    },
                    _ => text.to_string(),
                };
                let out = shape.as_mut().unwrap_or(&mut w);
                out.write_event(Event::Text(BytesText::new(&new))).map_err(xml_err)?;
            }
            Event::End(e) => {
                mode = None;
                if let Some(sw) = shape.as_mut() {
                    sw.write_event(Event::End(e)).map_err(xml_err)?;
                    shape_depth -= 1;
                    if shape_depth == 0 {
                        let bytes = shape.take().unwrap().into_inner();
                        if !(is_note && anchor_deleted) {
                            w.get_mut().extend_from_slice(&bytes);
                        }
                    }
                    buf.clear();
                    continue;
                }
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            other => {
                let out = shape.as_mut().unwrap_or(&mut w);
                out.write_event(other).map_err(xml_err)?;
            }
        }
        buf.clear();
    }
    Ok(w.into_inner())
}

fn shift_vml_anchor(text: &str, shift: &RowColShift) -> String {
    let parts: Vec<&str> = text.split(',').collect();
    if parts.len() != 8 {
        return text.to_string();
    }
    let idx = if shift.rows { [2usize, 6] } else { [0usize, 4] };
    let mut out: Vec<String> = parts.iter().map(|p| p.to_string()).collect();
    for i in idx {
        if let Ok(v) = parts[i].trim().parse::<u32>() {
            // Preserve the original leading whitespace style loosely.
            out[i] = format!(
                "{}{}",
                &parts[i][..parts[i].len() - parts[i].trim_start().len()],
                shift.clamp_index(v)
            );
        }
    }
    out.join(",")
}

/// Shift DrawingML two-cell / one-cell anchors: `<xdr:from>`/`<xdr:to>`
/// contain `<xdr:col>`/`<xdr:row>` (0-based).
fn shift_drawing(xml: &[u8], shift: &RowColShift) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();
    let mut in_anchor_corner = false; // inside <xdr:from> or <xdr:to>
    let mut rewriting_index = false; // inside the col/row element to shift
    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("drawing: {e}")))?;
        match ev {
            Event::Eof => break,
            Event::Start(e) => {
                match local_name(e.name().as_ref()) {
                    b"from" | b"to" => in_anchor_corner = true,
                    b"col" if in_anchor_corner && !shift.rows => rewriting_index = true,
                    b"row" if in_anchor_corner && shift.rows => rewriting_index = true,
                    _ => {}
                }
                w.write_event(Event::Start(e)).map_err(xml_err)?;
            }
            Event::Text(t) if rewriting_index => {
                let text = t
                    .unescape()
                    .map_err(|e| PatchError::Xml(format!("drawing text: {e}")))?;
                let new = text
                    .trim()
                    .parse::<u32>()
                    .ok()
                    .map(|i| shift.clamp_index(i).to_string())
                    .unwrap_or_else(|| text.to_string());
                w.write_event(Event::Text(BytesText::new(&new))).map_err(xml_err)?;
            }
            Event::End(e) => {
                match local_name(e.name().as_ref()) {
                    b"from" | b"to" => in_anchor_corner = false,
                    b"col" | b"row" => rewriting_index = false,
                    _ => {}
                }
                w.write_event(Event::End(e)).map_err(xml_err)?;
            }
            other => w.write_event(other).map_err(xml_err)?,
        }
        buf.clear();
    }
    Ok(w.into_inner())
}
