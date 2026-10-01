//! Streaming worksheet XML rewriter.
//!
//! Two passes over a patched sheet:
//!   1. `scan_sheet` — collect shared-formula groups, existing column ranges,
//!      merges, and the effective style of cells targeted by style patches.
//!   2. `rewrite_sheet` — copy the XML event-by-event, splicing cell writes
//!      into `<sheetData>`, rebuilding `<cols>`/`<mergeCells>`/pane/autoFilter
//!      when asked, and leaving every untouched byte-range verbatim.
//!
//! Patched formulas carry the frontend's evaluated result as a cached `<v>`
//! when the patch provides one, so reopening the file in-app renders without
//! a full recalculation. The workbook still gets `fullCalcOnLoad`, so Excel
//! recalculates on open regardless of the cache. Materialized shared
//! formulas keep their original cached `<v>`: only their `<f>` stub is
//! expanded.

use std::collections::{BTreeMap, HashMap, HashSet};

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;

use super::refs::{format_cell_ref, parse_a1_range, parse_cell_ref};
use super::workbook_xml::{attr_value, local_name, reader, with_attr, xml_err};
use super::PatchError;

// ---------------------------------------------------------------------------
// Prepared per-sheet work (built by mod.rs from the wire patch)
// ---------------------------------------------------------------------------

#[derive(Debug, Default)]
pub struct SheetOps {
    pub rows: BTreeMap<u32, RowWork>,
    /// 0-based column -> width/hidden override.
    pub col_overrides: BTreeMap<u32, ColOverride>,
    /// Normalized "A1:B2" range -> merge(true)/unmerge(false).
    pub merge_ops: Vec<(String, bool)>,
    /// Some((rows, cols)): set frozen panes (0,0 = unfreeze). None: untouched.
    pub freeze: Option<(u32, u32)>,
    /// Some(Some(range)) set, Some(None) clear, None untouched.
    pub auto_filter: Option<Option<String>>,
}

impl SheetOps {
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty()
            && self.col_overrides.is_empty()
            && self.merge_ops.is_empty()
            && self.freeze.is_none()
            && self.auto_filter.is_none()
    }
}

#[derive(Debug, Default)]
pub struct RowWork {
    pub cells: BTreeMap<u32, CellAction>,
    pub height: Option<f64>,
    pub hidden: Option<bool>,
}

#[derive(Debug)]
pub struct CellAction {
    /// New style index; None = keep the cell's current style.
    pub style: Option<u32>,
    pub content: CellContent,
}

/// Evaluated result of a patched formula, cached into the file as `<v>` so
/// reopening in-app renders instantly instead of recomputing the whole
/// dependency graph. Excel ignores staleness here anyway: the workbook gets
/// `fullCalcOnLoad`, so Excel recalculates on open regardless.
#[derive(Debug)]
pub enum CachedValue {
    Number(f64),
    Str(String),
    Bool(bool),
    Error(String),
}

#[derive(Debug)]
pub enum CellContent {
    /// Style-only change; contents untouched.
    Keep,
    /// Clear contents, keep the (possibly overridden) style.
    Clear,
    Number(f64),
    Text(String),
    Bool(bool),
    ErrorVal(String),
    /// New formula, optionally with the frontend's evaluated result written
    /// as a cached `<v>` (see [`CachedValue`]).
    Formula(String, Option<CachedValue>),
    /// Shared-formula member being materialized: replace the `<f>` element
    /// with this plain formula text but keep the cell's other children
    /// (notably the cached `<v>`).
    MaterializeFormula(String),
}

#[derive(Debug, Clone, Default)]
pub struct ColOverride {
    pub width_chars: Option<f64>,
    pub hidden: Option<bool>,
}

// ---------------------------------------------------------------------------
// Scan pass
// ---------------------------------------------------------------------------

#[derive(Debug)]
pub struct SharedGroup {
    pub master: (u32, u32),
    pub formula: String,
    /// Every cell carrying `t="shared"` with this si, master included.
    pub members: Vec<(u32, u32)>,
}

#[derive(Debug, Clone)]
pub struct ColRange {
    pub min: u32, // 1-based, as in the XML
    pub max: u32,
    /// All attributes except min/max, preserved verbatim on split.
    pub attrs: Vec<(String, String)>,
}

#[derive(Debug, Default)]
pub struct ScanResult {
    pub shared_groups: HashMap<u32, SharedGroup>,
    /// Effective base style (cell s → row s → col style → 0) for each
    /// requested style target.
    pub style_bases: HashMap<(u32, u32), u32>,
    pub existing_cols: Vec<ColRange>,
    pub existing_merges: Vec<String>,
    /// Sheet already has a <sheetViews> element. When absent, a freeze op
    /// must insert one in schema order (before cols/sheetData), not at
    /// worksheet close.
    pub has_sheet_views: bool,
}

pub fn scan_sheet(
    xml: &[u8],
    style_targets: &HashSet<(u32, u32)>,
) -> Result<ScanResult, PatchError> {
    let mut out = ScanResult::default();
    let mut r = reader(xml);
    let mut buf = Vec::new();

    let mut row_styles: HashMap<u32, u32> = HashMap::new();
    let mut cell_styles: HashMap<(u32, u32), u32> = HashMap::new();

    // Streaming state
    let mut cur_row: u32 = 0;
    let mut next_col: u32 = 0;
    let mut cur_cell: Option<(u32, u32)> = None;
    let mut in_shared_f: Option<u32> = None; // si of the <f> we're inside
    let mut shared_text = String::new();

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("scan: {e}")))?;
        let is_empty = matches!(&ev, Event::Empty(_));
        match ev {
            Event::Eof => break,

            Event::Start(ref e) | Event::Empty(ref e) => {
                match local_name(e.name().as_ref()) {
                    b"row" => {
                        cur_row = attr_value(e, b"r")
                            .and_then(|v| v.parse::<u32>().ok())
                            .map(|v| v - 1)
                            .unwrap_or(cur_row);
                        next_col = 0;
                        if attr_value(e, b"customFormat").as_deref() == Some("1") {
                            if let Some(s) = attr_value(e, b"s").and_then(|v| v.parse().ok()) {
                                row_styles.insert(cur_row, s);
                            }
                        }
                        if !is_empty {
                            // rows advance cur_row for implicit numbering
                        }
                    }
                    b"c" => {
                        let rc = attr_value(e, b"r")
                            .and_then(|v| parse_cell_ref(&v))
                            .unwrap_or((cur_row, next_col));
                        next_col = rc.1 + 1;
                        cur_cell = Some(rc);
                        if style_targets.contains(&rc) {
                            if let Some(s) = attr_value(e, b"s").and_then(|v| v.parse().ok()) {
                                cell_styles.insert(rc, s);
                            }
                        }
                    }
                    b"f" => {
                        if attr_value(e, b"t").as_deref() == Some("shared") {
                            if let (Some(si), Some(rc)) = (
                                attr_value(e, b"si").and_then(|v| v.parse::<u32>().ok()),
                                cur_cell,
                            ) {
                                let group =
                                    out.shared_groups.entry(si).or_insert_with(|| SharedGroup {
                                        master: rc,
                                        formula: String::new(),
                                        members: Vec::new(),
                                    });
                                group.members.push(rc);
                                if !is_empty {
                                    in_shared_f = Some(si);
                                    shared_text.clear();
                                }
                            }
                        }
                    }
                    b"col" => {
                        let min = attr_value(e, b"min").and_then(|v| v.parse().ok());
                        let max = attr_value(e, b"max").and_then(|v| v.parse().ok());
                        if let (Some(min), Some(max)) = (min, max) {
                            let mut attrs = Vec::new();
                            for a in e.attributes().with_checks(false).flatten() {
                                let k = String::from_utf8_lossy(a.key.as_ref()).into_owned();
                                if k != "min" && k != "max" {
                                    let v = String::from_utf8_lossy(&a.value).into_owned();
                                    attrs.push((k, v));
                                }
                            }
                            out.existing_cols.push(ColRange { min, max, attrs });
                        }
                    }
                    b"mergeCell" => {
                        if let Some(range) = attr_value(e, b"ref") {
                            out.existing_merges.push(range);
                        }
                    }
                    b"sheetViews" => {
                        out.has_sheet_views = true;
                    }
                    _ => {}
                }
            }
            Event::Text(ref t) => {
                if in_shared_f.is_some() {
                    let txt = t
                        .unescape()
                        .map_err(|e| PatchError::Xml(format!("scan text: {e}")))?;
                    shared_text.push_str(&txt);
                }
            }
            Event::End(ref e) => match local_name(e.name().as_ref()) {
                b"f" => {
                    if let Some(si) = in_shared_f.take() {
                        if !shared_text.is_empty() {
                            if let Some(g) = out.shared_groups.get_mut(&si) {
                                // The cell holding the text is the master.
                                if let Some(&last) = g.members.last() {
                                    g.master = last;
                                }
                                g.formula = shared_text.clone();
                            }
                        }
                    }
                }
                b"c" => cur_cell = None,
                b"row" => cur_row += 1,
                _ => {}
            },
            _ => {}
        }
        buf.clear();
    }

    // Effective base style per style target: cell → row → col → 0.
    for &(rr, cc) in style_targets {
        let base = cell_styles.get(&(rr, cc)).copied().or_else(|| {
            row_styles.get(&rr).copied().or_else(|| {
                out.existing_cols
                    .iter()
                    .find(|c| c.min <= cc + 1 && cc + 1 <= c.max)
                    .and_then(|c| {
                        c.attrs
                            .iter()
                            .find(|(k, _)| k == "style")
                            .and_then(|(_, v)| v.parse().ok())
                    })
            })
        });
        out.style_bases.insert((rr, cc), base.unwrap_or(0));
    }

    Ok(out)
}

// ---------------------------------------------------------------------------
// Rewrite pass
// ---------------------------------------------------------------------------

/// CT_Worksheet child sequence (ECMA-376). Used both to insert brand-new
/// elements at a legal spot and by the save-time validator.
pub(super) const WORKSHEET_CHILD_ORDER: &[&[u8]] = &[
    b"sheetPr",
    b"dimension",
    b"sheetViews",
    b"sheetFormatPr",
    b"cols",
    b"sheetData",
    b"sheetCalcPr",
    b"sheetProtection",
    b"protectedRanges",
    b"scenarios",
    b"autoFilter",
    b"sortState",
    b"dataConsolidate",
    b"customSheetViews",
    b"mergeCells",
    b"phoneticPr",
    b"conditionalFormatting",
    b"dataValidations",
    b"hyperlinks",
    b"printOptions",
    b"pageMargins",
    b"pageSetup",
    b"headerFooter",
    b"rowBreaks",
    b"colBreaks",
    b"customProperties",
    b"cellWatches",
    b"ignoredErrors",
    b"smartTags",
    b"drawing",
    b"legacyDrawing",
    b"legacyDrawingHF",
    b"drawingHF",
    b"picture",
    b"oleObjects",
    b"controls",
    b"webPublishItems",
    b"tableParts",
    b"extLst",
];

/// Schema position of worksheet child elements (CT_Worksheet sequence).
/// Used to insert brand-new elements (cols, autoFilter, mergeCells,
/// sheetViews) at a legal spot.
fn order_index(name: &[u8]) -> usize {
    WORKSHEET_CHILD_ORDER
        .iter()
        .position(|n| *n == name)
        .unwrap_or(usize::MAX - 1)
}

pub fn rewrite_sheet(xml: &[u8], ops: &SheetOps, scan: &ScanResult) -> Result<Vec<u8>, PatchError> {
    let mut r = reader(xml);
    let mut w = Writer::new(Vec::new());
    let mut buf = Vec::new();

    // Rows still waiting to be spliced in (consumed as we pass their spot).
    let mut pending: BTreeMap<u32, &RowWork> = ops.rows.iter().map(|(k, v)| (*k, v)).collect();

    // Deferred whole-element inserts: (schema order idx, XML bytes).
    let mut inserts: Vec<(usize, Vec<u8>)> = Vec::new();
    let rebuild_cols = !ops.col_overrides.is_empty();
    if rebuild_cols {
        let merged = apply_col_overrides(&scan.existing_cols, &ops.col_overrides);
        if !merged.is_empty() {
            inserts.push((order_index(b"cols"), serialize_cols(&merged)));
        }
    }
    let rebuild_merges = !ops.merge_ops.is_empty();
    if rebuild_merges {
        let merged = apply_merge_ops(&scan.existing_merges, &ops.merge_ops)?;
        if !merged.is_empty() {
            inserts.push((order_index(b"mergeCells"), serialize_merges(&merged)));
        }
    }
    if let Some(Some(range)) = &ops.auto_filter {
        let mut e = BytesStart::new("autoFilter");
        e.push_attribute(("ref", range.as_str()));
        let mut v = Vec::new();
        Writer::new(&mut v)
            .write_event(Event::Empty(e))
            .map_err(xml_err)?;
        inserts.push((order_index(b"autoFilter"), v));
    }
    // Freeze on a sheet with no <sheetViews>: build the whole element now so
    // it lands at its legal spot in the CT_Worksheet sequence (before cols /
    // sheetData), not at worksheet close where Excel rejects it.
    if !scan.has_sheet_views {
        if let Some((fr, fc)) = ops.freeze {
            if fr > 0 || fc > 0 {
                let mut sw = Writer::new(Vec::new());
                sw.write_event(Event::Start(BytesStart::new("sheetViews")))
                    .map_err(xml_err)?;
                let mut sv = BytesStart::new("sheetView");
                sv.push_attribute(("workbookViewId", "0"));
                sw.write_event(Event::Start(sv)).map_err(xml_err)?;
                write_pane(&mut sw, fr, fc)?;
                sw.write_event(Event::End(BytesEnd::new("sheetView")))
                    .map_err(xml_err)?;
                sw.write_event(Event::End(BytesEnd::new("sheetViews")))
                    .map_err(xml_err)?;
                inserts.push((order_index(b"sheetViews"), sw.into_inner()));
            }
        }
    }

    let mut depth = 0usize;
    let mut in_sheet_data = false;
    let mut skipping: Option<usize> = None; // skip everything until depth returns here
    let mut cur_row: u32 = 0;
    let mut next_col: u32 = 0;

    // Row splice state (only while inside a <row> that has work).
    struct RowState<'a> {
        row: u32,
        work: Option<&'a RowWork>,
        emitted: HashSet<u32>,
    }
    let mut row_state: Option<RowState> = None;

    // Cell splice state while we're inside a <c> being kept (for shared
    // formula materialization): the plain formula to write when we hit <f>.
    let mut cell_materialize: Option<String> = None;

    // sheetView pane handling
    let mut in_sheet_view = false;

    loop {
        let ev = r
            .read_event_into(&mut buf)
            .map_err(|e| PatchError::Xml(format!("rewrite: {e}")))?;

        if let Some(k) = skipping {
            match ev {
                Event::Start(_) => depth += 1,
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if depth < k {
                        skipping = None;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
            continue;
        }

        match ev {
            Event::Eof => break,

            Event::Decl(_) | Event::PI(_) | Event::Comment(_) | Event::DocType(_) => {
                w.write_event(ev).map_err(xml_err)?;
            }

            Event::Start(e) => {
                let name_owned = e.name().as_ref().to_vec();
                let name = local_name(&name_owned).to_vec();
                // Materialized shared formula: replace this cell's <f> and
                // swallow the original stub/text.
                if cell_materialize.is_some() && name == b"f" {
                    let ftext = cell_materialize.take().unwrap();
                    write_simple(&mut w, "f", &ftext)?;
                    depth += 1;
                    skipping = Some(depth);
                    buf.clear();
                    continue;
                }
                // depth 1 = children of <worksheet>
                if depth == 1 {
                    flush_inserts_before(&mut w, &mut inserts, order_index(&name))?;
                    if (name == b"cols" && rebuild_cols)
                        || (name == b"mergeCells" && rebuild_merges)
                        || (name == b"autoFilter" && ops.auto_filter.is_some())
                    {
                        // Replaced (or removed) wholesale: emit the queued
                        // replacement at the old element's position, skip the
                        // original.
                        flush_inserts_before(&mut w, &mut inserts, order_index(&name) + 1)?;
                        depth += 1;
                        skipping = Some(depth);
                        buf.clear();
                        continue;
                    }
                    if name == b"sheetData" {
                        in_sheet_data = true;
                        cur_row = 0;
                    }
                }

                if in_sheet_data && name == b"row" {
                    let row = attr_value(&e, b"r")
                        .and_then(|v| v.parse::<u32>().ok())
                        .map(|v| v - 1)
                        .unwrap_or(cur_row);
                    cur_row = row;
                    next_col = 0;
                    // Splice pending rows that belong strictly before this one.
                    let before: Vec<u32> = pending.range(..row).map(|(k, _)| *k).collect();
                    for key in before {
                        let work = pending.remove(&key).unwrap();
                        write_full_row(&mut w, key, work, scan)?;
                    }
                    let work = pending.remove(&row);
                    let start = rebuild_row_start(&e, work);
                    depth += 1;
                    w.write_event(Event::Start(start)).map_err(xml_err)?;
                    row_state = Some(RowState {
                        row,
                        work,
                        emitted: HashSet::new(),
                    });
                    buf.clear();
                    continue;
                }

                if in_sheet_data && name == b"c" {
                    let rc = attr_value(&e, b"r")
                        .and_then(|v| parse_cell_ref(&v))
                        .unwrap_or((cur_row, next_col));
                    next_col = rc.1 + 1;

                    if let Some(rs) = row_state.as_mut() {
                        // Emit pending new cells that come before this one.
                        if let Some(work) = rs.work {
                            let before: Vec<u32> = work
                                .cells
                                .range(..rc.1)
                                .filter(|(k, _)| !rs.emitted.contains(*k))
                                .map(|(k, _)| *k)
                                .collect();
                            for col in before {
                                if let Some(action) = work.cells.get(&col) {
                                    if !action_targets_existing_only(action) {
                                        write_new_cell(&mut w, rs.row, col, action, scan)?;
                                    }
                                    rs.emitted.insert(col);
                                }
                            }
                            if let Some(action) = work.cells.get(&rc.1) {
                                rs.emitted.insert(rc.1);
                                match &action.content {
                                    CellContent::MaterializeFormula(f) => {
                                        // Keep the cell, rewrite only its <f>.
                                        let start = match action.style {
                                            Some(s) => with_attr(&e, "s", &s.to_string()),
                                            None => e.to_owned(),
                                        };
                                        depth += 1;
                                        w.write_event(Event::Start(start)).map_err(xml_err)?;
                                        cell_materialize = Some(f.clone());
                                        buf.clear();
                                        continue;
                                    }
                                    CellContent::Keep => {
                                        // Style-only: retag, keep contents.
                                        let start = match action.style {
                                            Some(s) => with_attr(&e, "s", &s.to_string()),
                                            None => e.to_owned(),
                                        };
                                        depth += 1;
                                        w.write_event(Event::Start(start)).map_err(xml_err)?;
                                        buf.clear();
                                        continue;
                                    }
                                    _ => {
                                        // Replace the cell wholesale: write the
                                        // new version, skip the original.
                                        let orig_style = attr_value(&e, b"s")
                                            .and_then(|v| v.parse::<u32>().ok());
                                        write_cell(
                                            &mut w,
                                            rs.row,
                                            rc.1,
                                            action,
                                            action.style.or(orig_style),
                                        )?;
                                        depth += 1;
                                        skipping = Some(depth);
                                        buf.clear();
                                        continue;
                                    }
                                }
                            }
                        }
                    }
                    // Untouched cell: fall through to verbatim copy.
                }

                if in_sheet_view && local_name(&name) == b"pane" && ops.freeze.is_some() {
                    // Old pane replaced by the one we injected at sheetView start.
                    depth += 1;
                    skipping = Some(depth);
                    buf.clear();
                    continue;
                }
                if in_sheet_view
                    && local_name(&name) == b"selection"
                    && ops.freeze.is_some()
                    && attr_value(&e, b"pane").is_some()
                {
                    // Selections bound to old panes would dangle.
                    depth += 1;
                    skipping = Some(depth);
                    buf.clear();
                    continue;
                }

                let is_sheet_view = depth == 2 && local_name(&name) == b"sheetView";
                depth += 1;
                w.write_event(Event::Start(e)).map_err(xml_err)?;
                if is_sheet_view && !in_sheet_view {
                    in_sheet_view = true;
                    if let Some((fr, fc)) = ops.freeze {
                        write_pane(&mut w, fr, fc)?;
                    }
                }
            }

            Event::Empty(e) => {
                let name_owned = e.name().as_ref().to_vec();
                let name = local_name(&name_owned).to_vec();
                // Shared-formula member stub: <f t="shared" si="N"/>
                if cell_materialize.is_some() && name == b"f" {
                    let ftext = cell_materialize.take().unwrap();
                    write_simple(&mut w, "f", &ftext)?;
                    buf.clear();
                    continue;
                }
                if depth == 1 {
                    flush_inserts_before(&mut w, &mut inserts, order_index(&name))?;
                    if (name == b"cols" && rebuild_cols)
                        || (name == b"mergeCells" && rebuild_merges)
                        || (name == b"autoFilter" && ops.auto_filter.is_some())
                    {
                        flush_inserts_before(&mut w, &mut inserts, order_index(&name) + 1)?;
                        buf.clear();
                        continue;
                    }
                    if name == b"sheetData" {
                        // Self-closing sheetData with pending rows → expand.
                        if !pending.is_empty() {
                            w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
                            let keys: Vec<u32> = pending.keys().copied().collect();
                            for key in keys {
                                let work = pending.remove(&key).unwrap();
                                write_full_row(&mut w, key, work, scan)?;
                            }
                            w.write_event(Event::End(BytesEnd::new(
                                String::from_utf8_lossy(&name_owned).into_owned(),
                            )))
                            .map_err(xml_err)?;
                            buf.clear();
                            continue;
                        }
                    }
                }

                if in_sheet_data && name == b"row" {
                    // Row with no cells (attrs only).
                    let row = attr_value(&e, b"r")
                        .and_then(|v| v.parse::<u32>().ok())
                        .map(|v| v - 1)
                        .unwrap_or(cur_row);
                    let before: Vec<u32> = pending.range(..row).map(|(k, _)| *k).collect();
                    for key in before {
                        let work = pending.remove(&key).unwrap();
                        write_full_row(&mut w, key, work, scan)?;
                    }
                    if let Some(work) = pending.remove(&row) {
                        let start = rebuild_row_start(&e, Some(work));
                        w.write_event(Event::Start(start)).map_err(xml_err)?;
                        for (col, action) in &work.cells {
                            if !action_targets_existing_only(action) {
                                write_new_cell(&mut w, row, *col, action, scan)?;
                            }
                        }
                        w.write_event(Event::End(BytesEnd::new("row"))).map_err(xml_err)?;
                    } else {
                        w.write_event(Event::Empty(e)).map_err(xml_err)?;
                    }
                    cur_row = row + 1;
                    buf.clear();
                    continue;
                }

                if in_sheet_data && name == b"c" {
                    let rc = attr_value(&e, b"r")
                        .and_then(|v| parse_cell_ref(&v))
                        .unwrap_or((cur_row, next_col));
                    next_col = rc.1 + 1;
                    if let Some(rs) = row_state.as_mut() {
                        if let Some(work) = rs.work {
                            let before: Vec<u32> = work
                                .cells
                                .range(..rc.1)
                                .filter(|(k, _)| !rs.emitted.contains(*k))
                                .map(|(k, _)| *k)
                                .collect();
                            for col in before {
                                if let Some(action) = work.cells.get(&col) {
                                    if !action_targets_existing_only(action) {
                                        write_new_cell(&mut w, rs.row, col, action, scan)?;
                                    }
                                    rs.emitted.insert(col);
                                }
                            }
                            if let Some(action) = work.cells.get(&rc.1) {
                                rs.emitted.insert(rc.1);
                                if matches!(action.content, CellContent::Keep) {
                                    // Style-only on an empty cell: retag it.
                                    let start = match action.style {
                                        Some(s) => with_attr(&e, "s", &s.to_string()),
                                        None => e.to_owned(),
                                    };
                                    w.write_event(Event::Empty(start)).map_err(xml_err)?;
                                    buf.clear();
                                    continue;
                                }
                                let orig_style =
                                    attr_value(&e, b"s").and_then(|v| v.parse::<u32>().ok());
                                write_cell(
                                    &mut w,
                                    rs.row,
                                    rc.1,
                                    action,
                                    action.style.or(orig_style),
                                )?;
                                buf.clear();
                                continue;
                            }
                        }
                    }
                }

                if in_sheet_view && name == b"pane" && ops.freeze.is_some() {
                    buf.clear();
                    continue;
                }
                if in_sheet_view
                    && name == b"selection"
                    && ops.freeze.is_some()
                    && attr_value(&e, b"pane").is_some()
                {
                    buf.clear();
                    continue;
                }

                if depth == 2 && name == b"sheetView" {
                    // Self-closing sheetView but we need a pane inside it.
                    if let Some((fr, fc)) = ops.freeze {
                        if fr > 0 || fc > 0 {
                            w.write_event(Event::Start(e.to_owned())).map_err(xml_err)?;
                            write_pane(&mut w, fr, fc)?;
                            w.write_event(Event::End(BytesEnd::new("sheetView")))
                                .map_err(xml_err)?;
                            buf.clear();
                            continue;
                        }
                    }
                }

                w.write_event(Event::Empty(e)).map_err(xml_err)?;
            }

            Event::End(e) => {
                let name_owned = e.name().as_ref().to_vec();
                let name = local_name(&name_owned).to_vec();

                if in_sheet_data && name == b"row" {
                    // Flush this row's remaining new cells before closing it.
                    if let Some(rs) = row_state.take() {
                        if let Some(work) = rs.work {
                            for (col, action) in &work.cells {
                                if !rs.emitted.contains(col)
                                    && !action_targets_existing_only(action)
                                {
                                    write_new_cell(&mut w, rs.row, *col, action, scan)?;
                                }
                            }
                        }
                        cur_row = rs.row + 1;
                        next_col = 0;
                    } else {
                        cur_row += 1;
                    }
                    depth = depth.saturating_sub(1);
                    w.write_event(Event::End(e)).map_err(xml_err)?;
                    buf.clear();
                    continue;
                }

                if name == b"sheetData" && in_sheet_data {
                    // Rows past the last original row.
                    let keys: Vec<u32> = pending.keys().copied().collect();
                    for key in keys {
                        let work = pending.remove(&key).unwrap();
                        write_full_row(&mut w, key, work, scan)?;
                    }
                    in_sheet_data = false;
                }
                if name == b"sheetView" {
                    in_sheet_view = false;
                }
                if name == b"c" {
                    cell_materialize = None;
                }
                if name == b"worksheet" {
                    // Anything never flushed (element types the sheet lacked
                    // entirely) goes before the close tag.
                    flush_inserts_before(&mut w, &mut inserts, usize::MAX)?;
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

    // Sanity: every patched row must have been spliced somewhere.
    if !pending.is_empty() {
        return Err(PatchError::Xml("sheetData not found for pending rows".into()));
    }

    Ok(w.into_inner())
}

/// Style-only/materialize actions must not create cells out of thin air when
/// flushed as "new" — Keep on a nonexistent cell still writes an empty styled
/// cell (that's how a style lands on a blank cell), so only Materialize is
/// existing-only.
fn action_targets_existing_only(action: &CellAction) -> bool {
    matches!(action.content, CellContent::MaterializeFormula(_))
}

fn flush_inserts_before(
    w: &mut Writer<Vec<u8>>,
    inserts: &mut Vec<(usize, Vec<u8>)>,
    upcoming: usize,
) -> Result<(), PatchError> {
    inserts.sort_by_key(|(idx, _)| *idx);
    let mut i = 0;
    while i < inserts.len() {
        if inserts[i].0 < upcoming {
            let (_, bytes) = inserts.remove(i);
            w.get_mut()
                .extend_from_slice(&bytes);
        } else {
            i += 1;
        }
    }
    Ok(())
}

/// Row start tag with height/hidden overrides applied and `spans` dropped
/// (we may add cells outside the original span hint).
fn rebuild_row_start(e: &BytesStart, work: Option<&RowWork>) -> BytesStart<'static> {
    let mut out = BytesStart::new("row");
    let (height, hidden) = match work {
        Some(wk) => (wk.height, wk.hidden),
        None => (None, None),
    };
    for attr in e.attributes().with_checks(false).flatten() {
        let key = attr.key.as_ref();
        if key == b"spans" && work.is_some() {
            continue;
        }
        if (key == b"ht" || key == b"customHeight") && height.is_some() {
            continue;
        }
        if key == b"hidden" && hidden.is_some() {
            continue;
        }
        out.push_attribute(attr);
    }
    if let Some(h) = height {
        out.push_attribute(("ht", trim_float(h).as_str()));
        out.push_attribute(("customHeight", "1"));
    }
    if hidden == Some(true) {
        out.push_attribute(("hidden", "1"));
    }
    out
}

/// A brand-new row (didn't exist in the original XML) with all its cells.
fn write_full_row(
    w: &mut Writer<Vec<u8>>,
    row: u32,
    work: &RowWork,
    scan: &ScanResult,
) -> Result<(), PatchError> {
    let mut start = BytesStart::new("row");
    start.push_attribute(("r", (row + 1).to_string().as_str()));
    if let Some(h) = work.height {
        start.push_attribute(("ht", trim_float(h).as_str()));
        start.push_attribute(("customHeight", "1"));
    }
    if work.hidden == Some(true) {
        start.push_attribute(("hidden", "1"));
    }
    w.write_event(Event::Start(start)).map_err(xml_err)?;
    for (col, action) in &work.cells {
        if !action_targets_existing_only(action) {
            write_new_cell(w, row, *col, action, scan)?;
        }
    }
    w.write_event(Event::End(BytesEnd::new("row"))).map_err(xml_err)?;
    Ok(())
}

/// A cell that doesn't exist in the original sheet: inherit the effective
/// style (row/col defaults) unless the action pins one.
fn write_new_cell(
    w: &mut Writer<Vec<u8>>,
    row: u32,
    col: u32,
    action: &CellAction,
    scan: &ScanResult,
) -> Result<(), PatchError> {
    let style = action
        .style
        .or_else(|| scan.style_bases.get(&(row, col)).copied().filter(|&s| s != 0));
    write_cell(w, row, col, action, style)
}

fn write_cell(
    w: &mut Writer<Vec<u8>>,
    row: u32,
    col: u32,
    action: &CellAction,
    style: Option<u32>,
) -> Result<(), PatchError> {
    let mut c = BytesStart::new("c");
    c.push_attribute(("r", format_cell_ref(row, col).as_str()));
    if let Some(s) = style {
        c.push_attribute(("s", s.to_string().as_str()));
    }

    match &action.content {
        CellContent::Keep | CellContent::Clear => {
            w.write_event(Event::Empty(c)).map_err(xml_err)?;
        }
        CellContent::Number(n) => {
            if !n.is_finite() {
                return Err(PatchError::BadValue(format!("non-finite number {n}")));
            }
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            write_simple(w, "v", &trim_float(*n))?;
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
        CellContent::Text(s) => {
            c.push_attribute(("t", "inlineStr"));
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            w.write_event(Event::Start(BytesStart::new("is"))).map_err(xml_err)?;
            let mut t = BytesStart::new("t");
            t.push_attribute(("xml:space", "preserve"));
            w.write_event(Event::Start(t)).map_err(xml_err)?;
            w.write_event(Event::Text(BytesText::new(&sanitize_xlsx_text(s))))
                .map_err(xml_err)?;
            w.write_event(Event::End(BytesEnd::new("t"))).map_err(xml_err)?;
            w.write_event(Event::End(BytesEnd::new("is"))).map_err(xml_err)?;
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
        CellContent::Bool(b) => {
            c.push_attribute(("t", "b"));
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            write_simple(w, "v", if *b { "1" } else { "0" })?;
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
        CellContent::ErrorVal(e) => {
            c.push_attribute(("t", "e"));
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            write_simple(w, "v", e)?;
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
        CellContent::Formula(f, cached) => {
            // Cached-result typing follows OOXML: formula string results are
            // `t="str"` (NOT inlineStr — that's invalid next to <f>), errors
            // `t="e"`, booleans `t="b"`, numbers untyped.
            match cached {
                Some(CachedValue::Str(_)) => c.push_attribute(("t", "str")),
                Some(CachedValue::Bool(_)) => c.push_attribute(("t", "b")),
                Some(CachedValue::Error(_)) => c.push_attribute(("t", "e")),
                Some(CachedValue::Number(_)) | None => {}
            }
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            write_simple(w, "f", f)?;
            match cached {
                Some(CachedValue::Number(n)) => write_simple(w, "v", &trim_float(*n))?,
                Some(CachedValue::Str(s)) => write_simple(w, "v", &sanitize_xlsx_text(s))?,
                Some(CachedValue::Bool(b)) => write_simple(w, "v", if *b { "1" } else { "0" })?,
                Some(CachedValue::Error(e)) => write_simple(w, "v", e)?,
                None => {}
            }
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
        CellContent::MaterializeFormula(f) => {
            w.write_event(Event::Start(c)).map_err(xml_err)?;
            write_simple(w, "f", f)?;
            w.write_event(Event::End(BytesEnd::new("c"))).map_err(xml_err)?;
        }
    }
    Ok(())
}

fn write_simple(w: &mut Writer<Vec<u8>>, tag: &str, text: &str) -> Result<(), PatchError> {
    w.write_event(Event::Start(BytesStart::new(tag))).map_err(xml_err)?;
    w.write_event(Event::Text(BytesText::new(text))).map_err(xml_err)?;
    w.write_event(Event::End(BytesEnd::new(tag))).map_err(xml_err)?;
    Ok(())
}

fn write_pane(w: &mut Writer<Vec<u8>>, rows: u32, cols: u32) -> Result<(), PatchError> {
    if rows == 0 && cols == 0 {
        return Ok(()); // unfreeze = simply don't write a pane
    }
    let mut pane = BytesStart::new("pane");
    if cols > 0 {
        pane.push_attribute(("xSplit", cols.to_string().as_str()));
    }
    if rows > 0 {
        pane.push_attribute(("ySplit", rows.to_string().as_str()));
    }
    pane.push_attribute(("topLeftCell", format_cell_ref(rows, cols).as_str()));
    let active = match (rows > 0, cols > 0) {
        (true, true) => "bottomRight",
        (true, false) => "bottomLeft",
        (false, true) => "topRight",
        (false, false) => unreachable!(),
    };
    pane.push_attribute(("activePane", active));
    pane.push_attribute(("state", "frozen"));
    w.write_event(Event::Empty(pane)).map_err(xml_err)?;
    Ok(())
}

// ---------------------------------------------------------------------------
// cols / merges rebuilding
// ---------------------------------------------------------------------------

fn apply_col_overrides(
    existing: &[ColRange],
    overrides: &BTreeMap<u32, ColOverride>,
) -> Vec<ColRange> {
    // Explode affected ranges at override boundaries, then patch the
    // single-column ranges. Untouched ranges pass through as-is.
    let mut out: Vec<ColRange> = Vec::new();
    let mut covered: HashSet<u32> = HashSet::new();

    for range in existing {
        let hits: Vec<u32> = overrides
            .keys()
            .copied()
            .filter(|c| range.min <= c + 1 && c + 1 <= range.max)
            .collect();
        if hits.is_empty() {
            out.push(range.clone());
            continue;
        }
        let mut cursor = range.min;
        for c0 in &hits {
            let col1 = c0 + 1;
            if cursor < col1 {
                out.push(ColRange {
                    min: cursor,
                    max: col1 - 1,
                    attrs: range.attrs.clone(),
                });
            }
            out.push(patched_range(col1, &range.attrs, &overrides[c0]));
            covered.insert(*c0);
            cursor = col1 + 1;
        }
        if cursor <= range.max {
            out.push(ColRange {
                min: cursor,
                max: range.max,
                attrs: range.attrs.clone(),
            });
        }
    }

    for (c0, ov) in overrides {
        if !covered.contains(c0) {
            out.push(patched_range(c0 + 1, &[], ov));
        }
    }

    out.sort_by_key(|r| r.min);
    out
}

fn patched_range(col1: u32, base_attrs: &[(String, String)], ov: &ColOverride) -> ColRange {
    let mut attrs: Vec<(String, String)> = base_attrs
        .iter()
        .filter(|(k, _)| {
            let drop_width = ov.width_chars.is_some() && (k == "width" || k == "customWidth");
            let drop_hidden = ov.hidden.is_some() && k == "hidden";
            !(drop_width || drop_hidden)
        })
        .cloned()
        .collect();
    if let Some(wd) = ov.width_chars {
        attrs.push(("width".into(), trim_float(wd)));
        attrs.push(("customWidth".into(), "1".into()));
    }
    if ov.hidden == Some(true) {
        attrs.push(("hidden".into(), "1".into()));
    }
    ColRange {
        min: col1,
        max: col1,
        attrs,
    }
}

fn serialize_cols(ranges: &[ColRange]) -> Vec<u8> {
    let mut v = Vec::new();
    let mut w = Writer::new(&mut v);
    w.write_event(Event::Start(BytesStart::new("cols"))).ok();
    for r in ranges {
        let mut e = BytesStart::new("col");
        e.push_attribute(("min", r.min.to_string().as_str()));
        e.push_attribute(("max", r.max.to_string().as_str()));
        for (k, val) in &r.attrs {
            e.push_attribute((k.as_str(), val.as_str()));
        }
        w.write_event(Event::Empty(e)).ok();
    }
    w.write_event(Event::End(BytesEnd::new("cols"))).ok();
    v
}

fn apply_merge_ops(
    existing: &[String],
    ops: &[(String, bool)],
) -> Result<Vec<String>, PatchError> {
    let norm = |s: &str| -> Option<String> {
        parse_a1_range(s).map(|(r1, c1, r2, c2)| {
            format!("{}:{}", format_cell_ref(r1, c1), format_cell_ref(r2, c2))
        })
    };
    let mut set: Vec<String> = existing.iter().filter_map(|s| norm(s)).collect();
    for (range, merge) in ops {
        let Some(key) = norm(range) else { continue };
        if *merge {
            if !set.contains(&key) {
                set.push(key);
            }
        } else {
            set.retain(|s| s != &key);
        }
    }
    // Overlapping merges make Excel throw the repair dialog on open. Refuse
    // and let the save fall back (visibly) rather than write a broken file.
    let parsed: Vec<_> = set.iter().filter_map(|s| parse_a1_range(s)).collect();
    for (i, a) in parsed.iter().enumerate() {
        for b in &parsed[i + 1..] {
            let disjoint = a.2 < b.0 || b.2 < a.0 || a.3 < b.1 || b.3 < a.1;
            if !disjoint {
                return Err(PatchError::BadValue(format!(
                    "overlapping merge ranges {} would corrupt the sheet",
                    set[i]
                )));
            }
        }
    }
    Ok(set)
}

fn serialize_merges(ranges: &[String]) -> Vec<u8> {
    let mut v = Vec::new();
    let mut w = Writer::new(&mut v);
    let mut start = BytesStart::new("mergeCells");
    start.push_attribute(("count", ranges.len().to_string().as_str()));
    w.write_event(Event::Start(start)).ok();
    for r in ranges {
        let mut e = BytesStart::new("mergeCell");
        e.push_attribute(("ref", r.as_str()));
        w.write_event(Event::Empty(e)).ok();
    }
    w.write_event(Event::End(BytesEnd::new("mergeCells"))).ok();
    v
}

// ---------------------------------------------------------------------------
// text helpers
// ---------------------------------------------------------------------------

/// Excel encodes characters illegal in XML 1.0 as _xHHHH_. A literal "_x"
/// must itself be escaped as _x005F_x so decoding stays unambiguous.
pub fn sanitize_xlsx_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let ch = chars[i];
        if ch == '_'
            && i + 1 < chars.len()
            && chars[i + 1] == 'x'
            && looks_like_x_escape(&chars[i..])
        {
            out.push_str("_x005F_x");
            i += 2;
            continue;
        }
        if ch < '\u{20}' && ch != '\t' && ch != '\n' && ch != '\r' {
            out.push_str(&format!("_x{:04X}_", ch as u32));
        } else {
            out.push(ch);
        }
        i += 1;
    }
    out
}

fn looks_like_x_escape(tail: &[char]) -> bool {
    // _xHHHH_ — 2 prefix chars + 4 hex + closing underscore
    tail.len() >= 7
        && tail[2..6].iter().all(|c| c.is_ascii_hexdigit())
        && tail[6] == '_'
}

/// Format a float the way Excel likes: shortest round-trip, no trailing ".0".
fn trim_float(v: f64) -> String {
    if v == v.trunc() && v.abs() < 1e15 {
        format!("{}", v as i64)
    } else {
        format!("{v}")
    }
}
