//! Surgical xlsx save.
//!
//! `apply_patch` takes the ORIGINAL workbook bytes plus a cell/format patch
//! and produces new workbook bytes where every zip entry we had no reason to
//! touch is copied raw. See docs/surgical-save-plan.md for the design.

pub mod package;
pub mod patch;
pub mod refs;
pub mod row_col;
pub mod sheet_xml;
pub mod structure;
pub mod styles;
pub mod validate;
pub mod workbook_xml;

#[cfg(test)]
pub(crate) mod tests;

use std::collections::{BTreeMap, HashSet};

use patch::{CellPatch, CellValue, Patch, SheetOp};
use sheet_xml::{CachedValue, CellAction, CellContent, ColOverride, RowWork, SheetOps};

pub const PATCH_VERSION: u32 = 1;

#[derive(Debug, thiserror::Error)]
pub enum PatchError {
    #[error("unsupported patch version {0}")]
    Version(u32),
    #[error("patch json: {0}")]
    Json(String),
    #[error("zip: {0}")]
    Zip(String),
    #[error("xml: {0}")]
    Xml(String),
    #[error("missing part: {0}")]
    MissingPart(String),
    #[error("sheet not found in workbook: {0}")]
    SheetNotFound(String),
    #[error("bad value: {0}")]
    BadValue(String),
    #[error("malformed reference")]
    BadRef,
    #[error("shifted reference out of sheet bounds")]
    RefOutOfRange,
    #[error("shared formula group {0} has no master formula")]
    SharedFormulaMissing(u32),
    /// The workbook contains features the structural engine can't rewrite
    /// safely (pivots, tables, slicers…). Callers fall back — visibly.
    #[error("unsupported for surgical save: {0}")]
    Unsupported(String),
    /// The patched output failed the save-time validity gate. Nothing was
    /// written; the generator has a bug that would have made Excel offer
    /// to repair the file.
    #[error("output failed validation: {0}")]
    Validation(String),
}

pub fn apply_patch_json(base: &[u8], patch_json: &str) -> Result<Vec<u8>, PatchError> {
    let patch: Patch =
        serde_json::from_str(patch_json).map_err(|e| PatchError::Json(e.to_string()))?;
    if patch.version != PATCH_VERSION {
        return Err(PatchError::Version(patch.version));
    }
    apply_patch(base, &patch)
}

pub fn apply_patch(base: &[u8], patch: &Patch) -> Result<Vec<u8>, PatchError> {
    // The safest save is the one that rewrites nothing.
    if patch.is_empty() {
        return Ok(base.to_vec());
    }

    let mut store = package::PartStore::open(base)?;

    // Resolve workbook.xml through the package rels rather than assuming the
    // conventional path.
    let root_rels = store.read("_rels/.rels")?;
    let workbook_path = find_rel_target(&root_rels, "/officeDocument", "")
        .unwrap_or_else(|| "xl/workbook.xml".to_string());
    let (wb_dir, wb_file) = split_part_path(&workbook_path);
    let wb_rels_path = format!("{wb_dir}_rels/{wb_file}.rels");

    // --- structural ops, in order: sheets first, then row/col shifts ---
    // Cell patches below address sheets by their live (post-op) names and
    // post-shift coordinates.
    for op in &patch.sheet_ops {
        match op {
            SheetOp::Create { name, tab_color } => structure::create_sheet(
                &mut store,
                &workbook_path,
                &wb_dir,
                &wb_rels_path,
                name,
                tab_color.as_deref(),
            )?,
            SheetOp::Rename { old_name, new_name } => structure::rename_sheet(
                &mut store,
                &workbook_path,
                &wb_dir,
                old_name,
                new_name,
            )?,
            SheetOp::Delete { name } => structure::delete_sheet(
                &mut store,
                &workbook_path,
                &wb_dir,
                &wb_rels_path,
                name,
            )?,
        }
    }
    for op in &patch.row_col_ops {
        row_col::apply_row_col_op(&mut store, &workbook_path, &wb_dir, &wb_rels_path, op)?;
    }

    // Re-read after structural ops: sheet list and rels may have changed.
    let workbook_xml = store.read(&workbook_path)?;
    let wb_rels_xml = store.read(&wb_rels_path)?;
    let sheet_paths = workbook_xml::sheet_part_paths(&workbook_xml, &wb_rels_xml)?;

    // Lazily-created styles editor, shared across sheets so interned indices
    // are consistent.
    let mut styles_editor: Option<styles::StylesEditor> = None;
    let styles_path = format!("{wb_dir}styles.xml");

    for sp in &patch.sheets {
        if sp.is_empty() {
            continue;
        }
        let part = sheet_paths
            .iter()
            .find(|(name, _)| name == &sp.name)
            .map(|(_, p)| p.clone())
            .ok_or_else(|| PatchError::SheetNotFound(sp.name.clone()))?;
        let sheet_bytes = store.read(&part)?;

        // --- scan ---
        let style_targets: HashSet<(u32, u32)> =
            sp.styles.iter().map(|s| (s.r, s.c)).collect();
        let scan = sheet_xml::scan_sheet(&sheet_bytes, &style_targets)?;

        // --- content cells ---
        let mut rows: BTreeMap<u32, RowWork> = BTreeMap::new();
        let mut content_cells: HashSet<(u32, u32)> = HashSet::new();
        for cp in &sp.cells {
            rows.entry(cp.r).or_default().cells.insert(
                cp.c,
                CellAction {
                    style: None,
                    content: cell_content(cp)?,
                },
            );
            content_cells.insert((cp.r, cp.c));
        }

        // --- shared formula materialization ---
        // Any group with a patched member must be fully materialized: the
        // remaining members get their translated plain formulas.
        for (si, group) in &scan.shared_groups {
            let touched = group.members.iter().any(|m| content_cells.contains(m));
            if !touched {
                continue;
            }
            if group.formula.is_empty() {
                return Err(PatchError::SharedFormulaMissing(*si));
            }
            for member in &group.members {
                if content_cells.contains(member) {
                    continue;
                }
                let dr = member.0 as i64 - group.master.0 as i64;
                let dc = member.1 as i64 - group.master.1 as i64;
                let translated = refs::shift_formula(&group.formula, dr, dc)?;
                rows.entry(member.0).or_default().cells.insert(
                    member.1,
                    CellAction {
                        style: None,
                        content: CellContent::MaterializeFormula(translated),
                    },
                );
            }
        }

        // --- style patches ---
        if !sp.styles.is_empty() {
            if styles_editor.is_none() {
                styles_editor = Some(styles::StylesEditor::parse(store.read(&styles_path)?)?);
            }
            let editor = styles_editor.as_mut().unwrap();
            for stp in &sp.styles {
                let base_xf = scan
                    .style_bases
                    .get(&(stp.r, stp.c))
                    .copied()
                    .unwrap_or(0);
                let new_s = editor.intern(base_xf, stp)?;
                let cell = rows
                    .entry(stp.r)
                    .or_default()
                    .cells
                    .entry(stp.c)
                    .or_insert(CellAction {
                        style: None,
                        content: CellContent::Keep,
                    });
                cell.style = Some(new_s);
            }
        }

        // --- row heights / visibility ---
        for rh in &sp.row_heights {
            rows.entry(rh.r).or_default().height = Some(rh.pts);
        }
        for rv in &sp.hidden_rows {
            rows.entry(rv.r).or_default().hidden = Some(rv.hidden);
        }

        // --- column overrides ---
        let mut col_overrides: BTreeMap<u32, ColOverride> = BTreeMap::new();
        for cw in &sp.col_widths {
            col_overrides.entry(cw.c).or_default().width_chars = Some(cw.chars);
        }
        for cv in &sp.hidden_cols {
            col_overrides.entry(cv.c).or_default().hidden = Some(cv.hidden);
        }

        let ops = SheetOps {
            rows,
            col_overrides,
            merge_ops: sp.merges.iter().map(|m| (m.range.clone(), m.merge)).collect(),
            freeze: sp.freeze.as_ref().map(|f| (f.rows, f.cols)),
            auto_filter: sp.auto_filter.clone(),
        };
        if ops.is_empty() {
            continue;
        }
        let rewritten = sheet_xml::rewrite_sheet(&sheet_bytes, &ops, &scan)?;
        store.write(&part, rewritten);
    }

    // --- calc chain + fullCalcOnLoad ---
    if patch.has_content_changes() {
        if let Some(calc_target) = find_rel_target(&wb_rels_xml, "/calcChain", &wb_dir) {
            if store.exists(&calc_target) {
                store.remove(&calc_target);
                store.write(&wb_rels_path, workbook_xml::strip_calc_chain_rel(&wb_rels_xml)?);
                let ct = store.read("[Content_Types].xml")?;
                let stripped =
                    workbook_xml::strip_content_type_override(&ct, &format!("/{calc_target}"))?;
                store.write("[Content_Types].xml", stripped);
            }
        }
    }

    // --- workbook.xml (fullCalcOnLoad, defined names) ---
    if patch.has_content_changes() || !patch.defined_names.is_empty() {
        let current = store.read(&workbook_path)?;
        let out = workbook_xml::patch_workbook_xml(
            &current,
            patch.has_content_changes(),
            &patch.defined_names,
        )?;
        store.write(&workbook_path, out);
    }

    // --- styles.xml ---
    if let Some(editor) = &styles_editor {
        if editor.is_dirty() {
            store.write(&styles_path, editor.serialize()?);
        }
    }

    // --- validity gate ---
    // Everything the patch rewrote is re-checked against the Excel-strict
    // invariants before a single byte reaches disk. A failure here means a
    // generator bug — fail the save, never write a file Excel would repair.
    validate::validate_store(&mut store, &workbook_path, &wb_rels_path)?;

    store.finish()
}

fn cell_content(cp: &CellPatch) -> Result<CellContent, PatchError> {
    if let Some(f) = &cp.f {
        let trimmed = f.trim_start_matches('=');
        if trimmed.is_empty() {
            return Err(PatchError::BadValue("empty formula".into()));
        }
        // A `v` next to `f` is the frontend's evaluated result, cached into
        // the file so reopening in-app doesn't force a full recalculation.
        // An unencodable cached value is dropped, never an error: the cache
        // is a perf optimization, the formula is the content.
        let cached = match &cp.v {
            Some(CellValue::Number { n }) if n.is_finite() => Some(CachedValue::Number(*n)),
            Some(CellValue::Str { s }) => Some(CachedValue::Str(s.clone())),
            Some(CellValue::Bool { b }) => Some(CachedValue::Bool(*b)),
            Some(CellValue::Error { e }) => Some(CachedValue::Error(e.clone())),
            _ => None,
        };
        return Ok(CellContent::Formula(trimmed.to_string(), cached));
    }
    if let Some(v) = &cp.v {
        return Ok(match v {
            CellValue::Number { n } => {
                if !n.is_finite() {
                    return Err(PatchError::BadValue(format!("non-finite number {n}")));
                }
                CellContent::Number(*n)
            }
            CellValue::Str { s } => CellContent::Text(s.clone()),
            CellValue::Bool { b } => CellContent::Bool(*b),
            CellValue::Error { e } => CellContent::ErrorVal(e.clone()),
        });
    }
    if cp.clear {
        return Ok(CellContent::Clear);
    }
    Err(PatchError::BadValue(format!(
        "cell patch at r{} c{} has no content",
        cp.r, cp.c
    )))
}

/// Find a Relationship whose Type ends with `type_suffix`; resolve its
/// Target relative to `base_dir` ("" for package root, "xl/" for workbook
/// rels).
fn find_rel_target(rels_xml: &[u8], type_suffix: &str, base_dir: &str) -> Option<String> {
    let mut r = workbook_xml::reader(rels_xml);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Empty(e)) | Ok(quick_xml::events::Event::Start(e)) => {
                let is_rel = {
                    let name = e.name();
                    workbook_xml::local_name(name.as_ref()) == b"Relationship"
                };
                if is_rel {
                    let ty = workbook_xml::attr_value(&e, b"Type");
                    let target = workbook_xml::attr_value(&e, b"Target");
                    if let (Some(ty), Some(target)) = (ty, target) {
                        if ty.ends_with(type_suffix) {
                            let resolved = if let Some(stripped) = target.strip_prefix('/') {
                                stripped.to_string()
                            } else {
                                format!("{base_dir}{target}")
                            };
                            return Some(resolved);
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => return None,
            Ok(_) => {}
            Err(_) => return None,
        }
        buf.clear();
    }
}

/// "xl/workbook.xml" → ("xl/", "workbook.xml")
fn split_part_path(path: &str) -> (String, String) {
    match path.rfind('/') {
        Some(i) => (path[..=i].to_string(), path[i + 1..].to_string()),
        None => (String::new(), path.to_string()),
    }
}
