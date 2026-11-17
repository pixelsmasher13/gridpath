//! Wire format for surgical workbook patches.
//!
//! The frontend diffs its live cell model against the load-time copy and
//! sends one of these per save. Field names are camelCase to match the
//! TypeScript side. Rows/columns are 0-based here (matching the frontend
//! model); conversion to 1-based A1 happens at the XML layer.

use serde::Deserialize;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Patch {
    /// Bump when the shape changes; the backend rejects versions it doesn't
    /// know rather than guessing.
    pub version: u32,
    #[serde(default)]
    pub sheets: Vec<SheetPatch>,
    /// Workbook-scoped defined names to create or replace.
    #[serde(default)]
    pub defined_names: Vec<DefinedName>,
    /// Sheet create/rename/delete, applied in order BEFORE cell patches
    /// (cell patches address sheets by their live — post-op — names).
    #[serde(default)]
    pub sheet_ops: Vec<SheetOp>,
    /// Row/column insert/delete, applied in order after sheet ops and
    /// before cell patches (cell coordinates are post-shift).
    #[serde(default)]
    pub row_col_ops: Vec<RowColOp>,
    /// Links to other workbooks: the parts behind `[n]Sheet!A1` formulas
    /// (see `external_links.rs`). An entry creates link `index` or brings
    /// the one the file already has up to date.
    #[serde(default)]
    pub external_links: Vec<ExternalLinkPatch>,
    /// Charts made in this session, as they are now (their ranges and place
    /// already follow every row, column and sheet change of the patch).
    /// Written last: the chart part, its place in the sheet's drawing, and
    /// the data it plots as the engine has it.
    #[serde(default)]
    pub charts: Vec<crate::charts::Chart>,
}

impl Patch {
    /// An empty patch must short-circuit to "return the original bytes":
    /// the safest possible save is the one that doesn't rewrite anything.
    pub fn is_empty(&self) -> bool {
        self.defined_names.is_empty()
            && self.sheet_ops.is_empty()
            && self.row_col_ops.is_empty()
            && self.external_links.is_empty()
            && self.charts.is_empty()
            && self.sheets.iter().all(SheetPatch::is_empty)
    }

    /// True when any cell content changed — that's what obligates dropping
    /// calcChain and setting fullCalcOnLoad so Excel recalculates.
    /// Row/col shifts and sheet rename/delete rewrite formulas, so they
    /// count; a pure sheet create does not.
    pub fn has_content_changes(&self) -> bool {
        self.sheets.iter().any(|s| !s.cells.is_empty())
            || !self.row_col_ops.is_empty()
            || self
                .sheet_ops
                .iter()
                .any(|op| !matches!(op, SheetOp::Create { .. } | SheetOp::SetVisibility { .. } | SheetOp::SetTabColor { .. }))
    }
}

/// One other workbook this one reads.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ExternalLinkPatch {
    /// The `n` of `[n]Sheet!A1`: 1-based position in workbook.xml's
    /// `<externalReferences>`. An index the file already has is updated in
    /// place; a new one must be the next free index.
    pub index: u32,
    /// The other file as the link's relationship spells it (`NOW.xlsx`,
    /// `../models/NOW%202026.xlsx`, `file:///…`). Only written for a new
    /// link: an existing link keeps the path it was saved with.
    #[serde(default)]
    pub target: String,
    #[serde(default)]
    pub sheets: Vec<ExternalSheetPatch>,
}

/// One sheet of the other workbook, with the last values of the cells this
/// workbook reads from it. Excel shows these until the link is refreshed.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ExternalSheetPatch {
    pub name: String,
    /// What this workbook reads on the sheet, as 0-based inclusive
    /// `[r1, c1, r2, c2]`. Values the file kept inside these are dropped:
    /// `cells` is everything that is there now. Kept values outside them
    /// (read by something this patch does not know about) stay.
    #[serde(default)]
    pub areas: Vec<[u32; 4]>,
    /// The non-empty cells inside `areas`.
    #[serde(default)]
    pub cells: Vec<ExternalCellPatch>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ExternalCellPatch {
    pub r: u32,
    pub c: u32,
    pub v: CellValue,
}

/// Sheet-level structure operation.
#[derive(Debug, Deserialize)]
#[serde(tag = "op", rename_all = "camelCase")]
pub enum SheetOp {
    #[serde(rename_all = "camelCase")]
    Create {
        name: String,
        /// CSS hex like "#RRGGBB" for the tab stripe.
        #[serde(default)]
        tab_color: Option<String>,
    },
    #[serde(rename_all = "camelCase")]
    Rename { old_name: String, new_name: String },
    #[serde(rename_all = "camelCase")]
    Delete { name: String },
    /// Hide or show a sheet (workbook.xml `state` attribute). Not a content change.
    #[serde(rename_all = "camelCase")]
    SetVisibility { name: String, hidden: bool },
    /// Set (CSS hex) or clear (null) a sheet's tab colour. Not a content change.
    #[serde(rename_all = "camelCase")]
    SetTabColor { name: String, color: Option<String> },
}

/// Row/column structure operation. Indices are 0-based (frontend model).
#[derive(Debug, Deserialize)]
#[serde(tag = "op", rename_all = "camelCase")]
pub enum RowColOp {
    #[serde(rename_all = "camelCase")]
    InsertRows { sheet: String, before: u32, count: u32 },
    #[serde(rename_all = "camelCase")]
    DeleteRows { sheet: String, start: u32, count: u32 },
    #[serde(rename_all = "camelCase")]
    InsertColumns { sheet: String, before: u32, count: u32 },
    #[serde(rename_all = "camelCase")]
    DeleteColumns { sheet: String, start: u32, count: u32 },
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SheetPatch {
    /// Sheet name exactly as it appears in the workbook (not the part path).
    pub name: String,
    #[serde(default)]
    pub cells: Vec<CellPatch>,
    #[serde(default)]
    pub styles: Vec<StylePatch>,
    #[serde(default)]
    pub col_widths: Vec<ColWidth>,
    #[serde(default)]
    pub row_heights: Vec<RowHeight>,
    #[serde(default)]
    pub hidden_rows: Vec<RowVisibility>,
    #[serde(default)]
    pub hidden_cols: Vec<ColVisibility>,
    #[serde(default)]
    pub merges: Vec<MergeOp>,
    #[serde(default)]
    pub freeze: Option<Freeze>,
    /// `Some(Some(filter))` sets the AutoFilter (a bare range string, or a
    /// range with per-column criteria), `Some(None)` clears it, `None`
    /// leaves it untouched.
    #[serde(default, with = "double_option")]
    pub auto_filter: Option<Option<AutoFilterPatch>>,
}

/// An AutoFilter: just its range, or the range plus column criteria.
#[derive(Debug, Clone, Deserialize)]
#[serde(untagged)]
pub enum AutoFilterPatch {
    Range(String),
    Spec(AutoFilterSpec),
}

impl AutoFilterPatch {
    pub fn range(&self) -> &str {
        match self {
            AutoFilterPatch::Range(r) => r,
            AutoFilterPatch::Spec(s) => &s.r#ref,
        }
    }
    pub fn columns(&self) -> &[FilterColumnPatch] {
        match self {
            AutoFilterPatch::Range(_) => &[],
            AutoFilterPatch::Spec(s) => &s.columns,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AutoFilterSpec {
    pub r#ref: String,
    #[serde(default)]
    pub columns: Vec<FilterColumnPatch>,
}

/// One `<filterColumn>`: a value list (`values` + `blank`) and/or custom
/// comparisons. `colId` is the 0-based offset within the range.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FilterColumnPatch {
    pub col_id: u32,
    #[serde(default)]
    pub values: Option<Vec<String>>,
    #[serde(default)]
    pub blank: bool,
    #[serde(default)]
    pub custom: Option<CustomFilterPatch>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CustomFilterPatch {
    #[serde(default)]
    pub and: bool,
    pub items: Vec<CustomFilterItem>,
}

/// `operator` is one of equal, notEqual, greaterThan, greaterThanOrEqual,
/// lessThan, lessThanOrEqual (ST_FilterOperator).
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CustomFilterItem {
    pub operator: String,
    pub val: String,
}

impl SheetPatch {
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
            && self.styles.is_empty()
            && self.col_widths.is_empty()
            && self.row_heights.is_empty()
            && self.hidden_rows.is_empty()
            && self.hidden_cols.is_empty()
            && self.merges.is_empty()
            && self.freeze.is_none()
            && self.auto_filter.is_none()
    }
}

/// One cell content change. `f` alone is a formula, `v` alone a literal,
/// `clear` alone clears; `f` + `v` is a formula WITH its evaluated result,
/// persisted as a cached `<v>` so reopening in-app skips recalculation.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellPatch {
    pub r: u32,
    pub c: u32,
    /// Formula WITHOUT the leading `=`. Persisted as `<f>`; Excel still
    /// recalculates on open (fullCalcOnLoad) whether or not a cached value
    /// accompanies it.
    #[serde(default)]
    pub f: Option<String>,
    /// Literal value when `f` is absent; cached evaluated result when `f`
    /// is present.
    #[serde(default)]
    pub v: Option<CellValue>,
    /// Clear contents, keep the cell's style.
    #[serde(default)]
    pub clear: bool,
    /// Alone: the new evaluated result of the formula already in this cell
    /// (a dependent that recalculated). Only the cached `<v>` and its type
    /// change; the `<f>` stays byte-for-byte. Ignored for a cell that is not
    /// a formula in the file, and never creates a cell.
    #[serde(default)]
    pub cache: Option<CellValue>,
}

#[derive(Debug, Deserialize)]
#[serde(tag = "t")]
pub enum CellValue {
    #[serde(rename = "n")]
    Number { n: f64 },
    #[serde(rename = "s")]
    Str { s: String },
    #[serde(rename = "b")]
    Bool { b: bool },
    /// Literal error value the user typed (e.g. `#N/A`).
    #[serde(rename = "e")]
    Error { e: String },
}

/// Partial format override for one cell. Mirrors the frontend's
/// `CellFormatShape` + background + borders: only the fields present are
/// changed; everything else is inherited from the cell's current style.
#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct StylePatch {
    pub r: u32,
    pub c: u32,
    #[serde(default)]
    pub bold: Option<bool>,
    #[serde(default)]
    pub italic: Option<bool>,
    #[serde(default)]
    pub underline: Option<bool>,
    #[serde(default)]
    pub strike: Option<bool>,
    /// CSS hex like "#RRGGBB".
    #[serde(default)]
    pub font_color: Option<String>,
    /// `Some(None)` clears the fill, `Some(Some(color))` sets it.
    #[serde(default, with = "double_option")]
    pub background_color: Option<Option<String>>,
    #[serde(default)]
    pub font_size: Option<f64>,
    #[serde(default)]
    pub font_family: Option<String>,
    #[serde(default)]
    pub horizontal_align: Option<String>,
    #[serde(default)]
    pub vertical_align: Option<String>,
    #[serde(default)]
    pub number_format: Option<String>,
    #[serde(default)]
    pub wrap_text: Option<bool>,
    #[serde(default)]
    pub indent: Option<u32>,
    /// Border sides: each present side replaces that side; `null` clears it.
    #[serde(default)]
    pub borders: Option<BordersPatch>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct BordersPatch {
    #[serde(default, with = "double_option")]
    pub top: Option<Option<BorderSide>>,
    #[serde(default, with = "double_option")]
    pub bottom: Option<Option<BorderSide>>,
    #[serde(default, with = "double_option")]
    pub left: Option<Option<BorderSide>>,
    #[serde(default, with = "double_option")]
    pub right: Option<Option<BorderSide>>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct BorderSide {
    /// OOXML border style token: thin, medium, thick, dashed, dotted, double…
    pub style: String,
    /// CSS hex like "#RRGGBB"; defaults to black when absent.
    #[serde(default)]
    pub color: Option<String>,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ColWidth {
    pub c: u32,
    /// Width in Excel character units (what the XML stores), NOT pixels.
    pub chars: f64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RowHeight {
    pub r: u32,
    /// Height in points (what the XML stores), NOT pixels.
    pub pts: f64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RowVisibility {
    pub r: u32,
    pub hidden: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ColVisibility {
    pub c: u32,
    pub hidden: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MergeOp {
    /// A1 range like "B2:D2".
    pub range: String,
    /// true = merge, false = unmerge.
    pub merge: bool,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Freeze {
    pub rows: u32,
    pub cols: u32,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DefinedName {
    pub name: String,
    /// Formula-style reference, e.g. `'DCF'!$C$4`.
    pub r#ref: String,
}

/// serde helper distinguishing "absent" from "explicitly null" so patches
/// can express "clear this" separately from "leave this alone".
mod double_option {
    use serde::{Deserialize, Deserializer};

    pub fn deserialize<'de, T, D>(de: D) -> Result<Option<Option<T>>, D::Error>
    where
        T: Deserialize<'de>,
        D: Deserializer<'de>,
    {
        Ok(Some(Option::<T>::deserialize(de)?))
    }
}
