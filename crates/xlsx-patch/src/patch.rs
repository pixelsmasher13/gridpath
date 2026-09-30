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
}

impl Patch {
    /// An empty patch must short-circuit to "return the original bytes":
    /// the safest possible save is the one that doesn't rewrite anything.
    pub fn is_empty(&self) -> bool {
        self.defined_names.is_empty()
            && self.sheet_ops.is_empty()
            && self.row_col_ops.is_empty()
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
                .any(|op| !matches!(op, SheetOp::Create { .. }))
    }
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
    /// `Some(Some(range))` sets the AutoFilter, `Some(None)` clears it,
    /// `None` leaves it untouched.
    #[serde(default, with = "double_option")]
    pub auto_filter: Option<Option<String>>,
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
