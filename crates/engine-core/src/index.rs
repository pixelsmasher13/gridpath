//! Structural workbook index built straight from the IronCalc model — the
//! Rust twin of `packages/gridpath/src/core/workbookIndex.ts`
//! (`buildWorkbookIndex`). Same heuristics, same JSON shape, so the
//! TypeScript consumers (`describeWorkbookPayload`, `findRowsInIndex`) run
//! unchanged over either source.
//!
//! Why a second implementation: the TypeScript builder walks a Univer
//! snapshot — 2.3M JavaScript cell objects on the large benchmark model, 2.2 s on
//! the main thread after every write the agent makes. Here the walk is over
//! the model the engine already holds, and the per-sheet memo
//! (`IndexCache`) means only sheets that changed are rebuilt.
//!
//! Conventions in the output (dictated by the TypeScript types):
//! * `row`, `startRow`, `endRow`, `headerRow` are 1-based;
//! * `sampleCol`, `startCol`, `endCol` are 0-based;
//! * header keys are column letters.

use ironcalc_base::cell::CellValue;
use ironcalc_base::types::{Cell, FormulaValue, SheetState, SpillValue};
use ironcalc_base::Model;
use serde::Serialize;
use std::collections::HashMap;

/// Formula-run entries kept per sheet — backstop against pathological sheets.
pub const MAX_FORMULA_RUNS: usize = 20_000;
/// Sections kept per sheet at build time.
pub const MAX_SECTIONS: usize = 500;
/// Header cells kept per sheet.
pub const MAX_HEADERS: usize = 150;
/// Rows examined when hunting for the header row.
pub const HEADER_SCAN_ROWS: usize = 50;
/// Row labels are clipped to this many characters.
pub const MAX_LABEL_CHARS: usize = 60;
/// Header texts are clipped to this many characters.
pub const MAX_HEADER_CHARS: usize = 40;

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct RowEntry {
    pub row: i32,
    pub label: Option<String>,
    pub values: u32,
    pub formulas: u32,
    pub bold: bool,
    pub sample_col: i32,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetSection {
    pub start_row: i32,
    pub end_row: i32,
    pub title: Option<String>,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct FormulaRun {
    pub row: i32,
    pub start_col: i32,
    pub end_col: i32,
    pub formula: String,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct SheetIndex {
    pub name: String,
    pub hidden: bool,
    pub used_range: Option<String>,
    pub header_row: Option<i32>,
    /// Column letter → header text. A `BTreeMap` would sort "AA" before
    /// "B"; insertion order (column order) is what the TypeScript side
    /// produces, so keep a Vec of pairs and serialise as an object.
    #[serde(serialize_with = "serialize_pairs")]
    pub headers: Vec<(String, String)>,
    pub rows: Vec<RowEntry>,
    pub sections: Vec<SheetSection>,
    pub formula_runs: Vec<FormulaRun>,
    pub formula_runs_truncated: bool,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct DefinedName {
    pub name: String,
    #[serde(rename = "ref")]
    pub reference: String,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookIndex {
    pub sheets: Vec<SheetIndex>,
    pub defined_names: Vec<DefinedName>,
}

fn serialize_pairs<S: serde::Serializer>(pairs: &[(String, String)], s: S) -> Result<S::Ok, S::Error> {
    use serde::ser::SerializeMap;
    let mut map = s.serialize_map(Some(pairs.len()))?;
    for (k, v) in pairs {
        map.serialize_entry(k, v)?;
    }
    map.end()
}

// ---------------------------------------------------------------------------
// Shared helpers (ports of the TypeScript ones, kept byte-for-byte in spirit)
// ---------------------------------------------------------------------------

/// 0-based column index → letters.
pub fn col_letters(c0: i32) -> String {
    let mut n = c0;
    let mut s = Vec::new();
    loop {
        s.push(b'A' + (n % 26) as u8);
        n = n / 26 - 1;
        if n < 0 {
            break;
        }
    }
    s.reverse();
    String::from_utf8(s).unwrap_or_default()
}

/// Letters → 0-based column index (None when not a plain A..XFD token).
pub fn letters_to_col(s: &str) -> Option<i32> {
    if s.is_empty() || s.len() > 3 {
        return None;
    }
    let mut n: i32 = 0;
    for ch in s.chars() {
        let u = ch.to_ascii_uppercase();
        if !u.is_ascii_uppercase() {
            return None;
        }
        n = n * 26 + (u as i32 - 'A' as i32 + 1);
    }
    if n <= 16_384 {
        Some(n - 1)
    } else {
        None
    }
}

/// Is this text usable as a human row label? ≥2 chars after trimming and
/// not leading with number/formula/currency punctuation (rejects "Y",
/// "2024A", numbers, formula-ish text).
pub fn is_label_text(s: &str) -> bool {
    let t = s.trim();
    if t.chars().count() < 2 {
        return false;
    }
    !matches!(t.chars().next(), Some('-' | '+' | '#' | '0'..='9' | '.' | '(' | '$' | '%' | '='))
}

fn clip_chars(s: &str, n: usize) -> String {
    s.chars().take(n).collect()
}

/// Does this number format render a date/time? (y/m/d/h tokens outside
/// quoted literals and bracketed sections.)
pub fn is_date_format(nf: &str) -> bool {
    let mut in_quote = false;
    let mut in_bracket = false;
    for ch in nf.chars() {
        match ch {
            '"' if !in_bracket => in_quote = !in_quote,
            '[' if !in_quote => in_bracket = true,
            ']' if !in_quote => in_bracket = false,
            c if !in_quote && !in_bracket && matches!(c.to_ascii_lowercase(), 'y' | 'm' | 'd' | 'h') => return true,
            _ => {}
        }
    }
    false
}

/// Excel 1900-system serial → "YYYY-MM-DD" (same validity window as the
/// TypeScript helper: 1..=2958465, i.e. up to 9999-12-31).
pub fn serial_to_iso_date(serial: f64) -> Option<String> {
    if !serial.is_finite() || serial < 1.0 || serial > 2_958_465.0 {
        return None;
    }
    // Days since 1899-12-30 (so that serial 1 = 1900-01-01 and the phantom
    // 29 Feb 1900 at serial 60 maps like Excel's display does for ≥61).
    let days = serial.round() as i64;
    // Civil-from-days (Howard Hinnant), epoch 1970-01-01 = serial 25569.
    let z = days - 25_569 + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = if m <= 2 { y + 1 } else { y };
    Some(format!("{y:04}-{m:02}-{d:02}"))
}

/// A cell's value without cloning shared strings.
#[allow(dead_code)]
enum Val<'a> {
    None,
    Num(f64),
    Bool(bool),
    Str(&'a str),
    Owned(String),
}

impl<'a> Val<'a> {
    fn as_str(&self) -> Option<&str> {
        match self {
            Val::Str(s) => Some(s),
            Val::Owned(s) => Some(s.as_str()),
            _ => None,
        }
    }
    fn is_present(&self) -> bool {
        match self {
            Val::None => false,
            Val::Str(s) => !s.is_empty(),
            Val::Owned(s) => !s.is_empty(),
            _ => true,
        }
    }
}

fn value_of<'a>(m: &'a Model, sheet: u32, r: i32, c: i32, cell: &'a Cell) -> Val<'a> {
    let shared = &m.workbook.shared_strings;
    match cell {
        Cell::EmptyCell { .. } => Val::None,
        Cell::BooleanCell { v, .. } => Val::Bool(*v),
        Cell::NumberCell { v, .. } => Val::Num(*v),
        Cell::SharedString { si, .. } => shared.get(*si as usize).map(|s| Val::Str(s.as_str())).unwrap_or(Val::None),
        Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => match v {
            FormulaValue::Boolean(b) => Val::Bool(*b),
            FormulaValue::Number(n) => Val::Num(*n),
            FormulaValue::Text(t) => Val::Str(t.as_str()),
            FormulaValue::Unevaluated => Val::None,
            FormulaValue::Error { .. } => localized(m, sheet, r, c),
        },
        Cell::SpillCell { v, .. } => match v {
            SpillValue::Boolean(b) => Val::Bool(*b),
            SpillValue::Number(n) => Val::Num(*n),
            SpillValue::Text(t) => Val::Str(t.as_str()),
            SpillValue::Error(_) => localized(m, sheet, r, c),
        },
        Cell::ErrorCell { .. } => localized(m, sheet, r, c),
    }
}

/// Errors render through the model so the text matches what the grid
/// shows ("#N/A", "#REF!"...). Rare enough that the clone is irrelevant.
fn localized(m: &Model, sheet: u32, r: i32, c: i32) -> Val<'static> {
    match m.get_cell_value_by_index(sheet, r, c) {
        Ok(CellValue::String(s)) => Val::Owned(s),
        Ok(CellValue::Number(n)) => Val::Num(n),
        Ok(CellValue::Boolean(b)) => Val::Bool(b),
        _ => Val::None,
    }
}

/// Per-build memo of style facts keyed by style index: bold flag and
/// number format. `get_style_for_cell` assembles a full `Style` (font,
/// fill, border clones) so it must run once per distinct style, not per
/// cell.
#[derive(Default)]
pub struct StyleFacts {
    bold: HashMap<i32, bool>,
    num_fmt: HashMap<i32, String>,
}

impl StyleFacts {
    fn fill(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) {
        if self.bold.contains_key(&s) {
            return;
        }
        match m.get_style_for_cell(sheet, r, c) {
            Ok(st) => {
                self.bold.insert(s, st.font.b);
                self.num_fmt.insert(s, st.num_fmt);
            }
            Err(_) => {
                self.bold.insert(s, false);
                self.num_fmt.insert(s, "General".to_string());
            }
        }
    }
    pub fn is_bold(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) -> bool {
        self.fill(m, sheet, r, c, s);
        self.bold[&s]
    }
    pub fn num_fmt(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) -> &str {
        self.fill(m, sheet, r, c, s);
        &self.num_fmt[&s]
    }
    pub fn is_date(&mut self, m: &Model, sheet: u32, r: i32, c: i32, s: i32) -> bool {
        is_date_format(self.num_fmt(m, sheet, r, c, s))
    }
}

/// Header-cell rendering: strings verbatim (clipped), date serials as ISO,
/// years as text; anything else is not header-like.
fn header_text_of(v: &Val, date_fmt: bool) -> Option<String> {
    if let Some(s) = v.as_str() {
        let t = s.trim();
        return if t.is_empty() { None } else { Some(clip_chars(t, MAX_HEADER_CHARS)) };
    }
    match v {
        Val::Num(n) => {
            if date_fmt {
                if let Some(iso) = serial_to_iso_date(*n) {
                    return Some(iso);
                }
            }
            if n.fract() == 0.0 && (1900.0..=2100.0).contains(n) {
                Some(format!("{}", *n as i64))
            } else {
                None
            }
        }
        _ => None,
    }
}

/// Sorted (row, sorted columns) view of a sheet's cells.
fn sorted_rows(ws: &ironcalc_base::types::Worksheet) -> Vec<(i32, Vec<i32>)> {
    let mut rows: Vec<(i32, Vec<i32>)> = ws
        .sheet_data
        .iter()
        .map(|(r, cols)| {
            let mut cs: Vec<i32> = cols.keys().copied().collect();
            cs.sort_unstable();
            (*r, cs)
        })
        .collect();
    rows.sort_unstable_by_key(|(r, _)| *r);
    rows
}

/// The header (period) row of a sheet, by the index's rule: the topmost of
/// the first `HEADER_SCAN_ROWS` rows whose cells beyond column A are mostly
/// text / years / dates (≥3 such cells and ≥60% of the non-empty ones).
/// Returns the 1-based row and its column-letter → text map in column order.
pub fn detect_header_row(m: &Model, sheet: u32, facts: &mut StyleFacts) -> Option<(i32, Vec<(String, String)>)> {
    let ws = m.workbook.worksheets.get(sheet as usize)?;
    let rows = sorted_rows(ws);
    for (r, cols) in rows.iter().take(HEADER_SCAN_ROWS) {
        let cells = &ws.sheet_data[r];
        let mut non_empty = 0usize;
        let mut header_like = 0usize;
        let mut candidate: Vec<(String, String)> = Vec::new();
        for c in cols {
            if *c == 1 {
                continue; // label column
            }
            let cell = &cells[c];
            let v = value_of(m, sheet, *r, *c, cell);
            let has_content = v.is_present() || cell.get_formula().is_some();
            if !has_content {
                continue;
            }
            non_empty += 1;
            let date_fmt = matches!(v, Val::Num(_)) && facts.is_date(m, sheet, *r, *c, cell.get_style());
            if let Some(text) = header_text_of(&v, date_fmt) {
                header_like += 1;
                if candidate.len() < MAX_HEADERS {
                    candidate.push((col_letters(*c - 1), text));
                }
            }
        }
        if header_like >= 3 && (header_like as f64) >= 0.6 * (non_empty as f64) {
            return Some((*r, candidate));
        }
    }
    None
}

/// Build the index of one sheet.
pub fn build_sheet_index(m: &Model, sheet: u32) -> Result<SheetIndex, String> {
    let ws = m.workbook.worksheets.get(sheet as usize).ok_or_else(|| format!("sheet index {sheet} out of range"))?;
    let mut facts = StyleFacts::default();
    let rows_sorted = sorted_rows(ws);

    let mut rows: Vec<RowEntry> = Vec::new();
    let mut formula_runs: Vec<FormulaRun> = Vec::new();
    let mut formula_runs_truncated = false;
    let mut max_used_row = 0i32; // 1-based; 0 = none
    let mut max_used_col = 0i32;

    for (r, cols) in &rows_sorted {
        let cells = &ws.sheet_data[r];
        let mut values = 0u32;
        let mut formulas = 0u32;
        let mut bold = false;
        let mut last_value_col = -1i32;
        let mut last_formula_col = -1i32;
        let mut label: Option<String> = None;

        // Formula-run compression: consecutive columns whose formulas share
        // the same R1C1 text (IronCalc's `shared_formulas` entry) are exact
        // column translations of each other — the same test the TypeScript
        // builder makes by shifting the A1 text one column at a time.
        let mut run_start: Option<(i32, &str)> = None; // (start col 1-based, r1c1)
        let mut run_prev_col = 0i32;
        let close_run = |start: i32, end: i32, formula_runs: &mut Vec<FormulaRun>, truncated: &mut bool| {
            if formula_runs.len() < MAX_FORMULA_RUNS {
                let f = m.get_cell_formula(sheet, *r, start).ok().flatten().unwrap_or_default();
                let f = f.strip_prefix('=').unwrap_or(&f);
                formula_runs.push(FormulaRun { row: *r, start_col: start - 1, end_col: end - 1, formula: format!("={f}") });
            } else {
                *truncated = true;
            }
        };

        for c in cols {
            let cell = &cells[c];
            let f_index = cell.get_formula();
            let v = value_of(m, sheet, *r, *c, cell);
            let has_f = f_index.is_some();
            let has_v = v.is_present();
            // Runs are interrupted by any non-formula column, including a
            // gap in the column sequence (handled below) and literal cells.
            let r1c1 = f_index.and_then(|fi| ws.shared_formulas.get(fi as usize)).map(|s| s.as_str());
            match (run_start, r1c1) {
                (Some((_, prev)), Some(cur)) if *c == run_prev_col + 1 && prev == cur => {}
                (Some((start, _)), Some(cur)) => {
                    close_run(start, run_prev_col, &mut formula_runs, &mut formula_runs_truncated);
                    run_start = Some((*c, cur));
                }
                (Some((start, _)), None) => {
                    close_run(start, run_prev_col, &mut formula_runs, &mut formula_runs_truncated);
                    run_start = None;
                }
                (None, Some(cur)) => run_start = Some((*c, cur)),
                (None, None) => {}
            }
            run_prev_col = *c;

            if !has_f && !has_v {
                continue;
            }
            if *r > max_used_row {
                max_used_row = *r;
            }
            if *c > max_used_col {
                max_used_col = *c;
            }
            if has_f {
                formulas += 1;
                last_formula_col = *c;
            } else {
                values += 1;
                last_value_col = *c;
            }
            // Formula cells' computed values count as label candidates too
            // — a label pulled via formula (`=Assumptions!A5`) is still the
            // label.
            if label.is_none() {
                if let Some(s) = v.as_str() {
                    if is_label_text(s) {
                        label = Some(clip_chars(s.trim(), MAX_LABEL_CHARS));
                    }
                }
            }
            if !bold && facts.is_bold(m, sheet, *r, *c, cell.get_style()) {
                bold = true;
            }
        }
        if let Some((start, _)) = run_start {
            close_run(start, run_prev_col, &mut formula_runs, &mut formula_runs_truncated);
        }

        if values + formulas == 0 {
            continue;
        }
        let sample = if last_formula_col >= 1 { last_formula_col } else { last_value_col };
        rows.push(RowEntry {
            row: *r,
            label,
            values,
            formulas,
            bold,
            sample_col: if sample >= 1 { sample - 1 } else { -1 },
        });
    }

    // Header (period) row.
    let (header_row, headers) = match detect_header_row(m, sheet, &mut facts) {
        Some((r, h)) => (Some(r), h),
        None => (None, Vec::new()),
    };
    // The header row's "label" is usually just its first period cell —
    // noise as a row label or section title; drop it when it is one of the
    // header texts.
    if let Some(hr) = header_row {
        if let Some(entry) = rows.iter_mut().find(|e| e.row == hr) {
            if let Some(l) = &entry.label {
                if headers.iter().any(|(_, t)| t == l) {
                    entry.label = None;
                }
            }
        }
    }

    // Sections: contiguous row runs split on ≥2 blank rows, or on a
    // header-styled row (bold label with no data cells) which titles the
    // section it starts.
    let mut sections: Vec<SheetSection> = Vec::new();
    let mut have_current = false;
    for entry in &rows {
        let is_header_styled = entry.bold && entry.formulas == 0 && entry.values == 1 && entry.label.is_some();
        let gap = if have_current { entry.row - sections.last().map(|s| s.end_row).unwrap_or(entry.row) } else { 0 };
        if !have_current || gap > 2 || is_header_styled {
            if sections.len() >= MAX_SECTIONS {
                break;
            }
            sections.push(SheetSection { start_row: entry.row, end_row: entry.row, title: entry.label.clone() });
            have_current = true;
        } else if let Some(cur) = sections.last_mut() {
            cur.end_row = entry.row;
            if cur.title.is_none() && entry.label.is_some() {
                cur.title = entry.label.clone();
            }
        }
    }

    Ok(SheetIndex {
        name: ws.name.clone(),
        hidden: !matches!(ws.state, SheetState::Visible),
        used_range: if max_used_row >= 1 { Some(format!("A1:{}{}", col_letters(max_used_col - 1), max_used_row)) } else { None },
        header_row,
        headers,
        rows,
        sections,
        formula_runs,
        formula_runs_truncated,
    })
}

pub fn defined_names(m: &Model) -> Vec<DefinedName> {
    m.get_defined_name_list()
        .into_iter()
        .filter(|(name, _, formula)| !name.is_empty() && !formula.is_empty())
        .map(|(name, _, formula)| DefinedName { name, reference: formula })
        .collect()
}

/// Whole-workbook index, uncached.
pub fn build_workbook_index(m: &Model) -> Result<WorkbookIndex, String> {
    let mut sheets = Vec::with_capacity(m.workbook.worksheets.len());
    for i in 0..m.workbook.worksheets.len() {
        sheets.push(build_sheet_index(m, i as u32)?);
    }
    Ok(WorkbookIndex { sheets, defined_names: defined_names(m) })
}

/// Per-sheet memo of serialised sheet indexes. The owner marks sheets dirty
/// as it mutates the model (`mark_dirty` / `mark_all_dirty`); a build then
/// re-serialises only those. Keyed by sheet position AND name so a sheet
/// op that reorders or renames cannot serve another sheet's entry.
#[derive(Default)]
pub struct IndexCache {
    sheets: Vec<Option<(String, String)>>, // (name, json)
}

impl IndexCache {
    pub fn mark_dirty(&mut self, sheet: u32) {
        if let Some(slot) = self.sheets.get_mut(sheet as usize) {
            *slot = None;
        }
    }
    pub fn mark_all_dirty(&mut self) {
        self.sheets.clear();
    }
    /// Index JSON for the whole workbook, rebuilding only dirty sheets.
    /// Returns the JSON plus how many sheets were rebuilt.
    pub fn workbook_index_json(&mut self, m: &Model) -> Result<(String, usize), String> {
        let n = m.workbook.worksheets.len();
        self.sheets.resize(n, None);
        let mut rebuilt = 0usize;
        let mut out = String::from("{\"sheets\":[");
        for (i, ws) in m.workbook.worksheets.iter().enumerate() {
            let fresh = match &self.sheets[i] {
                Some((name, _)) if name == &ws.name => false,
                _ => true,
            };
            if fresh {
                let idx = build_sheet_index(m, i as u32)?;
                let json = serde_json::to_string(&idx).map_err(|e| e.to_string())?;
                self.sheets[i] = Some((ws.name.clone(), json));
                rebuilt += 1;
            }
            if i > 0 {
                out.push(',');
            }
            out.push_str(&self.sheets[i].as_ref().expect("filled above").1);
        }
        out.push_str("],\"definedNames\":");
        out.push_str(&serde_json::to_string(&defined_names(m)).map_err(|e| e.to_string())?);
        out.push('}');
        Ok((out, rebuilt))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model() -> Model<'static> {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        let set = |m: &mut Model, r: i32, c: i32, v: &str| m.set_user_input(0, r, c, v.to_string()).unwrap();
        set(&mut m, 1, 1, "Model title");
        // Header row: label column + periods.
        set(&mut m, 4, 1, "USD mm");
        set(&mut m, 4, 2, "FY2024A");
        set(&mut m, 4, 3, "FY2025E");
        set(&mut m, 4, 4, "FY2026E");
        // Section title (bold label, no data) then rows.
        set(&mut m, 6, 1, "Revenue build");
        let mut bold = m.get_style_for_cell(0, 6, 1).unwrap();
        bold.font.b = true;
        m.set_cell_style(0, 6, 1, &bold).unwrap();
        set(&mut m, 7, 1, "Units");
        set(&mut m, 7, 2, "100");
        set(&mut m, 7, 3, "=B7*1.1");
        set(&mut m, 7, 4, "=C7*1.1");
        set(&mut m, 8, 1, "Price");
        set(&mut m, 8, 2, "2");
        set(&mut m, 8, 3, "=B8");
        set(&mut m, 8, 4, "=C8");
        set(&mut m, 9, 1, "Revenue");
        set(&mut m, 9, 2, "=B7*B8");
        set(&mut m, 9, 3, "=C7*C8");
        set(&mut m, 9, 4, "=D7*D8+1"); // breaks the run
        // Gap of 3 blank rows → new section.
        set(&mut m, 13, 1, "Costs");
        set(&mut m, 13, 2, "5");
        m.evaluate();
        m
    }

    #[test]
    fn letters_roundtrip() {
        assert_eq!(col_letters(0), "A");
        assert_eq!(col_letters(25), "Z");
        assert_eq!(col_letters(26), "AA");
        assert_eq!(col_letters(753), "ABZ");
        assert_eq!(letters_to_col("A"), Some(0));
        assert_eq!(letters_to_col("aa"), Some(26));
        assert_eq!(letters_to_col("DB"), Some(105));
        assert_eq!(letters_to_col("1A"), None);
        assert_eq!(letters_to_col("XFE"), None);
    }

    #[test]
    fn helpers_match_the_typescript_rules() {
        assert!(is_label_text("Total revenue"));
        assert!(!is_label_text("Y"));
        assert!(!is_label_text("2024A"));
        assert!(!is_label_text("=A1"));
        assert!(!is_label_text("(1,837)"));
        assert!(is_date_format("yyyy-mm-dd"));
        assert!(is_date_format("[$-409]mmm-yy"));
        assert!(!is_date_format("#,##0.00"));
        assert!(!is_date_format("0.0%"));
        assert!(!is_date_format("\"Day\" 0")); // quoted literal only
        // Same convention as the TypeScript helper (epoch 1899-12-30, no
        // phantom-day correction below serial 61).
        assert_eq!(serial_to_iso_date(1.0).as_deref(), Some("1899-12-31"));
        assert_eq!(serial_to_iso_date(61.0).as_deref(), Some("1900-03-01"));
        assert_eq!(serial_to_iso_date(25569.0).as_deref(), Some("1970-01-01"));
        assert_eq!(serial_to_iso_date(45292.0).as_deref(), Some("2024-01-01"));
        assert_eq!(serial_to_iso_date(0.0), None);
    }

    #[test]
    fn index_structure() {
        let m = model();
        let idx = build_sheet_index(&m, 0).unwrap();
        assert_eq!(idx.name, "Sheet1");
        assert!(!idx.hidden);
        assert_eq!(idx.used_range.as_deref(), Some("A1:D13"));
        assert_eq!(idx.header_row, Some(4));
        assert_eq!(idx.headers, vec![("B".into(), "FY2024A".into()), ("C".into(), "FY2025E".into()), ("D".into(), "FY2026E".into())]);
        let rows: Vec<(i32, Option<&str>, u32, u32, bool, i32)> =
            idx.rows.iter().map(|r| (r.row, r.label.as_deref(), r.values, r.formulas, r.bold, r.sample_col)).collect();
        assert_eq!(rows[0], (1, Some("Model title"), 1, 0, false, 0));
        // Header row keeps its label-column caption (not one of the headers).
        assert_eq!(rows[1], (4, Some("USD mm"), 4, 0, false, 3));
        assert_eq!(rows[2], (6, Some("Revenue build"), 1, 0, true, 0));
        assert_eq!(rows[3], (7, Some("Units"), 2, 2, false, 3));
        assert_eq!(rows[5], (9, Some("Revenue"), 1, 3, false, 3));
        // Sections: title row 1 alone (gap of 3), header row 4 alone, bold
        // "Revenue build" starts a section through row 9, gap → "Costs".
        let secs: Vec<(i32, i32, Option<&str>)> = idx.sections.iter().map(|s| (s.start_row, s.end_row, s.title.as_deref())).collect();
        assert_eq!(secs, vec![(1, 1, Some("Model title")), (4, 4, Some("USD mm")), (6, 9, Some("Revenue build")), (13, 13, Some("Costs"))]);
        // Formula runs: C7:D7 one run, C8:D8 one run, B9:C9 one run, D9 alone.
        let runs: Vec<(i32, i32, i32, &str)> = idx.formula_runs.iter().map(|f| (f.row, f.start_col, f.end_col, f.formula.as_str())).collect();
        assert_eq!(runs, vec![(7, 2, 3, "=B7*1.1"), (8, 2, 3, "=B8"), (9, 1, 2, "=B7*B8"), (9, 3, 3, "=D7*D8+1")]);
        assert!(!idx.formula_runs_truncated);
    }

    #[test]
    fn cache_rebuilds_only_dirty_sheets_and_tracks_renames() {
        let mut m = model();
        m.add_sheet("Other").unwrap();
        m.set_user_input(1, 1, 1, "x".to_string()).unwrap();
        m.evaluate();
        let mut cache = IndexCache::default();
        let (json, rebuilt) = cache.workbook_index_json(&m).unwrap();
        assert_eq!(rebuilt, 2);
        let v: serde_json::Value = serde_json::from_str(&json).unwrap();
        assert_eq!(v["sheets"].as_array().unwrap().len(), 2);
        assert_eq!(v["sheets"][0]["headerRow"], 4);
        assert_eq!(v["sheets"][0]["headers"]["B"], "FY2024A");
        assert_eq!(v["sheets"][0]["rows"][0]["sampleCol"], 0);
        assert_eq!(v["sheets"][1]["usedRange"], "A1:A1");
        assert_eq!(v["definedNames"].as_array().unwrap().len(), 0);

        let (_, rebuilt) = cache.workbook_index_json(&m).unwrap();
        assert_eq!(rebuilt, 0);
        cache.mark_dirty(1);
        let (_, rebuilt) = cache.workbook_index_json(&m).unwrap();
        assert_eq!(rebuilt, 1);
        m.rename_sheet("Other", "Renamed").unwrap();
        let (json, rebuilt) = cache.workbook_index_json(&m).unwrap();
        assert_eq!(rebuilt, 1);
        assert!(json.contains("\"name\":\"Renamed\""));
        cache.mark_all_dirty();
        let (_, rebuilt) = cache.workbook_index_json(&m).unwrap();
        assert_eq!(rebuilt, 2);
    }
}
