//! A1 reference utilities and formula reference shifting.
//!
//! Shifting exists for one purpose in phases 0–2: materializing shared
//! formulas. A shared group stores the formula text once on its master cell;
//! every other member derives its formula by shifting relative references by
//! (Δrow, Δcol). Absolute parts ($) never move.

use super::PatchError;

pub const MAX_ROW: u32 = 1_048_576; // 1-based
pub const MAX_COL: u32 = 16_384; // 1-based (XFD)

/// "BC" -> 0-based column index.
pub fn col_letters_to_index(letters: &str) -> Option<u32> {
    if letters.is_empty() || letters.len() > 3 {
        return None;
    }
    let mut n: u32 = 0;
    for b in letters.bytes() {
        let d = match b {
            b'A'..=b'Z' => (b - b'A' + 1) as u32,
            b'a'..=b'z' => (b - b'a' + 1) as u32,
            _ => return None,
        };
        n = n * 26 + d;
    }
    if n > MAX_COL {
        return None;
    }
    Some(n - 1)
}

/// 0-based column index -> "BC".
pub fn col_index_to_letters(mut c: u32) -> String {
    let mut out = [0u8; 3];
    let mut i = 3;
    loop {
        i -= 1;
        out[i] = b'A' + (c % 26) as u8;
        if c < 26 {
            break;
        }
        c = c / 26 - 1;
    }
    String::from_utf8_lossy(&out[i..]).into_owned()
}

/// "C7" -> 0-based (row, col). Rejects out-of-range and malformed refs.
pub fn parse_cell_ref(a1: &str) -> Option<(u32, u32)> {
    let split = a1.find(|ch: char| ch.is_ascii_digit())?;
    let (letters, digits) = a1.split_at(split);
    let col = col_letters_to_index(letters)?;
    let row: u32 = digits.parse().ok()?;
    if row == 0 || row > MAX_ROW {
        return None;
    }
    Some((row - 1, col))
}

/// 0-based (row, col) -> "C7".
pub fn format_cell_ref(r: u32, c: u32) -> String {
    format!("{}{}", col_index_to_letters(c), r + 1)
}

/// "B2:D5" (or single "B2") -> 0-based inclusive (r1, c1, r2, c2).
pub fn parse_a1_range(range: &str) -> Option<(u32, u32, u32, u32)> {
    let mut parts = range.split(':');
    let start = parse_cell_ref(parts.next()?.trim())?;
    let end = match parts.next() {
        Some(p) => parse_cell_ref(p.trim())?,
        None => start,
    };
    if parts.next().is_some() {
        return None;
    }
    Some((
        start.0.min(end.0),
        start.1.min(end.1),
        start.0.max(end.0),
        start.1.max(end.1),
    ))
}

/// Shift every relative cell reference in `formula` by (dr, dc), leaving
/// absolute parts, quoted strings, and quoted sheet names alone. This is the
/// same translation Excel applies to shared-formula members.
///
/// The tokenizer is deliberately conservative: a candidate is 1–3 letters
/// followed by digits, not embedded in a longer identifier, and not a
/// function call (`LOG10(`). Column-only (`A:B`) and row-only (`3:9`) range
/// refs are also handled since shared formulas may contain them.
pub fn shift_formula(formula: &str, dr: i64, dc: i64) -> Result<String, PatchError> {
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0usize;

    while i < bytes.len() {
        let b = bytes[i];

        // Double-quoted string literal: copy verbatim ("" is an escaped quote).
        if b == b'"' {
            let start = i;
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'"' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            out.push_str(&formula[start..i]);
            continue;
        }

        // Quoted sheet name: copy verbatim ('' is an escaped quote).
        if b == b'\'' {
            let start = i;
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'\'' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            out.push_str(&formula[start..i]);
            continue;
        }

        // A reference candidate can start with $, letters, or digits (row-only
        // ranges like 3:9). Anything glued to the tail of an identifier is not
        // a reference (e.g. the "A1" inside "FOO_A1").
        let prev_is_ident = i > 0 && {
            let p = bytes[i - 1];
            p.is_ascii_alphanumeric() || p == b'_' || p == b'.'
        };

        if !prev_is_ident && (b == b'$' || b.is_ascii_alphabetic() || b.is_ascii_digit()) {
            if let Some((consumed, replacement)) = try_shift_ref(&formula[i..], dr, dc)? {
                out.push_str(&replacement);
                i += consumed;
                continue;
            }
            // Not a reference: consume the whole identifier/number so its
            // interior can't be re-scanned as a reference.
            let start = i;
            while i < bytes.len() {
                let ch = bytes[i];
                if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b'.' || ch == b'$' {
                    i += 1;
                } else {
                    break;
                }
            }
            if i == start {
                i += 1; // lone '$' not followed by a ref
            }
            out.push_str(&formula[start..i]);
            continue;
        }

        out.push(b as char);
        i += 1;
    }

    Ok(out)
}

/// Try to read one reference at the head of `s`. Returns how many bytes were
/// consumed and the shifted text, or None when `s` doesn't start with a ref.
fn try_shift_ref(s: &str, dr: i64, dc: i64) -> Result<Option<(usize, String)>, PatchError> {
    let bytes = s.as_bytes();

    // --- row-only range: 3:9, $3:$9 ---
    if let Some((consumed, text)) = try_row_only_range(s, dr)? {
        return Ok(Some((consumed, text)));
    }

    let mut i = 0usize;
    let col_abs = bytes.first() == Some(&b'$');
    if col_abs {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() && i - letters_start < 3 {
        i += 1;
    }
    let letters = &s[letters_start..i];
    if letters.is_empty() {
        return Ok(None);
    }

    // --- column-only range: A:B, $A:$B ---
    if i < bytes.len() && bytes[i] == b':' {
        if let Some((consumed, text)) = try_col_only_tail(s, i, letters, col_abs, dc)? {
            return Ok(Some((consumed, text)));
        }
    }

    let row_abs = i < bytes.len() && bytes[i] == b'$';
    let digits_start = if row_abs { i + 1 } else { i };
    let mut j = digits_start;
    while j < bytes.len() && bytes[j].is_ascii_digit() {
        j += 1;
    }
    if j == digits_start {
        return Ok(None); // letters with no row → identifier, not a ref
    }
    // A ref can't be immediately followed by more identifier chars or `(`
    // (that would make it a name or a function like LOG10().
    if let Some(&next) = bytes.get(j) {
        if next.is_ascii_alphanumeric() || next == b'_' || next == b'.' || next == b'(' {
            return Ok(None);
        }
    }

    let col = match col_letters_to_index(letters) {
        Some(c) => c,
        None => return Ok(None),
    };
    let row1: u64 = s[digits_start..j].parse().map_err(|_| PatchError::BadRef)?;
    if row1 == 0 || row1 > MAX_ROW as u64 {
        return Ok(None);
    }

    let new_col = if col_abs {
        col as i64
    } else {
        col as i64 + dc
    };
    let new_row1 = if row_abs {
        row1 as i64
    } else {
        row1 as i64 + dr
    };
    if new_col < 0 || new_col >= MAX_COL as i64 || new_row1 < 1 || new_row1 > MAX_ROW as i64 {
        return Err(PatchError::RefOutOfRange);
    }

    let mut text = String::new();
    if col_abs {
        text.push('$');
    }
    text.push_str(&col_index_to_letters(new_col as u32));
    if row_abs {
        text.push('$');
    }
    text.push_str(&new_row1.to_string());
    Ok(Some((j, text)))
}

fn try_row_only_range(s: &str, dr: i64) -> Result<Option<(usize, String)>, PatchError> {
    let bytes = s.as_bytes();
    let mut i = 0usize;
    let abs1 = bytes.first() == Some(&b'$');
    if abs1 {
        i += 1;
    }
    let d1 = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == d1 || bytes.get(i) != Some(&b':') {
        return Ok(None);
    }
    let first_end = i;
    i += 1;
    let abs2 = bytes.get(i) == Some(&b'$');
    if abs2 {
        i += 1;
    }
    let d2 = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == d2 {
        return Ok(None);
    }
    // Reject things like "3:9x" — a ref must end at a boundary.
    if let Some(&next) = bytes.get(i) {
        if next.is_ascii_alphanumeric() || next == b'_' || next == b'.' || next == b'(' {
            return Ok(None);
        }
    }
    let r1: i64 = s[d1..first_end].parse().map_err(|_| PatchError::BadRef)?;
    let r2: i64 = s[d2..i].parse().map_err(|_| PatchError::BadRef)?;
    let n1 = if abs1 { r1 } else { r1 + dr };
    let n2 = if abs2 { r2 } else { r2 + dr };
    if n1 < 1 || n2 < 1 || n1 > MAX_ROW as i64 || n2 > MAX_ROW as i64 {
        return Err(PatchError::RefOutOfRange);
    }
    let mut text = String::new();
    if abs1 {
        text.push('$');
    }
    text.push_str(&n1.to_string());
    text.push(':');
    if abs2 {
        text.push('$');
    }
    text.push_str(&n2.to_string());
    Ok(Some((i, text)))
}

/// `s[..colon]` holds the first column letters (abs1 already consumed); try
/// to read `:$?LETTERS` with no digits after — a column-only range.
fn try_col_only_tail(
    s: &str,
    colon: usize,
    letters1: &str,
    abs1: bool,
    dc: i64,
) -> Result<Option<(usize, String)>, PatchError> {
    let bytes = s.as_bytes();
    let mut i = colon + 1;
    let abs2 = bytes.get(i) == Some(&b'$');
    if abs2 {
        i += 1;
    }
    let l2 = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() && i - l2 < 3 {
        i += 1;
    }
    if i == l2 {
        return Ok(None);
    }
    // If digits follow, this is a normal range like A1:B2, not column-only.
    if bytes.get(i).map_or(false, |b| b.is_ascii_digit() || *b == b'$') {
        return Ok(None);
    }
    if let Some(&next) = bytes.get(i) {
        if next.is_ascii_alphanumeric() || next == b'_' || next == b'.' || next == b'(' {
            return Ok(None);
        }
    }
    let c1 = match col_letters_to_index(letters1) {
        Some(c) => c,
        None => return Ok(None),
    };
    let c2 = match col_letters_to_index(&s[l2..i]) {
        Some(c) => c,
        None => return Ok(None),
    };
    let n1 = if abs1 { c1 as i64 } else { c1 as i64 + dc };
    let n2 = if abs2 { c2 as i64 } else { c2 as i64 + dc };
    if n1 < 0 || n2 < 0 || n1 >= MAX_COL as i64 || n2 >= MAX_COL as i64 {
        return Err(PatchError::RefOutOfRange);
    }
    let mut text = String::new();
    if abs1 {
        text.push('$');
    }
    text.push_str(&col_index_to_letters(n1 as u32));
    text.push(':');
    if abs2 {
        text.push('$');
    }
    text.push_str(&col_index_to_letters(n2 as u32));
    Ok(Some((i, text)))
}

// ---------------------------------------------------------------------------
// positional row/column shifting (insert/delete)
// ---------------------------------------------------------------------------

/// One row/column insert or delete, in the coordinate space of the sheet
/// BEFORE the operation. Indices are 0-based.
#[derive(Debug, Clone, Copy)]
pub struct RowColShift {
    /// true = the row axis moves, false = the column axis.
    pub rows: bool,
    pub start: u32,
    pub count: u32,
    pub insert: bool,
}

impl RowColShift {
    /// Adjust a single 0-based index. `None` = the index was deleted.
    pub fn adjust_index(&self, i: u32) -> Option<u32> {
        if self.insert {
            if i >= self.start {
                Some(i + self.count)
            } else {
                Some(i)
            }
        } else if i < self.start {
            Some(i)
        } else if i < self.start + self.count {
            None
        } else {
            Some(i - self.count)
        }
    }

    /// Adjust an index that must survive (anchors, view corners): deleted
    /// positions collapse onto the deletion point instead of vanishing.
    pub fn clamp_index(&self, i: u32) -> u32 {
        self.adjust_index(i).unwrap_or(self.start)
    }

    /// Adjust an inclusive 0-based span on this axis. `None` = the whole
    /// span was deleted. Partially-deleted spans shrink (Excel semantics).
    pub fn adjust_span(&self, a: u32, b: u32) -> Option<(u32, u32)> {
        if self.insert {
            let na = if a >= self.start { a + self.count } else { a };
            let nb = if b >= self.start { b + self.count } else { b };
            return Some((na, nb));
        }
        let end = self.start + self.count; // exclusive
        let na = if a < self.start {
            a
        } else if a >= end {
            a - self.count
        } else {
            self.start
        };
        let nb = if b < self.start {
            b
        } else if b >= end {
            b - self.count
        } else {
            // Last surviving position before the deleted block.
            match self.start.checked_sub(1) {
                Some(p) => p,
                None => return None, // block starts at 0 and b is inside it
            }
        };
        if a >= self.start && b < end {
            return None; // fully inside the deleted block
        }
        if nb < na {
            return None;
        }
        Some((na, nb))
    }

    fn max_index(&self) -> u32 {
        if self.rows {
            MAX_ROW - 1
        } else {
            MAX_COL - 1
        }
    }
}

/// Quote a sheet name for use in a formula if it needs it.
pub fn format_sheet_qualifier(name: &str) -> String {
    let plain = !name.is_empty()
        && !name.as_bytes()[0].is_ascii_digit()
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || c == '_' || c == '.')
        // A name that parses as a cell ref (e.g. "A1") must be quoted.
        && parse_cell_ref(name).is_none();
    if plain {
        name.to_string()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

fn sheet_names_eq(a: &str, b: &str) -> bool {
    // Excel sheet names are case-insensitive.
    a.len() == b.len()
        && a.chars()
            .zip(b.chars())
            .all(|(x, y)| x.eq_ignore_ascii_case(&y))
}

/// One reference token: `$?COL$?ROW` optionally `:$?COL$?ROW`, or a
/// row-only / column-only range. Coordinates are 0-based; `None` on an axis
/// means "unbounded" (whole-row/whole-column ranges).
struct RefToken {
    consumed: usize,
    // (col, row) with absolute flags; row/col None for partial refs.
    c1: Option<(u32, bool)>,
    r1: Option<(u32, bool)>,
    c2: Option<(u32, bool)>,
    r2: Option<(u32, bool)>,
    is_range: bool,
}

fn parse_ref_endpoint(s: &str) -> Option<(usize, Option<(u32, bool)>, Option<(u32, bool)>)> {
    let bytes = s.as_bytes();
    let mut i = 0usize;
    let col_abs = bytes.first() == Some(&b'$');
    if col_abs {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() && i - letters_start < 3 {
        i += 1;
    }
    let letters = &s[letters_start..i];
    let row_abs = bytes.get(i) == Some(&b'$');
    let digits_start = if row_abs { i + 1 } else { i };
    let mut j = digits_start;
    while j < bytes.len() && bytes[j].is_ascii_digit() {
        j += 1;
    }
    let digits = &s[digits_start..j];

    let col = if letters.is_empty() {
        None
    } else {
        Some((col_letters_to_index(letters)?, col_abs))
    };
    let row = if digits.is_empty() {
        if row_abs {
            return None; // "$A$" with no digits — not a ref
        }
        None
    } else {
        let r: u64 = digits.parse().ok()?;
        if r == 0 || r > MAX_ROW as u64 {
            return None;
        }
        Some(((r - 1) as u32, row_abs))
    };
    if col.is_none() && row.is_none() {
        return None;
    }
    // Column-only endpoint must not consume the '$' meant for a row.
    let consumed = if digits.is_empty() { i } else { j };
    Some((consumed, col, row))
}

fn ends_at_boundary(s: &str, i: usize) -> bool {
    match s.as_bytes().get(i) {
        None => true,
        Some(&b) => !(b.is_ascii_alphanumeric() || b == b'_' || b == b'.' || b == b'('),
    }
}

/// Parse one full reference token at the head of `s`.
fn parse_ref_token(s: &str) -> Option<RefToken> {
    let (n1, c1, r1) = parse_ref_endpoint(s)?;
    // Single endpoint must be a full cell ref unless a range follows
    // (A:B / 3:9 / A1:B2). "A" alone is an identifier, "3" alone a number.
    let has_colon = s.as_bytes().get(n1) == Some(&b':');
    if !has_colon {
        if c1.is_some() && r1.is_some() && ends_at_boundary(s, n1) {
            return Some(RefToken {
                consumed: n1,
                c1,
                r1,
                c2: None,
                r2: None,
                is_range: false,
            });
        }
        return None;
    }
    let rest = &s[n1 + 1..];
    let (n2, c2, r2) = match parse_ref_endpoint(rest) {
        Some(x) => x,
        None => {
            // "A1:" with junk after the colon — treat head as single cell.
            if c1.is_some() && r1.is_some() {
                return Some(RefToken {
                    consumed: n1,
                    c1,
                    r1,
                    c2: None,
                    r2: None,
                    is_range: false,
                });
            }
            return None;
        }
    };
    let total = n1 + 1 + n2;
    if !ends_at_boundary(s, total) {
        return None;
    }
    // The two endpoints must be the same shape: cell:cell, col:col, row:row.
    let shape_ok = (c1.is_some() == c2.is_some()) && (r1.is_some() == r2.is_some());
    if !shape_ok {
        // A1:B — take the head as a single cell ref if it is one.
        if c1.is_some() && r1.is_some() && ends_at_boundary(s, n1) {
            return Some(RefToken {
                consumed: n1,
                c1,
                r1,
                c2: None,
                r2: None,
                is_range: false,
            });
        }
        return None;
    }
    Some(RefToken {
        consumed: total,
        c1,
        r1,
        c2,
        r2,
        is_range: true,
    })
}

fn render_endpoint(col: Option<(u32, bool)>, row: Option<(u32, bool)>, out: &mut String) {
    if let Some((c, abs)) = col {
        if abs {
            out.push('$');
        }
        out.push_str(&col_index_to_letters(c));
    }
    if let Some((r, abs)) = row {
        if abs {
            out.push('$');
        }
        out.push_str(&(r + 1).to_string());
    }
}

/// Positionally adjust one parsed ref token. Returns the replacement text
/// ("#REF!" when the referenced cells were deleted).
fn adjust_ref_token(tok: &RefToken, shift: &RowColShift) -> Result<String, PatchError> {
    // Pick the coordinates on the shifted axis; the other axis passes through.
    let axis1 = if shift.rows { tok.r1 } else { tok.c1 };
    let axis2 = if shift.rows { tok.r2 } else { tok.c2 };

    let rebuild = |a1: Option<(u32, bool)>,
                   a2: Option<(u32, bool)>|
     -> Result<String, PatchError> {
        let (c1, r1, c2, r2) = if shift.rows {
            (tok.c1, a1, tok.c2, a2)
        } else {
            (a1, tok.r1, a2, tok.r2)
        };
        let mut out = String::new();
        render_endpoint(c1, r1, &mut out);
        if tok.is_range {
            out.push(':');
            render_endpoint(c2, r2, &mut out);
        }
        Ok(out)
    };

    // Axis not present in the token (e.g. column-only range under a row
    // shift): the ref is unbounded on the shifted axis — nothing to do.
    let Some((i1, abs1)) = axis1 else {
        return rebuild(axis1, axis2);
    };

    if !tok.is_range {
        return match shift.adjust_index(i1) {
            Some(n) if n <= shift.max_index() => rebuild(Some((n, abs1)), None),
            Some(_) => Err(PatchError::RefOutOfRange),
            None => Ok("#REF!".to_string()),
        };
    }

    let Some((i2, abs2)) = axis2 else {
        return rebuild(axis1, axis2); // malformed; leave as-is
    };
    let (lo, hi) = (i1.min(i2), i1.max(i2));
    match shift.adjust_span(lo, hi) {
        Some((na, nb)) if nb <= shift.max_index() => {
            // Preserve the original endpoint order.
            let (n1, n2) = if i1 <= i2 { (na, nb) } else { (nb, na) };
            rebuild(Some((n1, abs1)), Some((n2, abs2)))
        }
        Some(_) => Err(PatchError::RefOutOfRange),
        None => Ok("#REF!".to_string()),
    }
}

enum QualifierAction<'a> {
    /// Row/col shift: adjust refs on the target sheet.
    Shift {
        shift: &'a RowColShift,
        target_sheet: &'a str,
        on_target_sheet: bool,
    },
    /// Sheet rename: rewrite qualifiers, leave refs alone.
    Rename { old: &'a str, new: &'a str },
}

/// Positionally adjust every reference in `formula` affected by `shift`.
/// Only refs qualified with `target_sheet` — or unqualified refs when the
/// formula itself lives on the target sheet — move. Deleted refs become
/// `#REF!`.
pub fn adjust_formula(
    formula: &str,
    shift: &RowColShift,
    target_sheet: &str,
    on_target_sheet: bool,
) -> Result<String, PatchError> {
    walk_formula(
        formula,
        &QualifierAction::Shift {
            shift,
            target_sheet,
            on_target_sheet,
        },
    )
}

/// Rewrite every `old!` / `'old'!` sheet qualifier in `formula` to `new`
/// (quoted as needed). References themselves are untouched.
pub fn rename_sheet_in_formula(formula: &str, old: &str, new: &str) -> Result<String, PatchError> {
    walk_formula(formula, &QualifierAction::Rename { old, new })
}

fn walk_formula(formula: &str, action: &QualifierAction) -> Result<String, PatchError> {
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len() + 8);
    let mut i = 0usize;
    // Sheet qualifier of the immediately-following ref token, if any.
    let mut qualifier: Option<String> = None;
    // Set when the previous token was an identifier/quoted name followed by
    // ':' — the "Sheet1:" head of a 3-D reference.
    let mut three_d_head: Option<String> = None;

    while i < bytes.len() {
        let b = bytes[i];

        // Double-quoted string literal: copy verbatim.
        if b == b'"' {
            let start = i;
            i += 1;
            while i < bytes.len() {
                if bytes[i] == b'"' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            out.push_str(&formula[start..i]);
            qualifier = None;
            three_d_head = None;
            continue;
        }

        // Quoted sheet name.
        if b == b'\'' {
            let start = i;
            i += 1;
            let name_start = i;
            while i < bytes.len() {
                if bytes[i] == b'\'' {
                    if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                        i += 2;
                        continue;
                    }
                    break;
                }
                i += 1;
            }
            let raw_name = &formula[name_start..i];
            i += 1; // closing quote
            let name = raw_name.replace("''", "'");
            let followed_by_bang = bytes.get(i) == Some(&b'!');
            let followed_by_colon = bytes.get(i) == Some(&b':');
            handle_qualifier(
                &formula[start..i],
                &name,
                followed_by_bang,
                followed_by_colon,
                action,
                &mut out,
                &mut qualifier,
                &mut three_d_head,
            )?;
            continue;
        }

        let prev_is_ident = i > 0 && {
            let p = bytes[i - 1];
            p.is_ascii_alphanumeric() || p == b'_' || p == b'.'
        };

        if !prev_is_ident && (b == b'$' || b.is_ascii_alphabetic() || b.is_ascii_digit()) {
            // Try a reference token first.
            if let Some(tok) = parse_ref_token(&formula[i..]) {
                // "A1:" can also be the head of "A1:B2" — parse_ref_token
                // already consumed the full range in that case. But a token
                // followed by '!' is a sheet qualifier, not a ref (rare:
                // sheets named like cell refs are quoted, so this is safe).
                if bytes.get(i + tok.consumed) != Some(&b'!') {
                    let replacement = match action {
                        QualifierAction::Shift {
                            shift,
                            target_sheet,
                            on_target_sheet,
                        } => {
                            let applies = match (&qualifier, &three_d_head) {
                                (_, Some(_)) => {
                                    // 3-D span: refuse if it could involve
                                    // the target sheet, else leave alone.
                                    let involved = qualifier
                                        .as_deref()
                                        .is_some_and(|q| sheet_names_eq(q, target_sheet))
                                        || three_d_head
                                            .as_deref()
                                            .is_some_and(|h| sheet_names_eq(h, target_sheet));
                                    if involved {
                                        return Err(PatchError::Unsupported(
                                            "3-D reference spans the edited sheet".into(),
                                        ));
                                    }
                                    false
                                }
                                (Some(q), None) => sheet_names_eq(q, target_sheet),
                                (None, None) => *on_target_sheet,
                            };
                            if applies {
                                adjust_ref_token(&tok, shift)?
                            } else {
                                formula[i..i + tok.consumed].to_string()
                            }
                        }
                        QualifierAction::Rename { .. } => {
                            formula[i..i + tok.consumed].to_string()
                        }
                    };
                    out.push_str(&replacement);
                    i += tok.consumed;
                    qualifier = None;
                    three_d_head = None;
                    continue;
                }
            }

            // Identifier / number: consume as a unit.
            let start = i;
            while i < bytes.len() {
                let ch = bytes[i];
                if ch.is_ascii_alphanumeric() || ch == b'_' || ch == b'.' || ch == b'$' {
                    i += 1;
                } else {
                    break;
                }
            }
            if i == start {
                i += 1; // lone '$'
            }
            let ident = &formula[start..i];
            let followed_by_bang = bytes.get(i) == Some(&b'!');
            let followed_by_colon = bytes.get(i) == Some(&b':');
            handle_qualifier(
                ident,
                ident,
                followed_by_bang,
                followed_by_colon,
                action,
                &mut out,
                &mut qualifier,
                &mut three_d_head,
            )?;
            continue;
        }

        if b != b'!' {
            // Any other operator/punctuation clears qualifier context,
            // except ':' which may glue a 3-D head to its tail.
            if b != b':' {
                qualifier = None;
                three_d_head = None;
            }
        }
        out.push(b as char);
        i += 1;
    }

    Ok(out)
}

/// Shared tail for quoted and unquoted potential sheet qualifiers. The
/// caller has already advanced past the name text; this only decides how to
/// render it and what qualifier context the following token sees.
#[allow(clippy::too_many_arguments)]
fn handle_qualifier(
    original_text: &str,
    name: &str,
    followed_by_bang: bool,
    followed_by_colon: bool,
    action: &QualifierAction,
    out: &mut String,
    qualifier: &mut Option<String>,
    three_d_head: &mut Option<String>,
) -> Result<(), PatchError> {
    if followed_by_bang {
        let rendered = match action {
            QualifierAction::Rename { old, new }
                if three_d_head.is_none() && sheet_names_eq(name, old) =>
            {
                format_sheet_qualifier(new)
            }
            QualifierAction::Rename { old, new } if three_d_head.is_some() => {
                // 3-D tail: rename head/tail independently is fine.
                if sheet_names_eq(name, old) {
                    format_sheet_qualifier(new)
                } else {
                    original_text.to_string()
                }
            }
            _ => original_text.to_string(),
        };
        out.push_str(&rendered);
        *qualifier = Some(name.to_string());
        // Keep three_d_head as-is: "A:B!ref" means head A, tail B.
        return Ok(());
    }
    if followed_by_colon {
        // Might be the head of a 3-D reference (Sheet1:Sheet3!A1).
        let rendered = match action {
            QualifierAction::Rename { old, new } if sheet_names_eq(name, old) => {
                // Only rename if this really is a 3-D head; a bare
                // identifier before ':' in a range (A:B) never reaches
                // here (parse_ref_token owns those). Conservatively
                // rename — worst case we renamed a matching name-like
                // token that Excel treats as a name, which cannot
                // collide with a sheet name anyway.
                format_sheet_qualifier(new)
            }
            _ => original_text.to_string(),
        };
        out.push_str(&rendered);
        *three_d_head = Some(name.to_string());
        *qualifier = None;
        return Ok(());
    }
    out.push_str(original_text);
    *qualifier = None;
    *three_d_head = None;
    Ok(())
}

/// Adjust a space-separated sqref list ("A1:B2 D4"). Deleted ranges are
/// dropped; `Ok(None)` when nothing survives. An insert that would push a
/// range past the sheet edge is an error (Excel refuses those inserts too).
pub fn adjust_sqref_list(
    list: &str,
    shift: &RowColShift,
) -> Result<Option<String>, PatchError> {
    let mut parts: Vec<String> = Vec::new();
    for token in list.split_whitespace() {
        match parse_a1_range(token) {
            Some((r1, c1, r2, c2)) => {
                let (a, b) = if shift.rows { (r1, r2) } else { (c1, c2) };
                match shift.adjust_span(a, b) {
                    Some((na, nb)) if nb <= shift.max_index() => {
                        let (nr1, nc1, nr2, nc2) = if shift.rows {
                            (na, c1, nb, c2)
                        } else {
                            (r1, na, r2, nb)
                        };
                        if nr1 == nr2 && nc1 == nc2 {
                            parts.push(format_cell_ref(nr1, nc1));
                        } else {
                            parts.push(format!(
                                "{}:{}",
                                format_cell_ref(nr1, nc1),
                                format_cell_ref(nr2, nc2)
                            ));
                        }
                    }
                    Some(_) => return Err(PatchError::RefOutOfRange),
                    None => {}
                }
            }
            None => parts.push(token.to_string()),
        }
    }
    Ok(if parts.is_empty() {
        None
    } else {
        Some(parts.join(" "))
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn col_roundtrip() {
        for (s, i) in [("A", 0u32), ("Z", 25), ("AA", 26), ("BC", 54), ("XFD", 16383)] {
            assert_eq!(col_letters_to_index(s), Some(i), "{s}");
            assert_eq!(col_index_to_letters(i), s);
        }
        assert_eq!(col_letters_to_index("XFE"), None);
        assert_eq!(col_letters_to_index("AAAA"), None);
    }

    #[test]
    fn cell_ref_roundtrip() {
        assert_eq!(parse_cell_ref("C7"), Some((6, 2)));
        assert_eq!(parse_cell_ref("XFD1048576"), Some((1_048_575, 16_383)));
        assert_eq!(parse_cell_ref("XFD1048577"), None);
        assert_eq!(format_cell_ref(6, 2), "C7");
    }

    #[test]
    fn shift_basic() {
        assert_eq!(shift_formula("A1+B2", 1, 0).unwrap(), "A2+B3");
        assert_eq!(shift_formula("$A$1+B2", 5, 3).unwrap(), "$A$1+E7");
        assert_eq!(shift_formula("SUM(A1:A10)", 0, 2).unwrap(), "SUM(C1:C10)");
        assert_eq!(shift_formula("$A1+A$1", 1, 1).unwrap(), "$A2+B$1");
    }

    #[test]
    fn shift_skips_strings_and_names() {
        assert_eq!(
            shift_formula("IF(A1=\"B2\",LOG10(A1),MY_A1)", 1, 0).unwrap(),
            "IF(A2=\"B2\",LOG10(A2),MY_A1)"
        );
        // Quoted sheet names don't shift; the ref after ! does.
        assert_eq!(
            shift_formula("'Bal A1 Sheet'!B2+Plain!C3", 1, 1).unwrap(),
            "'Bal A1 Sheet'!C3+Plain!D4"
        );
        // Escaped quote inside sheet name.
        assert_eq!(
            shift_formula("'It''s A1'!A1", 1, 0).unwrap(),
            "'It''s A1'!A2"
        );
    }

    #[test]
    fn shift_whole_row_col_ranges() {
        assert_eq!(shift_formula("SUM(A:B)", 0, 1).unwrap(), "SUM(B:C)");
        assert_eq!(shift_formula("SUM($A:B)", 0, 1).unwrap(), "SUM($A:C)");
        assert_eq!(shift_formula("SUM(3:9)", 2, 0).unwrap(), "SUM(5:11)");
        assert_eq!(shift_formula("SUM($3:9)", 2, 0).unwrap(), "SUM($3:11)");
        // HLOOKUP row-index style: numbers that are NOT refs stay put.
        assert_eq!(
            shift_formula("HLOOKUP(C4,A1:CC26,10,FALSE)", 1, 0).unwrap(),
            "HLOOKUP(C5,A2:CC27,10,FALSE)"
        );
    }

    #[test]
    fn shift_out_of_range_errors() {
        assert!(matches!(
            shift_formula("A1", -1, 0),
            Err(PatchError::RefOutOfRange)
        ));
        assert!(matches!(
            shift_formula("XFD1", 0, 1),
            Err(PatchError::RefOutOfRange)
        ));
    }

    #[test]
    fn range_parse() {
        assert_eq!(parse_a1_range("B2:D5"), Some((1, 1, 4, 3)));
        assert_eq!(parse_a1_range("B2"), Some((1, 1, 1, 1)));
        assert_eq!(parse_a1_range("D5:B2"), Some((1, 1, 4, 3)));
        assert_eq!(parse_a1_range("nope"), None);
    }
}
