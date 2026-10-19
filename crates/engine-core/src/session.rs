//! Formula cells whose value this session changed, so a save can refresh
//! their cached results in the file (`<v>`) and readers that don't calculate
//! (pandas, previews, data pipelines) see current numbers.
//!
//! Only cells the session's edits changed are refreshed: a full comparison
//! with the file's caches would also rewrite cells this engine evaluates
//! differently from Excel (add-in calls, last-digit float differences) and
//! touch sheets nobody edited.
//!
//! Incremental recalculation reports what it changed. A structural edit
//! (rows, columns, sheets) is followed by a full recalculation that reports
//! nothing, so the first structural edit after an evaluation fingerprints
//! every formula's value (8 bytes each, until that recalculation), later
//! edits shift the fingerprints with the cells, and the recalculation
//! compares against them.

use ironcalc_base::expressions::token::Error;
use ironcalc_base::types::{Cell, FormulaValue};
use ironcalc_base::Model;
use serde_json::{json, Value};
use std::collections::hash_map::DefaultHasher;
use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};

/// Keys are (sheet_id, row, column): sheet ids survive renames and reorders.
type Key = (u32, i32, i32);

#[derive(Default)]
pub struct SessionChanges {
    changed: HashSet<Key>,
    before: Option<HashMap<Key, u64>>,
}

fn formula_value(cell: &Cell) -> Option<&FormulaValue> {
    match cell {
        Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => Some(v),
        _ => None,
    }
}

fn fingerprint(v: &FormulaValue) -> u64 {
    let mut h = DefaultHasher::new();
    match v {
        FormulaValue::Unevaluated => 0u8.hash(&mut h),
        FormulaValue::Boolean(b) => (1u8, *b).hash(&mut h),
        FormulaValue::Number(n) => (2u8, n.to_bits()).hash(&mut h),
        FormulaValue::Text(s) => (3u8, s).hash(&mut h),
        FormulaValue::Error { ei, .. } => (4u8, ei.to_string()).hash(&mut h),
    }
    h.finish()
}

fn sheet_id(m: &Model, sheet: u32) -> Option<u32> {
    m.workbook.worksheets.get(sheet as usize).map(|ws| ws.sheet_id)
}

/// The cached-value wire shape of xlsx_patch (`{"t":"n","n":…}` …), or None
/// for what a file can't cache: unevaluated cells and this engine's own
/// errors (#ERROR!, #N/IMPL, #CIRC!), which Excel doesn't know.
fn wire_value(v: &FormulaValue) -> Option<Value> {
    match v {
        FormulaValue::Number(n) if n.is_finite() => Some(json!({"t": "n", "n": n})),
        FormulaValue::Text(s) => Some(json!({"t": "s", "s": s})),
        FormulaValue::Boolean(b) => Some(json!({"t": "b", "b": b})),
        FormulaValue::Error { ei, .. } => match ei {
            Error::ERROR | Error::NIMPL | Error::CIRC => None,
            e => Some(json!({"t": "e", "e": e.to_string()})),
        },
        _ => None,
    }
}

impl SessionChanges {
    /// Cells (sheet index, row, column) an evaluation reported or an edit wrote.
    pub fn record(&mut self, m: &Model, cells: &[(u32, i32, i32)]) {
        for &(sheet, row, column) in cells {
            if let Some(id) = sheet_id(m, sheet) {
                self.changed.insert((id, row, column));
            }
        }
    }

    /// Before a structural edit: fingerprint every formula's value, unless
    /// an earlier edit since the last evaluation already did.
    pub fn before_structural(&mut self, m: &Model) {
        if self.before.is_some() {
            return;
        }
        let mut before = HashMap::new();
        for ws in &m.workbook.worksheets {
            for (row, columns) in &ws.sheet_data {
                for (column, cell) in columns {
                    if let Some(v) = formula_value(cell) {
                        before.insert((ws.sheet_id, *row, *column), fingerprint(v));
                    }
                }
            }
        }
        self.before = Some(before);
    }

    /// After the full recalculation that follows structural edits: every
    /// formula whose value differs from its fingerprint changed.
    pub fn after_full_eval(&mut self, m: &Model) {
        let Some(before) = self.before.take() else {
            return;
        };
        for ws in &m.workbook.worksheets {
            for (row, columns) in &ws.sheet_data {
                for (column, cell) in columns {
                    if let Some(v) = formula_value(cell) {
                        let key = (ws.sheet_id, *row, *column);
                        if before.get(&key) != Some(&fingerprint(v)) {
                            self.changed.insert(key);
                        }
                    }
                }
            }
        }
    }

    /// A recalculation that is not the session's doing (the deferred first
    /// evaluation after opening): drop the fingerprints.
    pub fn discard_before(&mut self) {
        self.before = None;
    }

    /// Rows `at..` of a sheet moved by `count` (negative: `-count` rows
    /// deleted starting at `at`). Applies to the tracked cells and the
    /// fingerprints alike.
    pub fn shift_rows(&mut self, id: u32, at: i32, count: i32) {
        self.shift(id, at, count, true);
    }

    pub fn shift_columns(&mut self, id: u32, at: i32, count: i32) {
        self.shift(id, at, count, false);
    }

    fn shift(&mut self, id: u32, at: i32, count: i32, rows: bool) {
        let moved = |(s, r, c): Key| -> Option<Key> {
            if s != id {
                return Some((s, r, c));
            }
            let p = if rows { r } else { c };
            let p = if p < at {
                p
            } else if count < 0 && p < at - count {
                return None; // deleted
            } else {
                p + count
            };
            Some(if rows { (s, p, c) } else { (s, r, p) })
        };
        self.changed = self.changed.drain().filter_map(moved).collect();
        if let Some(before) = self.before.take() {
            self.before = Some(before.into_iter().filter_map(|(k, v)| moved(k).map(|k| (k, v))).collect());
        }
    }

    pub fn remove_sheet(&mut self, id: u32) {
        self.changed.retain(|k| k.0 != id);
        if let Some(before) = self.before.as_mut() {
            before.retain(|k, _| k.0 != id);
        }
    }

    /// `[{"sheet": name, "r": row0, "c": col0, "value": {…}}]` (0-based) for
    /// every tracked cell that holds a formula now, with its current value.
    pub fn refresh_json(&self, m: &Model) -> String {
        let names: HashMap<u32, (usize, &str)> = m
            .workbook
            .worksheets
            .iter()
            .enumerate()
            .map(|(i, ws)| (ws.sheet_id, (i, ws.name.as_str())))
            .collect();
        let mut keys: Vec<&Key> = self.changed.iter().collect();
        keys.sort_unstable();
        let mut out = Vec::with_capacity(keys.len());
        for &(id, row, column) in keys {
            let Some(&(index, name)) = names.get(&id) else { continue };
            let cell = m.workbook.worksheets[index].sheet_data.get(&row).and_then(|r| r.get(&column));
            if let Some(value) = cell.and_then(formula_value).and_then(wire_value) {
                out.push(json!({"sheet": name, "r": row - 1, "c": column - 1, "value": value}));
            }
        }
        Value::Array(out).to_string()
    }
}
