//! Value diffs over the model, built for memory: the desktop's settle and
//! sheet-op / structural commands compare formula values before and after
//! an operation and send the grid only the cells that moved. Done with
//! `serde_json::Value` maps that cost 184 MB each on a 926k-formula model
//! (and the diff built two). `Val` is a compact value, the "before" map is
//! the only map, and the "after" side streams over the model.

use crate::is_unevaluated;
use ironcalc_base::cell::CellValue;
use ironcalc_base::Model;
use serde_json::Value;
use std::collections::HashMap;

/// A formula result, compact (no `serde_json::Value` overhead).
#[derive(Debug, Clone, PartialEq)]
pub enum Val {
    None,
    Num(f64),
    Bool(bool),
    Text(String),
}

impl Val {
    pub fn from_cell_value(v: CellValue) -> Val {
        match v {
            CellValue::None => Val::None,
            CellValue::Number(n) => Val::Num(n),
            CellValue::Boolean(b) => Val::Bool(b),
            CellValue::String(s) => Val::Text(s),
        }
    }

    pub fn from_json(v: &Value) -> Val {
        match v {
            Value::Null => Val::None,
            Value::Number(n) => n.as_f64().map(Val::Num).unwrap_or(Val::None),
            Value::Bool(b) => Val::Bool(*b),
            Value::String(s) => Val::Text(s.clone()),
            other => Val::Text(other.to_string()),
        }
    }

    pub fn to_json(&self) -> Value {
        match self {
            Val::None => Value::Null,
            Val::Num(n) => serde_json::Number::from_f64(*n).map(Value::Number).unwrap_or_else(|| Value::String(n.to_string())),
            Val::Bool(b) => Value::Bool(*b),
            Val::Text(s) => Value::String(s.clone()),
        }
    }

    /// Empty to the grid: nothing, or the empty string.
    pub fn is_blank(&self) -> bool {
        matches!(self, Val::None) || matches!(self, Val::Text(s) if s.is_empty())
    }
}

/// "Did this value move": numbers within 1e-9 relative (the file's cache
/// and our evaluation differ in the 15th digit on ~0.2% of cells), and
/// blank ≡ blank whether it is null or "".
pub fn equivalent(a: &Val, b: &Val) -> bool {
    match (a, b) {
        (Val::Num(x), Val::Num(y)) => (x - y).abs() <= 1e-9 * x.abs().max(y.abs()).max(1.0),
        _ if a.is_blank() && b.is_blank() => true,
        _ => a == b,
    }
}

pub type ValueMap = HashMap<(u32, i32, i32), Val>;

/// Live value of a formula cell; a never-evaluated formula reads as blank
/// (IronCalc's "#ERROR!" placeholder is not a value the grid ever saw).
pub fn live_value(m: &Model, sheet: u32, row: i32, col: i32) -> Val {
    if is_unevaluated(m, sheet, row, col) {
        return Val::None;
    }
    m.get_cell_value_by_index(sheet, row, col).map(Val::from_cell_value).unwrap_or(Val::None)
}

/// Every formula cell's current value — the "before" side of a diff.
pub fn formula_values(m: &Model) -> ValueMap {
    let mut out = ValueMap::new();
    for (sheet_index, ws) in m.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_none() {
                    continue;
                }
                let key = (sheet_index as u32, *row, *column);
                out.insert(key, live_value(m, key.0, key.1, key.2));
            }
        }
    }
    out
}

/// Formula cells whose value differs from `before`, streaming over the
/// model (no "after" map). `remap` translates a BEFORE coordinate to its
/// AFTER coordinate (None = the cell is gone); the grid performed the same
/// structural change itself, so comparing at after-coordinates is exact.
pub fn changed_since(
    before: ValueMap,
    m: &Model,
    remap: impl Fn((u32, i32, i32)) -> Option<(u32, i32, i32)>,
) -> Vec<((u32, i32, i32), Val)> {
    let mut prior: ValueMap = HashMap::with_capacity(before.len());
    for (key, value) in before {
        if let Some(after_key) = remap(key) {
            prior.insert(after_key, value);
        }
    }
    let mut out = Vec::new();
    for (sheet_index, ws) in m.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_none() {
                    continue;
                }
                let key = (sheet_index as u32, *row, *column);
                let now = live_value(m, key.0, key.1, key.2);
                let same = match prior.get(&key) {
                    Some(prev) => equivalent(prev, &now),
                    // Not a formula before (or absent): unchanged only if blank now.
                    None => now.is_blank(),
                };
                if !same {
                    out.push((key, now));
                }
            }
        }
    }
    out
}

/// After a deferred first evaluate: formula cells whose live value differs
/// from the file's cached value. No before-map at all — the cache IS the
/// before side, and a cell the file cached nothing for counts as changed
/// only when it now holds something.
pub fn changed_vs_cached(m: &Model, cached: &crate::Cached) -> Vec<((u32, i32, i32), Val)> {
    let mut out = Vec::new();
    for (sheet_index, ws) in m.workbook.worksheets.iter().enumerate() {
        for (row, columns) in &ws.sheet_data {
            for (column, cell) in columns {
                if cell.get_formula().is_none() {
                    continue;
                }
                let key = (sheet_index as u32, *row, *column);
                let now = live_value(m, key.0, key.1, key.2);
                let same = match cached.get(&key) {
                    Some(saved) => equivalent(&Val::from_json(saved), &now),
                    None => now.is_blank(),
                };
                if !same {
                    out.push((key, now));
                }
            }
        }
    }
    out
}

/// The file-cache map exists for one reason: a formula our engine cannot
/// evaluate (add-in call, external link) shows the file's last value with
/// a `fileSaved` flag. Once the model is evaluated, every entry whose live
/// value is usable is dead weight — 88 MB of it on a 900k-formula model.
/// Keep only the entries that can still be served.
pub fn prune_cached(m: &Model, cached: &mut crate::Cached) -> usize {
    let before = cached.len();
    cached.retain(|&(sheet, row, col), _| {
        let live = m.get_cell_value_by_index(sheet, row, col).unwrap_or(CellValue::None);
        crate::value_missing(&live)
    });
    cached.shrink_to_fit();
    before - cached.len()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn equivalence_rules() {
        assert!(equivalent(&Val::Num(1.0), &Val::Num(1.0 + 1e-12)));
        assert!(!equivalent(&Val::Num(1.0), &Val::Num(1.001)));
        assert!(equivalent(&Val::None, &Val::Text(String::new())));
        assert!(!equivalent(&Val::None, &Val::Num(0.0)));
        assert!(!equivalent(&Val::Text("a".into()), &Val::Text("b".into())));
    }

    #[test]
    fn settle_diff_and_prune() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        m.set_user_input(0, 1, 1, "2".to_string()).unwrap();
        m.set_user_input(0, 1, 2, "=A1*3".to_string()).unwrap();
        m.set_user_input(0, 2, 2, "=IF(A1>5,A1,\"\")".to_string()).unwrap();
        // Pretend the file cached 6 for B1 (matches) and nothing for B2.
        let mut cached = crate::Cached::new();
        cached.insert((0, 1, 2), serde_json::json!(6.0));
        m.evaluate();
        let changed = changed_vs_cached(&m, &cached);
        assert!(changed.is_empty(), "{changed:?}"); // B1 equal, B2 blank with no cache
        // A cache that disagrees is reported.
        cached.insert((0, 1, 2), serde_json::json!(7.0));
        let changed = changed_vs_cached(&m, &cached);
        assert_eq!(changed, vec![((0, 1, 2), Val::Num(6.0))]);
        // Prune keeps nothing: every live value is usable.
        let dropped = prune_cached(&m, &mut cached);
        assert_eq!(dropped, 1);
        assert!(cached.is_empty());
    }

    #[test]
    fn structural_diff_streams_and_remaps() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        for r in 1..=3 {
            m.set_user_input(0, r, 1, format!("{r}")).unwrap();
            m.set_user_input(0, r, 2, format!("=A{r}*10")).unwrap();
        }
        m.set_user_input(0, 5, 1, "=SUM(A1:A3)".to_string()).unwrap();
        m.evaluate();
        let before = formula_values(&m);
        m.insert_rows(0, 2, 1).unwrap();
        m.evaluate();
        let changed = changed_since(before, &m, |(s, r, c)| Some(if s == 0 && r >= 2 { (s, r + 1, c) } else { (s, r, c) }));
        assert!(changed.is_empty(), "{changed:?}");
        let before = formula_values(&m);
        m.delete_rows(0, 3, 1).unwrap(); // the row holding A=2
        m.evaluate();
        let changed = changed_since(before, &m, |(s, r, c)| if s != 0 { Some((s, r, c)) } else if r == 3 { None } else if r > 3 { Some((s, r - 1, c)) } else { Some((s, r, c)) });
        assert_eq!(changed, vec![((0, 5, 1), Val::Num(4.0))]);
    }
}
