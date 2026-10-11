//! `query_rows`: filter / project / aggregate the rows of one sheet, served
//! from the model the engine already holds. Zero copies, millisecond-class
//! on 800k-cell sheets — the answer to "which rows have AWS in column DB",
//! a question the agent otherwise asks through chunked `read_range` calls
//! or a script that first ships the whole sheet to a sandbox.
//!
//! Input (JSON, see `QueryInput`):
//! ```json
//! {"sheet": 3, "where": [{"column": "Owner", "op": "eq", "value": "AWS"}],
//!  "columns": ["A", "Owner", "Capacity (MW)"], "limit": 50,
//!  "group_by": "Market", "aggregate": [{"column": "Capacity (MW)", "fn": "sum"}]}
//! ```
//! Columns are column letters or header texts (resolved on the detected or
//! given header row). Output rows carry 1-based row numbers and cell values
//! keyed by column letter; date-formatted numbers come back as ISO dates.

use crate::index::{col_letters, detect_header_row, letters_to_col, serial_to_iso_date, StyleFacts};
use ironcalc_base::cell::CellValue;
use ironcalc_base::types::{Cell, FormulaValue, SpillValue};
use ironcalc_base::Model;
use serde::Deserialize;
use serde_json::{json, Map, Value};
use std::collections::HashMap;

pub const DEFAULT_LIMIT: usize = 50;
pub const MAX_LIMIT: usize = 500;
/// Hard cap on cells per reply (rows × columns); the limit shrinks to fit.
pub const MAX_CELLS: usize = 10_000;
pub const MAX_COLUMNS: usize = 60;
/// Default projection when `columns` is omitted: the header columns (or
/// A.. when there is no header row), at most this many.
pub const DEFAULT_COLUMNS: usize = 40;
pub const MAX_PREDICATES: usize = 12;

#[derive(Deserialize, Debug, Default)]
pub struct QueryInput {
    pub sheet: u32,
    pub header_row: Option<i32>,
    #[serde(rename = "where", default)]
    pub predicates: Vec<Predicate>,
    pub columns: Option<Vec<String>>,
    pub start_row: Option<i32>,
    pub end_row: Option<i32>,
    pub limit: Option<usize>,
    pub group_by: Option<String>,
    #[serde(default)]
    pub aggregate: Vec<Aggregate>,
    pub order_by: Option<OrderBy>,
    #[serde(default)]
    pub include_formulas: bool,
}

#[derive(Deserialize, Debug)]
pub struct Predicate {
    pub column: String,
    pub op: String,
    #[serde(default)]
    pub value: Value,
}

#[derive(Deserialize, Debug)]
pub struct Aggregate {
    pub column: String,
    #[serde(rename = "fn")]
    pub func: String,
}

#[derive(Deserialize, Debug)]
pub struct OrderBy {
    pub column: String,
    #[serde(default)]
    pub desc: bool,
}

/// A cell's value for comparison and output.
#[derive(Debug, Clone, PartialEq)]
enum V {
    None,
    Num(f64),
    Bool(bool),
    Text(String),
}

impl V {
    fn present(&self) -> bool {
        match self {
            V::None => false,
            V::Text(t) => !t.is_empty(),
            _ => true,
        }
    }
    fn as_num(&self) -> Option<f64> {
        match self {
            V::Num(n) => Some(*n),
            V::Text(t) => t.trim().replace(',', "").parse::<f64>().ok(),
            V::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
            V::None => None,
        }
    }
    fn text(&self, date_fmt: bool) -> String {
        match self {
            V::None => String::new(),
            V::Num(n) => {
                if date_fmt {
                    if let Some(iso) = serial_to_iso_date(*n) {
                        return iso;
                    }
                }
                num_text(*n)
            }
            V::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
            V::Text(t) => t.clone(),
        }
    }
    fn to_json(&self, date_fmt: bool) -> Value {
        match self {
            V::None => Value::Null,
            V::Num(n) => {
                if date_fmt {
                    if let Some(iso) = serial_to_iso_date(*n) {
                        return Value::String(iso);
                    }
                }
                serde_json::Number::from_f64(*n).map(Value::Number).unwrap_or_else(|| Value::String(n.to_string()))
            }
            V::Bool(b) => Value::Bool(*b),
            V::Text(t) => Value::String(t.clone()),
        }
    }
}

fn num_text(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

fn cell_v(m: &Model, sheet: u32, r: i32, c: i32, cell: &Cell) -> V {
    let shared = &m.workbook.shared_strings;
    match cell {
        Cell::EmptyCell { .. } => V::None,
        Cell::BooleanCell { v, .. } => V::Bool(*v),
        Cell::NumberCell { v, .. } => V::Num(*v),
        Cell::SharedString { si, .. } => shared.get(*si as usize).map(|s| V::Text(s.clone())).unwrap_or(V::None),
        Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => match v {
            FormulaValue::Boolean(b) => V::Bool(*b),
            FormulaValue::Number(n) => V::Num(*n),
            FormulaValue::Text(t) => V::Text(t.clone()),
            FormulaValue::Unevaluated => V::None,
            FormulaValue::Error { .. } => localized(m, sheet, r, c),
        },
        Cell::SpillCell { v, .. } => match v {
            SpillValue::Boolean(b) => V::Bool(*b),
            SpillValue::Number(n) => V::Num(*n),
            SpillValue::Text(t) => V::Text(t.clone()),
            SpillValue::Error(_) => localized(m, sheet, r, c),
        },
        Cell::ErrorCell { .. } => localized(m, sheet, r, c),
    }
}

fn localized(m: &Model, sheet: u32, r: i32, c: i32) -> V {
    match m.get_cell_value_by_index(sheet, r, c) {
        Ok(CellValue::String(s)) => V::Text(s),
        Ok(CellValue::Number(n)) => V::Num(n),
        Ok(CellValue::Boolean(b)) => V::Bool(b),
        _ => V::None,
    }
}

/// "YYYY-MM-DD" → Excel serial (1900 system, Excel's day count including the
/// phantom 29 Feb 1900 for dates after it).
pub fn iso_to_serial(s: &str) -> Option<f64> {
    let b = s.as_bytes();
    if b.len() != 10 || b[4] != b'-' || b[7] != b'-' {
        return None;
    }
    let y: i64 = s[0..4].parse().ok()?;
    let mo: i64 = s[5..7].parse().ok()?;
    let d: i64 = s[8..10].parse().ok()?;
    if !(1..=12).contains(&mo) || !(1..=31).contains(&d) {
        return None;
    }
    // days_from_civil (Howard Hinnant)
    let y2 = if mo <= 2 { y - 1 } else { y };
    let era = if y2 >= 0 { y2 } else { y2 - 399 } / 400;
    let yoe = y2 - era * 400;
    let mp = if mo > 2 { mo - 3 } else { mo + 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    let days_since_epoch = era * 146_097 + doe - 719_468;
    let serial = days_since_epoch + 25_569;
    if serial < 1 {
        return None;
    }
    Some(serial as f64)
}

/// A criterion value, pre-normalised once per predicate.
#[derive(Debug, Clone)]
struct Crit {
    raw: Value,
    text: String, // lowercase trimmed
    num: Option<f64>,
    boolean: Option<bool>,
}

impl Crit {
    fn from(v: &Value) -> Crit {
        let (text, num, boolean) = match v {
            Value::Number(n) => (n.to_string(), n.as_f64(), None),
            Value::Bool(b) => (b.to_string(), None, Some(*b)),
            Value::String(s) => {
                let t = s.trim();
                let num = t.replace(',', "").parse::<f64>().ok().or_else(|| iso_to_serial(t));
                let boolean = match t.to_ascii_lowercase().as_str() {
                    "true" => Some(true),
                    "false" => Some(false),
                    _ => None,
                };
                (t.to_lowercase(), num, boolean)
            }
            Value::Null => (String::new(), None, None),
            other => (other.to_string().to_lowercase(), None, None),
        };
        Crit { raw: v.clone(), text, num, boolean }
    }
}

fn eq(v: &V, crit: &Crit, date_fmt: bool) -> bool {
    match v {
        V::None => crit.text.is_empty() && crit.raw.is_null() || crit.text.is_empty(),
        V::Num(n) => {
            if let Some(c) = crit.num {
                return (n - c).abs() <= 1e-9 * n.abs().max(1.0);
            }
            v.text(date_fmt).to_lowercase() == crit.text
        }
        V::Bool(b) => match crit.boolean {
            Some(c) => *b == c,
            None => crit.num.is_some_and(|c| (c != 0.0) == *b),
        },
        V::Text(t) => {
            let tl = t.trim().to_lowercase();
            if tl == crit.text {
                return true;
            }
            match (crit.num, tl.replace(',', "").parse::<f64>()) {
                (Some(c), Ok(n)) => (n - c).abs() <= 1e-9 * n.abs().max(1.0),
                _ => false,
            }
        }
    }
}

fn cmp_order(v: &V, crit: &Crit) -> Option<std::cmp::Ordering> {
    if !v.present() {
        return None;
    }
    if let (Some(a), Some(b)) = (v.as_num(), crit.num) {
        return a.partial_cmp(&b);
    }
    match v {
        V::Text(t) => Some(t.trim().to_lowercase().cmp(&crit.text)),
        _ => None,
    }
}

fn matches(v: &V, op: &str, crit: &Crit, crit_list: &[Crit], date_fmt: bool) -> Result<bool, String> {
    use std::cmp::Ordering::*;
    Ok(match op {
        "empty" | "is_empty" | "blank" => !v.present(),
        "not_empty" | "is_not_empty" | "not_blank" => v.present(),
        "eq" | "=" | "==" | "equals" | "is" => eq(v, crit, date_fmt),
        "ne" | "!=" | "<>" | "not" | "not_equals" => !eq(v, crit, date_fmt),
        "contains" => v.present() && v.text(date_fmt).to_lowercase().contains(&crit.text),
        "not_contains" => !(v.present() && v.text(date_fmt).to_lowercase().contains(&crit.text)),
        "starts_with" => v.present() && v.text(date_fmt).to_lowercase().starts_with(&crit.text),
        "ends_with" => v.present() && v.text(date_fmt).to_lowercase().ends_with(&crit.text),
        "gt" | ">" => matches!(cmp_order(v, crit), Some(Greater)),
        "gte" | ">=" => matches!(cmp_order(v, crit), Some(Greater | Equal)),
        "lt" | "<" => matches!(cmp_order(v, crit), Some(Less)),
        "lte" | "<=" => matches!(cmp_order(v, crit), Some(Less | Equal)),
        "between" => {
            if crit_list.len() != 2 {
                return Err("`between` needs value: [low, high]".into());
            }
            matches!(cmp_order(v, &crit_list[0]), Some(Greater | Equal)) && matches!(cmp_order(v, &crit_list[1]), Some(Less | Equal))
        }
        "in" => crit_list.iter().any(|c| eq(v, c, date_fmt)),
        "not_in" => !crit_list.iter().any(|c| eq(v, c, date_fmt)),
        other => return Err(format!("unknown op \"{other}\" (use eq, ne, contains, not_contains, starts_with, ends_with, gt, gte, lt, lte, between, in, not_in, empty, not_empty)")),
    })
}

/// Header texts of a row, every column (the index's header map skips column
/// A and clips; resolution wants the full row).
fn header_texts(m: &Model, sheet: u32, row: i32, facts: &mut StyleFacts) -> Vec<(i32, String)> {
    let ws = &m.workbook.worksheets[sheet as usize];
    let Some(cells) = ws.sheet_data.get(&row) else { return Vec::new() };
    let mut cols: Vec<&i32> = cells.keys().collect();
    cols.sort_unstable();
    let mut out = Vec::new();
    for c in cols {
        let cell = &cells[c];
        let v = cell_v(m, sheet, row, *c, cell);
        if !v.present() {
            continue;
        }
        let date_fmt = matches!(v, V::Num(_)) && facts.is_date(m, sheet, row, *c, cell.get_style());
        let t = v.text(date_fmt).trim().to_string();
        if !t.is_empty() {
            out.push((*c, t));
        }
    }
    out
}

/// Column spec → 1-based column. Header text first (exact, case-insensitive),
/// then column letters, then a unique substring match on headers.
fn resolve_column(spec: &str, headers: &[(i32, String)]) -> Result<i32, String> {
    let s = spec.trim();
    if s.is_empty() {
        return Err("empty column spec".into());
    }
    let sl = s.to_lowercase();
    if let Some((c, _)) = headers.iter().find(|(_, h)| h.to_lowercase() == sl) {
        return Ok(*c);
    }
    if let Some(c0) = letters_to_col(s) {
        return Ok(c0 + 1);
    }
    let subs: Vec<&(i32, String)> = headers.iter().filter(|(_, h)| h.to_lowercase().contains(&sl)).collect();
    match subs.len() {
        1 => Ok(subs[0].0),
        0 => {
            let shown: Vec<String> = headers.iter().take(60).map(|(c, h)| format!("{}={h:?}", col_letters(c - 1))).collect();
            Err(format!("unknown column {s:?}; use a column letter or one of the header texts: {}", shown.join(", ")))
        }
        _ => {
            let shown: Vec<String> = subs.iter().take(10).map(|(c, h)| format!("{}={h:?}", col_letters(c - 1))).collect();
            Err(format!("column {s:?} is ambiguous: {}", shown.join(", ")))
        }
    }
}

struct Pred {
    col: i32,
    op: String,
    crit: Crit,
    crit_list: Vec<Crit>,
}

/// Run a query; returns the reply JSON.
pub fn query_rows_json(m: &Model, input_json: &str) -> Result<String, String> {
    let input: QueryInput = serde_json::from_str(input_json).map_err(|e| format!("bad query: {e}"))?;
    query_rows(m, &input).map(|v| v.to_string())
}

/// Timing is the caller's business: `std::time::Instant` traps on
/// wasm32-unknown-unknown, and this runs there too.
pub fn query_rows(m: &Model, q: &QueryInput) -> Result<Value, String> {
    let sheet = q.sheet;
    let ws = m.workbook.worksheets.get(sheet as usize).ok_or_else(|| format!("sheet index {sheet} out of range"))?;
    if q.predicates.len() > MAX_PREDICATES {
        return Err(format!("at most {MAX_PREDICATES} predicates"));
    }
    let mut facts = StyleFacts::default();

    // Header row: given, else detected.
    let (header_row, detected) = match q.header_row {
        Some(r) if r >= 1 => (Some(r), false),
        _ => (detect_header_row(m, sheet, &mut facts).map(|(r, _)| r), true),
    };
    let headers: Vec<(i32, String)> = header_row.map(|r| header_texts(m, sheet, r, &mut facts)).unwrap_or_default();
    let header_of = |c: i32| headers.iter().find(|(hc, _)| *hc == c).map(|(_, h)| h.clone());

    // Predicates.
    let mut preds: Vec<Pred> = Vec::with_capacity(q.predicates.len());
    for p in &q.predicates {
        let col = resolve_column(&p.column, &headers)?;
        let crit_list: Vec<Crit> = match &p.value {
            Value::Array(items) => items.iter().map(Crit::from).collect(),
            _ => Vec::new(),
        };
        preds.push(Pred { col, op: p.op.trim().to_lowercase(), crit: Crit::from(&p.value), crit_list });
    }
    // Validate ops up front with a neutral value so a typo fails fast.
    for p in &preds {
        matches(&V::None, &p.op, &p.crit, &p.crit_list, false)?;
    }

    // Row range.
    let mut row_keys: Vec<i32> = ws.sheet_data.keys().copied().collect();
    row_keys.sort_unstable();
    let max_row = row_keys.last().copied().unwrap_or(0);
    let start = q.start_row.unwrap_or_else(|| header_row.map(|h| h + 1).unwrap_or(1)).max(1);
    let end = q.end_row.unwrap_or(max_row).min(max_row);

    // Projection.
    let max_col = ws.sheet_data.values().flat_map(|r| r.keys()).copied().max().unwrap_or(1);
    let cols: Vec<i32> = match &q.columns {
        Some(specs) => {
            if specs.len() > MAX_COLUMNS {
                return Err(format!("at most {MAX_COLUMNS} columns"));
            }
            let mut out = Vec::new();
            for s in specs {
                let c = resolve_column(s, &headers)?;
                if !out.contains(&c) {
                    out.push(c);
                }
            }
            out
        }
        None => {
            if !headers.is_empty() {
                headers.iter().map(|(c, _)| *c).take(DEFAULT_COLUMNS).collect()
            } else {
                (1..=max_col.min(DEFAULT_COLUMNS as i32)).collect()
            }
        }
    };
    let group_col = match &q.group_by {
        Some(s) if !s.trim().is_empty() => Some(resolve_column(s, &headers)?),
        _ => None,
    };
    let aggs: Vec<(i32, String)> = q
        .aggregate
        .iter()
        .map(|a| resolve_column(&a.column, &headers).map(|c| (c, a.func.trim().to_lowercase())))
        .collect::<Result<_, _>>()?;
    for (_, f) in &aggs {
        if !matches!(f.as_str(), "sum" | "count" | "min" | "max" | "avg" | "mean") {
            return Err(format!("unknown aggregate fn {f:?} (sum, count, min, max, avg)"));
        }
    }
    let order_col = match &q.order_by {
        Some(o) => Some((resolve_column(&o.column, &headers)?, o.desc)),
        None => None,
    };
    let limit = q.limit.unwrap_or(DEFAULT_LIMIT).clamp(1, MAX_LIMIT);
    let limit = limit.min((MAX_CELLS / cols.len().max(1)).max(1));

    // Style facts per column are mostly uniform; memo date-ness per (col, style).
    let mut date_memo: HashMap<i32, bool> = HashMap::new();
    let mut is_date = |m: &Model, r: i32, c: i32, cell: &Cell, facts: &mut StyleFacts| -> bool {
        let s = cell.get_style();
        *date_memo.entry(s).or_insert_with(|| facts.is_date(m, sheet, r, c, s))
    };

    // Scan.
    let mut scanned = 0usize;
    let mut matched: Vec<(i32, Vec<(i32, V, bool)>)> = Vec::new(); // row, projected cells (col, value, date)
    let mut matched_count = 0usize;
    // group key → (count, per-agg accumulators)
    let mut groups: HashMap<String, (usize, Vec<Acc>)> = HashMap::new();
    let mut group_order: Vec<String> = Vec::new();
    let empty: HashMap<i32, Cell> = HashMap::new();

    for r in row_keys.iter().copied().filter(|r| *r >= start && *r <= end) {
        let cells = ws.sheet_data.get(&r).unwrap_or(&empty);
        scanned += 1;
        let mut ok = true;
        for p in &preds {
            let (v, date_fmt) = match cells.get(&p.col) {
                Some(cell) => (cell_v(m, sheet, r, p.col, cell), matches!(cell, Cell::NumberCell { .. } | Cell::CellFormula { .. }) && is_date(m, r, p.col, cell, &mut facts)),
                None => (V::None, false),
            };
            if !matches(&v, &p.op, &p.crit, &p.crit_list, date_fmt)? {
                ok = false;
                break;
            }
        }
        if !ok {
            continue;
        }
        matched_count += 1;
        if let Some(gc) = group_col {
            let key = match cells.get(&gc) {
                Some(cell) => {
                    let v = cell_v(m, sheet, r, gc, cell);
                    let d = matches!(v, V::Num(_)) && is_date(m, r, gc, cell, &mut facts);
                    v.text(d).trim().to_string()
                }
                None => String::new(),
            };
            let entry = groups.entry(key.clone()).or_insert_with(|| {
                group_order.push(key.clone());
                (0, aggs.iter().map(|_| Acc::default()).collect())
            });
            entry.0 += 1;
            for (i, (ac, _)) in aggs.iter().enumerate() {
                let v = cells.get(ac).map(|cell| cell_v(m, sheet, r, *ac, cell)).unwrap_or(V::None);
                entry.1[i].add(&v);
            }
            continue;
        }
        // Keep every match when ordering (sort before limit); otherwise stop
        // materialising past the limit but keep counting.
        if order_col.is_none() && matched.len() >= limit {
            continue;
        }
        let projected: Vec<(i32, V, bool)> = cols
            .iter()
            .map(|c| match cells.get(c) {
                Some(cell) => {
                    let v = cell_v(m, sheet, r, *c, cell);
                    let d = matches!(v, V::Num(_)) && is_date(m, r, *c, cell, &mut facts);
                    (*c, v, d)
                }
                None => (*c, V::None, false),
            })
            .collect();
        matched.push((r, projected));
    }

    let mut out = Map::new();
    out.insert("sheet".into(), json!(ws.name));
    out.insert("header_row".into(), json!(header_row));
    if header_row.is_some() && detected {
        out.insert("header_row_detected".into(), json!(true));
    }
    out.insert("rows_scanned".into(), json!(scanned));
    out.insert("row_range".into(), json!(format!("{start}-{end}")));
    out.insert("matched".into(), json!(matched_count));

    if let Some(gc) = group_col {
        let mut rows: Vec<Value> = Vec::new();
        let mut keys: Vec<&String> = group_order.iter().collect();
        // Sort groups: by the first aggregate (desc) when there is one, else by count desc.
        keys.sort_by(|a, b| {
            let (ca, aa) = &groups[*a];
            let (cb, ab) = &groups[*b];
            let ka = aa.first().and_then(|x| x.value(&aggs[0].1)).unwrap_or(*ca as f64);
            let kb = ab.first().and_then(|x| x.value(&aggs[0].1)).unwrap_or(*cb as f64);
            kb.partial_cmp(&ka).unwrap_or(std::cmp::Ordering::Equal).then_with(|| cb.cmp(ca))
        });
        for k in keys.iter().take(limit) {
            let (count, accs) = &groups[*k];
            let mut g = Map::new();
            g.insert("key".into(), json!(if k.is_empty() { "(blank)" } else { k.as_str() }));
            g.insert("count".into(), json!(count));
            for (i, (ac, f)) in aggs.iter().enumerate() {
                let name = format!("{}_{}", f, header_of(*ac).unwrap_or_else(|| col_letters(ac - 1)));
                g.insert(name, accs[i].value(f).map(|x| json!(x)).unwrap_or(Value::Null));
            }
            rows.push(Value::Object(g));
        }
        out.insert("group_by".into(), json!({ "column": col_letters(gc - 1), "header": header_of(gc) }));
        out.insert("group_count".into(), json!(groups.len()));
        out.insert("groups".into(), Value::Array(rows));
        out.insert("truncated".into(), json!(groups.len() > limit));
    } else {
        if let Some((oc, desc)) = order_col {
            let key_of = |cells: &Vec<(i32, V, bool)>, r: i32| -> V {
                cells.iter().find(|(c, _, _)| *c == oc).map(|(_, v, _)| v.clone()).unwrap_or_else(|| {
                    ws.sheet_data.get(&r).and_then(|row| row.get(&oc)).map(|cell| cell_v(m, sheet, r, oc, cell)).unwrap_or(V::None)
                })
            };
            matched.sort_by(|(ra, a), (rb, b)| {
                let ka = key_of(a, *ra);
                let kb = key_of(b, *rb);
                // Blanks sort last in either direction.
                match (ka.present(), kb.present()) {
                    (false, false) => std::cmp::Ordering::Equal,
                    (false, true) => std::cmp::Ordering::Greater,
                    (true, false) => std::cmp::Ordering::Less,
                    _ => {
                        let ord = order_cmp(&ka, &kb);
                        if desc { ord.reverse() } else { ord }
                    }
                }
            });
        }
        let truncated = matched.len() > limit || matched_count > matched.len();
        let rows: Vec<Value> = matched
            .iter()
            .take(limit)
            .map(|(r, cells)| {
                let mut c = Map::new();
                for (col, v, d) in cells {
                    let mut val = v.to_json(*d);
                    if q.include_formulas {
                        if let Ok(Some(f)) = m.get_cell_formula(sheet, *r, *col) {
                            val = json!({ "v": val, "f": f });
                        }
                    }
                    c.insert(col_letters(col - 1), val);
                }
                json!({ "row": r, "cells": c })
            })
            .collect();
        out.insert(
            "columns".into(),
            Value::Array(cols.iter().map(|c| json!({ "column": col_letters(c - 1), "header": header_of(*c) })).collect()),
        );
        out.insert("rows".into(), Value::Array(rows));
        out.insert("truncated".into(), json!(truncated));
    }
    Ok(Value::Object(out))
}

fn order_cmp(a: &V, b: &V) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    match (a.present(), b.present()) {
        (false, false) => return Equal,
        (false, true) => return Greater, // blanks last
        (true, false) => return Less,
        _ => {}
    }
    match (a.as_num(), b.as_num()) {
        (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(Equal),
        (Some(_), None) => Less, // numbers before text
        (None, Some(_)) => Greater,
        (None, None) => a.text(false).to_lowercase().cmp(&b.text(false).to_lowercase()),
    }
}

#[derive(Default, Debug)]
struct Acc {
    n: usize,
    non_empty: usize,
    sum: f64,
    min: Option<f64>,
    max: Option<f64>,
}

impl Acc {
    fn add(&mut self, v: &V) {
        if v.present() {
            self.non_empty += 1;
        }
        if let V::Num(x) = v {
            self.n += 1;
            self.sum += x;
            self.min = Some(self.min.map_or(*x, |m| m.min(*x)));
            self.max = Some(self.max.map_or(*x, |m| m.max(*x)));
        }
    }
    fn value(&self, f: &str) -> Option<f64> {
        match f {
            "sum" => Some(self.sum),
            "count" => Some(self.non_empty as f64),
            "min" => self.min,
            "max" => self.max,
            "avg" | "mean" => if self.n > 0 { Some(self.sum / self.n as f64) } else { None },
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn model() -> Model<'static> {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        let set = |m: &mut Model, r: i32, c: i32, v: &str| m.set_user_input(0, r, c, v.to_string()).unwrap();
        set(&mut m, 1, 1, "Site");
        set(&mut m, 1, 2, "Owner");
        set(&mut m, 1, 3, "Capacity (MW)");
        set(&mut m, 1, 4, "Online");
        set(&mut m, 1, 5, "Market");
        let data = [
            ("Ashburn 1", "AWS", "120", "45292", "NoVA"),
            ("Ashburn 2", "aws ", "80", "45658", "NoVA"),
            ("Dallas 1", "Microsoft", "200", "45400", "Dallas"),
            ("Phoenix 1", "AWS", "", "45500", "Phoenix"),
            ("Phoenix 2", "Google", "50", "", "Phoenix"),
        ];
        for (i, (site, owner, cap, online, market)) in data.iter().enumerate() {
            let r = i as i32 + 2;
            set(&mut m, r, 1, site);
            set(&mut m, r, 2, owner);
            if !cap.is_empty() {
                set(&mut m, r, 3, cap);
            }
            if !online.is_empty() {
                set(&mut m, r, 4, online);
                let mut st = m.get_style_for_cell(0, r, 4).unwrap();
                st.num_fmt = "yyyy-mm-dd".into();
                m.set_cell_style(0, r, 4, &st).unwrap();
            }
            set(&mut m, r, 5, market);
        }
        set(&mut m, 8, 1, "Total");
        set(&mut m, 8, 3, "=SUM(C2:C6)");
        m.evaluate();
        m
    }

    fn run(m: &Model, q: Value) -> Value {
        serde_json::from_str(&query_rows_json(m, &q.to_string()).unwrap()).unwrap()
    }

    #[test]
    fn iso_serial_roundtrip() {
        assert_eq!(iso_to_serial("2024-01-01"), Some(45292.0));
        assert_eq!(iso_to_serial("1970-01-01"), Some(25569.0));
        assert_eq!(iso_to_serial("1900-03-01"), Some(61.0));
        assert_eq!(iso_to_serial("nope"), None);
    }

    #[test]
    fn filters_by_header_name_case_insensitively_and_projects() {
        let m = model();
        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "Owner", "op": "eq", "value": "AWS" }], "columns": ["Site", "Capacity (MW)", "D"] }));
        assert_eq!(out["header_row"], 1);
        assert_eq!(out["header_row_detected"], true);
        assert_eq!(out["matched"], 3);
        let rows = out["rows"].as_array().unwrap();
        assert_eq!(rows.len(), 3);
        assert_eq!(rows[0]["row"], 2);
        assert_eq!(rows[0]["cells"]["A"], "Ashburn 1");
        assert_eq!(rows[0]["cells"]["C"], 120.0);
        // Date-formatted serial comes back as ISO.
        assert_eq!(rows[0]["cells"]["D"], "2024-01-01");
        // Trailing-space / lowercase owner still matches.
        assert_eq!(rows[1]["row"], 3);
        // Blank capacity → null.
        assert_eq!(rows[2]["cells"]["C"], Value::Null);
        assert_eq!(out["columns"][1]["header"], "Capacity (MW)");
        assert_eq!(out["truncated"], false);
    }

    #[test]
    fn numeric_date_and_set_predicates() {
        let m = model();
        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "C", "op": "gte", "value": 100 }], "columns": ["A"] }));
        let sites: Vec<&str> = out["rows"].as_array().unwrap().iter().map(|r| r["cells"]["A"].as_str().unwrap()).collect();
        assert_eq!(sites, vec!["Ashburn 1", "Dallas 1", "Total"]);

        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "Online", "op": "gt", "value": "2024-06-01" }], "columns": ["A"] }));
        let sites: Vec<&str> = out["rows"].as_array().unwrap().iter().map(|r| r["cells"]["A"].as_str().unwrap()).collect();
        assert_eq!(sites, vec!["Ashburn 2", "Phoenix 1"]);

        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "Owner", "op": "in", "value": ["google", "Microsoft"] }, { "column": "Capacity (MW)", "op": "not_empty" }], "columns": ["A"] }));
        let sites: Vec<&str> = out["rows"].as_array().unwrap().iter().map(|r| r["cells"]["A"].as_str().unwrap()).collect();
        assert_eq!(sites, vec!["Dallas 1", "Phoenix 2"]);

        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "Capacity", "op": "empty" }], "columns": ["A"] }));
        let sites: Vec<&str> = out["rows"].as_array().unwrap().iter().map(|r| r["cells"]["A"].as_str().unwrap()).collect();
        assert_eq!(sites, vec!["Phoenix 1"]);
        // Row 8 ("Total") has a formula in C — not empty.
        assert_eq!(out["matched"], 1);
    }

    #[test]
    fn order_limit_and_formulas() {
        let m = model();
        let out = run(&m, json!({ "sheet": 0, "columns": ["A", "C"], "order_by": { "column": "C", "desc": true }, "limit": 2, "include_formulas": true }));
        assert_eq!(out["matched"], 6);
        assert_eq!(out["truncated"], true);
        let rows = out["rows"].as_array().unwrap();
        assert_eq!(rows.len(), 2);
        // Total row (450) first, then Dallas (200).
        assert_eq!(rows[0]["cells"]["A"], "Total");
        assert_eq!(rows[0]["cells"]["C"]["f"], "=SUM(C2:C6)");
        assert_eq!(rows[0]["cells"]["C"]["v"], 450.0);
        assert_eq!(rows[1]["cells"]["A"], "Dallas 1");
    }

    #[test]
    fn group_by_with_aggregates() {
        let m = model();
        let out = run(&m, json!({ "sheet": 0, "end_row": 6, "group_by": "Market", "aggregate": [{ "column": "Capacity (MW)", "fn": "sum" }, { "column": "A", "fn": "count" }] }));
        let groups = out["groups"].as_array().unwrap();
        assert_eq!(out["group_count"], 3);
        assert_eq!(groups[0]["key"], "NoVA");
        assert_eq!(groups[0]["count"], 2);
        assert_eq!(groups[0]["sum_Capacity (MW)"], 200.0);
        assert_eq!(groups[0]["count_Site"], 2.0);
        assert_eq!(groups[1]["key"], "Dallas");
        assert_eq!(groups[2]["key"], "Phoenix");
        assert_eq!(groups[2]["sum_Capacity (MW)"], 50.0);
    }

    #[test]
    fn errors_are_actionable() {
        let m = model();
        let err = query_rows_json(&m, &json!({ "sheet": 0, "where": [{ "column": "Ownr", "op": "eq", "value": "x" }] }).to_string()).unwrap_err();
        assert!(err.contains("unknown column"), "{err}");
        assert!(err.contains("B=\"Owner\""), "{err}");
        let err = query_rows_json(&m, &json!({ "sheet": 0, "where": [{ "column": "B", "op": "like", "value": "x" }] }).to_string()).unwrap_err();
        assert!(err.contains("unknown op"), "{err}");
        let err = query_rows_json(&m, &json!({ "sheet": 9 }).to_string()).unwrap_err();
        assert!(err.contains("out of range"), "{err}");
    }

    #[test]
    fn no_header_row_falls_back_to_letters() {
        let mut m = Model::new_empty("wb", "en", "UTC", "en").unwrap();
        for r in 1..=5 {
            m.set_user_input(0, r, 1, format!("{}", r * 10)).unwrap();
        }
        m.evaluate();
        let out = run(&m, json!({ "sheet": 0, "where": [{ "column": "A", "op": "between", "value": [20, 40] }] }));
        assert_eq!(out["header_row"], Value::Null);
        assert_eq!(out["matched"], 3);
        assert_eq!(out["rows"][0]["row"], 2);
    }
}
