//! Where the engine disagrees with the file's cached values after the first
//! evaluate: `cargo run --release --example cache_diff -- file.xlsx`
use ironcalc_base::cell::CellValue;
use std::collections::BTreeMap;
fn main() {
    let path = std::env::args().nth(1).expect("file");
    let bytes = std::fs::read(&path).unwrap();
    let wb = ironcalc::import::load_from_xlsx_bytes(&bytes, "workbook", "en", "UTC").unwrap();
    let mut m = ironcalc_base::Model::from_workbook(wb, "en").unwrap();
    let cached = gridpath_engine_core::build_cached(&m);
    m.evaluate();
    let names = m.workbook.get_worksheet_names();
    let mut by_kind: BTreeMap<&str, usize> = BTreeMap::new();
    let mut by_sheet: BTreeMap<String, usize> = BTreeMap::new();
    let mut samples: BTreeMap<&str, Vec<String>> = BTreeMap::new();
    let mut total = 0usize;
    for (si, ws) in m.workbook.worksheets.iter().enumerate() {
        for (r, row) in &ws.sheet_data {
            for (c, cell) in row {
                if cell.get_formula().is_none() { continue; }
                total += 1;
                let live = m.get_cell_value_by_index(si as u32, *r, *c).unwrap();
                let saved = cached.get(&(si as u32, *r, *c));
                let live_j = gridpath_engine_core::cell_value_json(m.get_cell_value_by_index(si as u32, *r, *c).unwrap());
                let same = match saved { Some(s) => *s == live_j, None => live_j.is_null() };
                if same { continue; }
                let kind: &str = match (saved, &live) {
                    (None, _) => "no cached value",
                    (Some(s), CellValue::Number(n)) if s.is_f64() => {
                        let sv = s.as_f64().unwrap();
                        if (sv - n).abs() <= 1e-9 * sv.abs().max(1.0) { "float noise" } else { "number differs" }
                    }
                    (Some(s), CellValue::String(t)) if t.starts_with('#') && s.is_f64() => "error, cache had number",
                    (Some(s), CellValue::String(t)) if t.starts_with('#') && s.as_str().map_or(false, |x| x.starts_with('#')) => "error differs",
                    (Some(s), CellValue::Number(_)) if s.as_str().map_or(false, |x| x.starts_with('#')) => "number, cache had error",
                    (Some(s), CellValue::String(_)) if s.is_string() => "text differs",
                    (Some(_), CellValue::None) => "empty, cache had value",
                    (Some(s), CellValue::String(_)) if s.is_null() => "text, cache empty",
                    _ => "type differs",
                };
                *by_kind.entry(kind).or_default() += 1;
                *by_sheet.entry(names[si].clone()).or_default() += 1;
                let v = samples.entry(kind).or_default();
                if v.len() < 4 {
                    let f = m.get_cell_formula(si as u32, *r, *c).unwrap().unwrap_or_default();
                    v.push(format!("    {}!R{}C{}: cache={} live={:?} f={}", names[si], r, c, saved.map(|x| x.to_string()).unwrap_or("-".into()), live, f.chars().take(90).collect::<String>()));
                }
            }
        }
    }
    eprintln!("formula cells {total}, differing {}", by_kind.values().sum::<usize>());
    for (k, n) in &by_kind { eprintln!("{n:>7}  {k}"); for s in &samples[k] { eprintln!("{s}"); } }
    eprintln!("-- by sheet"); for (s, n) in &by_sheet { eprintln!("{n:>7}  {s}"); }
}
