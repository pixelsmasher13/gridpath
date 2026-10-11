//! Error cells on one sheet, with formula and the file's cached value:
//! `cargo run --release --example sheet_errors -- file.xlsx "Sheet"`
use ironcalc_base::cell::CellValue;
use std::collections::BTreeMap;
fn col_letters(c0: i32) -> String { let mut n = c0; let mut s = Vec::new(); loop { s.push(b'A' + (n % 26) as u8); n = n / 26 - 1; if n < 0 { break; } } s.reverse(); String::from_utf8(s).unwrap() }
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap(); let sheet = a.next().unwrap();
    let bytes = std::fs::read(&path).unwrap();
    let wb = ironcalc::import::load_from_xlsx_bytes(&bytes, "workbook", "en", "UTC").unwrap();
    let mut m = ironcalc_base::Model::from_workbook(wb, "en").unwrap();
    let cached = gridpath_engine_core::build_cached(&m);
    m.evaluate();
    let names = m.workbook.get_worksheet_names();
    let si = names.iter().position(|n| *n == sheet).expect("sheet") as u32;
    let ws = &m.workbook.worksheets[si as usize];
    let mut by_err: BTreeMap<String, Vec<(i32, i32)>> = BTreeMap::new();
    let mut by_formula: BTreeMap<String, usize> = BTreeMap::new();
    let mut total_formulas = 0;
    let mut excel_also_error = 0;
    let mut rows: Vec<&i32> = ws.sheet_data.keys().collect(); rows.sort();
    for r in rows {
        let mut cols: Vec<&i32> = ws.sheet_data[r].keys().collect(); cols.sort();
        for c in cols {
            let cell = &ws.sheet_data[r][c];
            if cell.get_formula().is_none() { continue; }
            total_formulas += 1;
            if let Ok(CellValue::String(s)) = m.get_cell_value_by_index(si, *r, *c) {
                if s.starts_with('#') {
                    by_err.entry(s.clone()).or_default().push((*r, *c));
                    let f = m.get_cell_formula(si, *r, *c).unwrap().unwrap_or_default();
                    // normalise row numbers to see the pattern
                    let pat = regex_lite(&f);
                    *by_formula.entry(format!("{s} :: {pat}")).or_default() += 1;
                    if cached.get(&(si, *r, *c)).and_then(|v| v.as_str()).map_or(false, |x| x.starts_with('#')) { excel_also_error += 1; }
                }
            }
        }
    }
    eprintln!("{sheet}: {total_formulas} formulas, {} error cells ({excel_also_error} were errors in Excel's cache too)", by_err.values().map(|v| v.len()).sum::<usize>());
    for (e, cells) in &by_err {
        let shown: Vec<String> = cells.iter().take(12).map(|(r, c)| format!("{}{}", col_letters(c - 1), r)).collect();
        eprintln!("{:>5} × {e}: {}{}", cells.len(), shown.join(" "), if cells.len() > 12 { " …" } else { "" });
    }
    eprintln!("-- distinct formula patterns (digits → #):");
    let mut pats: Vec<(&String, &usize)> = by_formula.iter().collect(); pats.sort_by(|a, b| b.1.cmp(a.1));
    for (p, n) in pats.iter().take(25) { eprintln!("{n:>5}  {}", p.chars().take(220).collect::<String>()); }
    // Show a few concrete cells with full formula + cached value
    eprintln!("-- samples:");
    for (e, cells) in &by_err { for (r, c) in cells.iter().take(3) {
        let f = m.get_cell_formula(si, *r, *c).unwrap().unwrap_or_default();
        eprintln!("  {}{} {e}  cache={:?}\n     {}", col_letters(c - 1), r, cached.get(&(si, *r, *c)), f);
    } }
}
fn regex_lite(f: &str) -> String { f.chars().map(|ch| if ch.is_ascii_digit() { '#' } else { ch }).collect() }
