//! Temporary diagnostic: load an xlsx, recalc with IronCalc, and report
//! formula cells whose computed value is an error while the file's cached
//! value (what Excel last computed) was not an error.
//!
//! Usage: cargo run --release --example debug_diverr -- /path/to/file.xlsx [Sheet!A1 ...]
//!
//! Extra `Sheet!A1` arguments print that cell's computed value vs the file's
//! cached value (for chasing wrong values that IFERROR masks as non-errors).

use ironcalc::base::expressions::utils::number_to_column;
use ironcalc::base::types::{Cell, FormulaValue};
use ironcalc::import::load_from_xlsx;
use std::collections::HashMap;

fn cached_repr(v: &FormulaValue) -> String {
    match v {
        FormulaValue::Unevaluated => "<uneval>".to_string(),
        FormulaValue::Boolean(b) => b.to_string(),
        FormulaValue::Number(n) => n.to_string(),
        FormulaValue::Text(t) => format!("{t:?}"),
        FormulaValue::Error { ei, .. } => format!("{ei}"),
    }
}

fn is_err(v: &FormulaValue) -> bool {
    matches!(v, FormulaValue::Error { .. })
}

fn main() {
    let path = std::env::args().nth(1).expect("usage: debug_diverr <file.xlsx>");
    let mut model = load_from_xlsx(&path, "en", "UTC", "en").expect("load failed");

    // Snapshot cached formula values before evaluation overwrites them.
    let mut cached: HashMap<(usize, i32, i32), String> = HashMap::new();
    let mut cached_err: HashMap<(usize, i32, i32), bool> = HashMap::new();
    for (si, ws) in model.workbook.worksheets.iter().enumerate() {
        for (row, cols) in &ws.sheet_data {
            for (col, cell) in cols {
                match cell {
                    Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => {
                        cached.insert((si, *row, *col), cached_repr(v));
                        cached_err.insert((si, *row, *col), is_err(v));
                    }
                    _ => {}
                }
            }
        }
    }

    let t0 = std::time::Instant::now();
    model.evaluate();
    eprintln!("evaluate: {:?}", t0.elapsed());

    // Collect mismatches: computed error, cached non-error.
    let sheet_names: Vec<String> = model
        .workbook
        .worksheets
        .iter()
        .map(|w| w.name.clone())
        .collect();
    let mut rows: Vec<(usize, i32, i32, String, String)> = Vec::new();
    for (si, ws) in model.workbook.worksheets.iter().enumerate() {
        for (row, cols) in &ws.sheet_data {
            for (col, cell) in cols {
                let v = match cell {
                    Cell::CellFormula { v, .. } | Cell::ArrayFormula { v, .. } => v,
                    _ => continue,
                };
                if let FormulaValue::Error { ei, o, m } = v {
                    let was_err = cached_err.get(&(si, *row, *col)).copied().unwrap_or(false);
                    if !was_err {
                        rows.push((
                            si,
                            *row,
                            *col,
                            format!("{ei} (origin {o}, {m})"),
                            cached
                                .get(&(si, *row, *col))
                                .cloned()
                                .unwrap_or_default(),
                        ));
                    }
                }
            }
        }
    }
    // Extra args: print computed vs cached for specific cells.
    for arg in std::env::args().skip(2) {
        let Some((sheet_name, cellref)) = arg.rsplit_once('!') else {
            continue;
        };
        let Some(si) = sheet_names.iter().position(|n| n == sheet_name) else {
            println!("inspect {arg}: sheet not found");
            continue;
        };
        let letters: String = cellref.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let row: i32 = cellref[letters.len()..].parse().unwrap_or(0);
        let col = ironcalc::base::expressions::utils::column_to_number(&letters).unwrap_or(0);
        let computed = model
            .get_formatted_cell_value(si as u32, row, col)
            .unwrap_or_else(|e| format!("<{e}>"));
        let formula = model
            .get_localized_cell_content(si as u32, row, col)
            .unwrap_or_default();
        let was = cached
            .get(&(si, row, col))
            .cloned()
            .unwrap_or_else(|| "<not a formula cell>".to_string());
        println!("inspect {arg}: computed={computed} cached={was} :: {formula}");
    }

    rows.sort();
    println!("--- {} new error cells ---", rows.len());
    for (si, row, col, err, was) in &rows {
        let colname = number_to_column(*col).unwrap_or_else(|| format!("C{col}"));
        let formula = model
            .get_localized_cell_content(*si as u32, *row, *col)
            .unwrap_or_default();
        println!(
            "{}!{}{} = {} (cached {}) :: {}",
            sheet_names[*si], colname, row, err, was,
            formula.chars().take(180).collect::<String>()
        );
    }
}
