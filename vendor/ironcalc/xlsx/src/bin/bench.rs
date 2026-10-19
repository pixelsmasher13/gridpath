#![allow(clippy::unwrap_used, clippy::panic)]

//! Timing + fidelity bench for gridpath evaluation.
//! Usage: bench file.xlsx [sheet_index row col new_value]

use ironcalc::compare::compare;
use ironcalc::import::load_from_xlsx;
use std::time::Instant;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let file = &args[1];

    let t = Instant::now();
    let mut model = load_from_xlsx(file, "en", "UTC", "en").unwrap();
    println!("load: {} ms", t.elapsed().as_millis());

    let t = Instant::now();
    model.evaluate();
    println!("full evaluate #1: {} ms", t.elapsed().as_millis());

    let t = Instant::now();
    model.evaluate();
    println!("full evaluate #2: {} ms", t.elapsed().as_millis());

    let cached = load_from_xlsx(file, "en", "UTC", "en").unwrap();
    match compare(&cached, &model) {
        Ok(diffs) => {
            println!("fidelity: {} cells differ from cached xlsx values", diffs.len());
            for d in diffs.iter() {
                println!(
                    "DIFF\t{}\t{}\t{}\t{}\t{:?}\t{:?}",
                    d.sheet_name, d.row, d.column, d.reason, d.value1, d.value2
                );
            }
        }
        Err(_) => println!("compare failed: models differ structurally"),
    }

    if args.len() >= 6 {
        let names = model.workbook.get_worksheet_names();
        let sheet: u32 = match args[2].parse() {
            Ok(n) => n,
            Err(_) => names.iter().position(|n| n == &args[2]).unwrap() as u32,
        };
        let row: i32 = args[3].parse().unwrap();
        let col: i32 = args[4].parse().unwrap();
        let value = &args[5];
        for i in 1..=3 {
            let t = Instant::now();
            model
                .set_user_input(sheet, row, col, value.to_string())
                .unwrap();
            model.evaluate_edited(&[(sheet, row, col)]);
            println!("incremental edit+recalc #{i}: {} ms", t.elapsed().as_millis());
        }
        // Verify BEFORE any further full recalc on `model`: a fresh model with
        // the same edit and a full evaluation must agree with the
        // incrementally-updated model on every cell.
        let mut full = load_from_xlsx(file, "en", "UTC", "en").unwrap();
        full.evaluate();
        full.set_user_input(sheet, row, col, value.to_string()).unwrap();
        full.evaluate();
        match compare(&full, &model) {
            Ok(diffs) => {
                println!("incremental-vs-full verification: {} diffs", diffs.len());
                for d in diffs.iter().take(10) {
                    println!(
                        "VDIFF\t{}\t{}\t{}\t{}\t{:?}\t{:?}",
                        d.sheet_name, d.row, d.column, d.reason, d.value1, d.value2
                    );
                }
            }
            Err(_) => println!("incremental-vs-full verification: structural compare failure"),
        }
        {
            let t = Instant::now();
            model
                .set_user_input(sheet, row, col, value.to_string())
                .unwrap();
            model.evaluate();
            println!("full edit+recalc (for comparison): {} ms", t.elapsed().as_millis());
        }
    }
}
