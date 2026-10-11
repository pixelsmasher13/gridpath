//! Replays the agent batches of an unsaved-work journal against the engine
//! natively and times each engine step the desktop app would have taken:
//! `cargo run --release --example replay_timing -- journal.json [file.xlsx]`
//!
//! For every batch: create_sheet (a full evaluate, as `calc_engine_sheet_op`
//! does), then all set_cell writes as one incremental `evaluate_edited`.
use serde_json::Value;
use std::time::Instant;

fn main() {
    let mut args = std::env::args().skip(1);
    let journal_path = args.next().expect("usage: replay_timing <journal.json> [file.xlsx]");
    let journal: Value = serde_json::from_str(&std::fs::read_to_string(&journal_path).expect("read journal")).unwrap();
    let path = args.next().unwrap_or_else(|| journal["path"].as_str().unwrap().to_string());
    let bytes = std::fs::read(&path).expect("read xlsx");
    let t = Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).expect("open");
    eprintln!("open (incl. first evaluate): {:?}", t.elapsed());
    let t = Instant::now();
    wb.evaluate();
    eprintln!("baseline full evaluate: {:?}", t.elapsed());

    for b in journal["batches"].as_array().unwrap() {
        let id = b["id"].as_str().unwrap_or("?");
        let muts = b["mutations"].as_array().unwrap();
        let mut sheet_names = wb.sheet_names();
        let mut writes = 0usize;
        let mut formulas = 0usize;
        let t_batch = Instant::now();
        let mut pending_edit = false;
        for m in muts {
            match m["type"].as_str().unwrap_or("") {
                "create_sheet" => {
                    let name = m["name"].as_str().unwrap();
                    let t = Instant::now();
                    let idx = wb.new_sheet().expect("new_sheet");
                    wb.rename_sheet(idx, name).expect("rename");
                    // The desktop's sheet op runs a FULL evaluate.
                    wb.evaluate();
                    sheet_names = wb.sheet_names();
                    eprintln!("[{id}] create_sheet {name:?} + full evaluate: {:?}", t.elapsed());
                }
                "set_cell" => {
                    let a = &m["address"];
                    let sheet = a["sheet"].as_str().unwrap();
                    let idx = sheet_names.iter().position(|s| s == sheet).expect("sheet") as u32;
                    let row = a["row"].as_i64().unwrap() as i32;
                    let col = a["col"].as_i64().unwrap() as i32;
                    let input = if let Some(f) = m["new_formula"].as_str() {
                        formulas += 1;
                        if f.starts_with('=') { f.to_string() } else { format!("={f}") }
                    } else {
                        match &m["new_value"] {
                            Value::Null => String::new(),
                            Value::String(s) => s.clone(),
                            Value::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
                            other => other.to_string(),
                        }
                    };
                    wb.set_input(idx, row, col, &input).expect("set_input");
                    writes += 1;
                    pending_edit = true;
                }
                _ => {}
            }
        }
        if pending_edit {
            let t = Instant::now();
            let out = wb.evaluate_edited().expect("evaluate_edited");
            let v: Value = serde_json::from_str(&out).unwrap();
            let n = v["cells"].as_array().map(|a| a.len()).unwrap_or(0);
            eprintln!(
                "[{id}] {writes} writes ({formulas} formulas) → evaluate_edited: {:?}, {n} cells changed (full={})",
                t.elapsed(),
                v["full"]
            );
        }
        eprintln!("[{id}] batch total: {:?}", t_batch.elapsed());
        // What the NEXT create_sheet would cost now.
        let t = Instant::now();
        wb.evaluate();
        eprintln!("[{id}] full evaluate with this sheet present: {:?}", t.elapsed());
    }
}
