//! Per-call cost of a one-cell incremental edit on a big model:
//! `edit_cost file.xlsx "Sheet"` — writes into an empty corner of the sheet,
//! timing set_input and evaluate_edited for each write.
use std::time::Instant;
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap(); let sheet = a.next().unwrap();
    let bytes = std::fs::read(&path).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    wb.evaluate();
    let idx = wb.sheet_names().iter().position(|n| *n == sheet).unwrap() as u32;
    let inputs = ["hello", "42", "=1+1", "=SUM(K11:K15)", "=SUMIFS(Sales!K:K,Sales!$CV:$CV,\"Acme\")", "7", "=AX400+1"];
    for (i, input) in inputs.iter().enumerate() {
        let r = 400 + i as i32;
        let t = Instant::now();
        wb.set_input(idx, r, 50, input).unwrap();
        let t_set = t.elapsed();
        let t = Instant::now();
        let out = wb.evaluate_edited().unwrap();
        let n = serde_json::from_str::<serde_json::Value>(&out).unwrap()["cells"].as_array().map(|a| a.len()).unwrap_or(0);
        eprintln!("{:<75} set_input {:>9.3?}  evaluate_edited {:>9.3?}  ({n} cells)", input.chars().take(75).collect::<String>(), t_set, t.elapsed());
    }
}
