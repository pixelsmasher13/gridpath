//! Cost of toggling one input between two values: `edit_toggle file.xlsx "Sheet" ROW COL a b [repeat]`
//! (1-based row/col). Times set_input + evaluate_edited per edit; IC_DEBUG=1 adds phase timings.
use std::time::Instant;
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap();
    let sheet = a.next().unwrap();
    let row: i32 = a.next().unwrap().parse().unwrap();
    let col: i32 = a.next().unwrap().parse().unwrap();
    let (va, vb) = (a.next().unwrap(), a.next().unwrap());
    let repeat: usize = a.next().map(|s| s.parse().unwrap()).unwrap_or(4);
    let bytes = std::fs::read(&path).unwrap();
    let t = Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    eprintln!("open+evaluate {:?}", t.elapsed());
    let idx = wb.sheet_names().iter().position(|n| *n == sheet).unwrap() as u32;
    for i in 0..repeat {
        let v = if i % 2 == 0 { &va } else { &vb };
        let t = Instant::now();
        wb.set_input(idx, row, col, v).unwrap();
        let out = wb.evaluate_edited().unwrap();
        let n = serde_json::from_str::<serde_json::Value>(&out).unwrap()["cells"].as_array().map(|a| a.len()).unwrap_or(0);
        eprintln!("edit {i}: {:>9.3?}  ({n} cells returned)", t.elapsed());
    }
}
