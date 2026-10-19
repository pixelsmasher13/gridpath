//! Cost of editing a specific cell with its dependents: `edit_at file.xlsx "Sheet" ROW COL value [repeat]`
//! (1-based row/col). Times set_input + evaluate_edited and reports cells changed.
use std::time::Instant;
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap(); let sheet = a.next().unwrap();
    let row: i32 = a.next().unwrap().parse().unwrap(); let col: i32 = a.next().unwrap().parse().unwrap();
    let value = a.next().unwrap(); let repeat: usize = a.next().map(|s| s.parse().unwrap()).unwrap_or(3);
    let bytes = std::fs::read(&path).unwrap();
    let t = Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    eprintln!("open+evaluate {:?}", t.elapsed());
    let idx = wb.sheet_names().iter().position(|n| *n == sheet).unwrap() as u32;
    let original = wb.cell(idx, row, col).unwrap();
    eprintln!("cell before: {}", original.chars().take(120).collect::<String>());
    for i in 0..repeat {
        let v = if i % 2 == 0 { value.clone() } else { format!("{value}x") };
        let t = Instant::now();
        wb.set_input(idx, row, col, &v).unwrap();
        let t_set = t.elapsed();
        let t = Instant::now();
        let out = wb.evaluate_edited().unwrap();
        let n = serde_json::from_str::<serde_json::Value>(&out).unwrap()["cells"].as_array().map(|a| a.len()).unwrap_or(0);
        eprintln!("edit {i}: set_input {t_set:>9.3?}  evaluate_edited {:>9.3?}  ({n} cells changed)", t.elapsed());
    }
}
