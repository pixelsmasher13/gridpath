//! Reproduces the desktop's engine sequence for an agent-built sheet:
//! read the sheet's cells from the saved file, delete the sheet, re-create
//! it (full evaluate, as the sheet op does), then write the cells back in
//! chunks of N with an incremental evaluate after each chunk.
//! `cargo run --release --example chunk_replay -- file.xlsx "Sheet" 1000`
use std::time::Instant;
fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().unwrap();
    let sheet = args.next().unwrap();
    let chunk: usize = args.next().and_then(|s| s.parse().ok()).unwrap_or(1000);
    let bytes = std::fs::read(&path).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    wb.evaluate();
    let names = wb.sheet_names();
    let idx = names.iter().position(|n| *n == sheet).expect("sheet") as u32;
    // Capture inputs: formula text or literal, in row-major order.
    let cells: serde_json::Value = serde_json::from_str(&wb.sheet_cells(idx).unwrap()).unwrap();
    let mut inputs: Vec<(i32, i32, String)> = Vec::new();
    for c in cells["cells"].as_array().unwrap() {
        let r = c["r"].as_i64().unwrap() as i32;
        let col = c["c"].as_i64().unwrap() as i32;
        let input = if let Some(f) = c["f"].as_str() {
            f.to_string()
        } else {
            match &c["v"] {
                serde_json::Value::Null => continue,
                serde_json::Value::String(s) => s.clone(),
                serde_json::Value::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
                other => other.to_string(),
            }
        };
        inputs.push((r, col, input));
    }
    eprintln!("{sheet}: {} input cells captured", inputs.len());
    wb.delete_sheet(idx).unwrap();
    let t = Instant::now();
    wb.evaluate();
    eprintln!("after delete: full evaluate {:?}", t.elapsed());
    let t = Instant::now();
    let new_idx = wb.new_sheet().unwrap();
    wb.rename_sheet(new_idx, &sheet).unwrap();
    wb.evaluate();
    eprintln!("create_sheet + full evaluate: {:?}", t.elapsed());
    let total = Instant::now();
    for (i, ch) in inputs.chunks(chunk).enumerate() {
        let t = Instant::now();
        for (r, c, v) in ch {
            wb.set_input(new_idx, *r, *c, v).unwrap();
        }
        let t_set = t.elapsed();
        let t = Instant::now();
        let out = wb.evaluate_edited().unwrap();
        let v: serde_json::Value = serde_json::from_str(&out).unwrap();
        eprintln!(
            "chunk {i}: {} cells, set_input {:?}, evaluate_edited {:?}, {} changed",
            ch.len(),
            t_set,
            t.elapsed(),
            v["cells"].as_array().map(|a| a.len()).unwrap_or(0)
        );
    }
    eprintln!("all chunks: {:?}", total.elapsed());
    let probe: serde_json::Value = serde_json::from_str(&wb.cell(new_idx, 16, 11).unwrap()).unwrap();
    eprintln!("K16 after replay: {}", probe);
}
