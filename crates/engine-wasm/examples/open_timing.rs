//! Native open-path stages: `cargo run --release --example open_timing -- file.xlsx [out.json]`
//! read → xlsx parse → model build → first evaluate → snapshot JSON (optionally written out).
use std::time::Instant;
fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().expect("file");
    let out = args.next();
    let t = Instant::now();
    let bytes = std::fs::read(&path).unwrap();
    eprintln!("read file: {:?} ({} MB)", t.elapsed(), bytes.len() / 1_000_000);
    let t = Instant::now();
    let wb = ironcalc::import::load_from_xlsx_bytes(&bytes, "workbook", "en", "UTC").unwrap();
    eprintln!("xlsx parse: {:?}", t.elapsed());
    let t = Instant::now();
    let mut model = ironcalc_base::Model::from_workbook(wb, "en").unwrap();
    eprintln!("model build (parse formulas): {:?}", t.elapsed());
    let t = Instant::now();
    let cached = gridpath_engine_core::build_cached(&model);
    eprintln!("cached-value map: {:?} ({} entries)", t.elapsed(), cached.len());
    let t = Instant::now();
    model.evaluate();
    eprintln!("first evaluate: {:?}", t.elapsed());
    let t = Instant::now();
    let json = gridpath_engine_core::snapshot_json(&model, &cached).unwrap();
    eprintln!("snapshot json: {:?} ({} MB)", t.elapsed(), json.len() / 1_000_000);
    let cells: usize = model.workbook.worksheets.iter().map(|w| w.sheet_data.values().map(|r| r.len()).sum::<usize>()).sum();
    let formulas: usize = model.workbook.worksheets.iter().map(|w| w.sheet_data.values().flat_map(|r| r.values()).filter(|c| c.get_formula().is_some()).count()).sum();
    eprintln!("cells {cells}, formulas {formulas}, sheets {}", model.workbook.worksheets.len());
    if let Some(o) = out { std::fs::write(o, json).unwrap(); }
}
