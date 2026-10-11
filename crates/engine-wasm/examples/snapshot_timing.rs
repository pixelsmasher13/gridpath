//! Native timing harness for the engine's open/evaluate/snapshot path:
//! `cargo run --release --example snapshot_timing -- file.xlsx`
fn main() {
    let path = std::env::args().nth(1).expect("usage: snapshot_timing <file.xlsx>");
    let bytes = std::fs::read(&path).expect("read");
    let t = std::time::Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).expect("open");
    eprintln!("open: {:?}", t.elapsed());
    let t = std::time::Instant::now();
    wb.evaluate();
    eprintln!("evaluate: {:?}", t.elapsed());
    let t = std::time::Instant::now();
    let s = wb.snapshot().expect("snapshot");
    eprintln!("snapshot: {:?} ({} MB)", t.elapsed(), s.len() / 1_000_000);
}
