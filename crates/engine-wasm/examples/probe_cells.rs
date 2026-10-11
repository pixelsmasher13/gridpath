//! `cargo run --release --example probe_cells -- file.xlsx "Sheet Name" A1:D20 ["Sheet2" A1:B3 ...]`
fn col_idx(s: &str) -> i32 { s.bytes().fold(0, |n, b| n * 26 + (b - b'A' + 1) as i32) }
fn parse(a1: &str) -> (i32, i32) { let i = a1.find(|c: char| c.is_ascii_digit()).unwrap(); (a1[i..].parse().unwrap(), col_idx(&a1[..i])) }
fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().unwrap();
    let bytes = std::fs::read(&path).unwrap();
    let t = std::time::Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    wb.evaluate();
    eprintln!("open+evaluate {:?}", t.elapsed());
    let names = wb.sheet_names();
    while let (Some(sheet), Some(range)) = (args.next(), args.next()) {
        let idx = names.iter().position(|n| *n == sheet).expect("sheet") as u32;
        let (a, b) = range.split_once(':').unwrap();
        let (r1, c1) = parse(a); let (r2, c2) = parse(b);
        eprintln!("== {sheet} {range}");
        for r in r1..=r2 {
            let mut line = format!("{r:>4}:");
            for c in c1..=c2 {
                let v: serde_json::Value = serde_json::from_str(&wb.cell(idx, r, c).unwrap()).unwrap();
                let content = v["content"].as_str().unwrap_or("");
                let val = &v["value"];
                let cell = if content.starts_with('=') { format!(" {}={}", &content[..content.len().min(28)], val) } else { format!(" {}", val) };
                line.push_str(&cell);
            }
            eprintln!("{}", line.chars().take(260).collect::<String>());
        }
    }
}
