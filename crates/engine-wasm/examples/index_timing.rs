//! Native timing harness for the structural index and a row query:
//! `cargo run --release --example index_timing -- file.xlsx [sheet] [column] [value]`
fn main() {
    let mut args = std::env::args().skip(1);
    let path = args.next().expect("usage: index_timing <file.xlsx> [sheet] [column] [value]");
    let sheet: u32 = args.next().and_then(|s| s.parse().ok()).unwrap_or(0);
    let column = args.next().unwrap_or_else(|| "A".into());
    let value = args.next().unwrap_or_default();
    let bytes = std::fs::read(&path).expect("read");
    let t = std::time::Instant::now();
    let mut wb = gridpath_engine::Workbook::open(&bytes).expect("open");
    wb.evaluate();
    eprintln!("open+evaluate: {:?}", t.elapsed());
    let t = std::time::Instant::now();
    let idx = wb.workbook_index().expect("index");
    eprintln!("workbook index: {:?} ({} KB)", t.elapsed(), idx.len() / 1000);
    let v: serde_json::Value = serde_json::from_str(&idx).unwrap();
    for s in v["sheets"].as_array().unwrap() {
        eprintln!(
            "  {:<40} used={:<12} header_row={:<6} rows={:<7} sections={:<5} runs={}{}",
            s["name"].as_str().unwrap_or(""),
            s["usedRange"].as_str().unwrap_or("-"),
            s["headerRow"].to_string(),
            s["rows"].as_array().map(|a| a.len()).unwrap_or(0),
            s["sections"].as_array().map(|a| a.len()).unwrap_or(0),
            s["formulaRuns"].as_array().map(|a| a.len()).unwrap_or(0),
            if s["formulaRunsTruncated"] == true { "+" } else { "" }
        );
    }
    let q = serde_json::json!({
        "sheet": sheet,
        "where": if value.is_empty() { serde_json::json!([]) } else { serde_json::json!([{ "column": column, "op": "eq", "value": value }]) },
        "limit": 5
    });
    let t = std::time::Instant::now();
    let out = wb.query_rows(&q.to_string()).expect("query");
    eprintln!("query_rows: {:?}", t.elapsed());
    let out: serde_json::Value = serde_json::from_str(&out).unwrap();
    eprintln!("{}", serde_json::to_string_pretty(&out).unwrap().chars().take(2500).collect::<String>());
    let t = std::time::Instant::now();
    let g = serde_json::json!({ "sheet": sheet, "group_by": column, "limit": 8 });
    let out = wb.query_rows(&g.to_string()).expect("group");
    eprintln!("group_by {column}: {:?}", t.elapsed());
    let out: serde_json::Value = serde_json::from_str(&out).unwrap();
    eprintln!("{}", serde_json::to_string_pretty(&out).unwrap().chars().take(1500).collect::<String>());
}
