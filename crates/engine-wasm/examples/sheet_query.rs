//! Run a query_rows JSON against a file: `sheet_query file.xlsx '<query json with "sheet": "Name">'`
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap(); let q = a.next().unwrap();
    let bytes = std::fs::read(&path).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    wb.evaluate();
    let mut v: serde_json::Value = serde_json::from_str(&q).unwrap();
    if let Some(name) = v["sheet"].as_str() {
        let idx = wb.sheet_names().iter().position(|n| n == name).expect("sheet") as u32;
        v["sheet"] = serde_json::json!(idx);
    }
    let out = wb.query_rows(&v.to_string()).unwrap();
    let out: serde_json::Value = serde_json::from_str(&out).unwrap();
    for r in out["rows"].as_array().unwrap_or(&vec![]) {
        let cells = r["cells"].as_object().unwrap();
        let mut parts: Vec<String> = Vec::new();
        for (k, val) in cells { parts.push(format!("{k}={}", match val { serde_json::Value::Null => "·".into(), serde_json::Value::Number(n) => format!("{:.4}", n.as_f64().unwrap()), other => other.to_string() })); }
        println!("{:>4}: {}", r["row"], parts.join("  "));
    }
    println!("matched {} truncated {}", out["matched"], out["truncated"]);
}
