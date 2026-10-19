//! Re-save every formula on the given sheets through the patcher so they
//! become dynamic-array formulas (Excel evaluates arrays like the engine):
//! `da_convert in.xlsx out.xlsx "Sheet A" "Sheet B" ...`
use serde_json::json;
fn main() {
    let mut a = std::env::args().skip(1);
    let inp = a.next().expect("in.xlsx"); let outp = a.next().expect("out.xlsx");
    let sheets: Vec<String> = a.collect();
    let base = std::fs::read(&inp).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&base).unwrap();
    wb.evaluate();
    let names = wb.sheet_names();
    let mut patch_sheets = Vec::new();
    let mut total = 0usize;
    for name in &sheets {
        let idx = names.iter().position(|n| n == name).unwrap_or_else(|| panic!("no sheet {name}")) as u32;
        let cells: serde_json::Value = serde_json::from_str(&wb.sheet_cells(idx).unwrap()).unwrap();
        let mut out = Vec::new();
        for c in cells["cells"].as_array().unwrap() {
            let Some(f) = c["f"].as_str() else { continue };
            let r = c["r"].as_i64().unwrap() as u32; let col = c["c"].as_i64().unwrap() as u32;
            let mut cell = json!({ "r": r - 1, "c": col - 1, "f": f.trim_start_matches('=') });
            // cached value, typed as the patcher expects
            match &c["v"] {
                serde_json::Value::Number(n) => cell["v"] = json!({ "t": "n", "n": n.as_f64().unwrap() }),
                serde_json::Value::String(s) if s.starts_with('#') => cell["v"] = json!({ "t": "e", "e": s }),
                serde_json::Value::String(s) => cell["v"] = json!({ "t": "s", "s": s }),
                serde_json::Value::Bool(b) => cell["v"] = json!({ "t": "b", "b": b }),
                _ => {}
            }
            out.push(cell);
        }
        total += out.len();
        eprintln!("{name}: {} formula cells", out.len());
        patch_sheets.push(json!({ "name": name, "cells": out }));
    }
    let patch = json!({ "version": 1, "sheets": patch_sheets });
    let t = std::time::Instant::now();
    let bytes = gridpath_engine::patch(&base, &patch.to_string()).unwrap();
    eprintln!("patched {total} formula cells in {:?}; {} → {} bytes", t.elapsed(), base.len(), bytes.len());
    std::fs::write(&outp, &bytes).unwrap();
    // Verify: reopen, evaluate, compare every converted cell's value.
    let mut wb2 = gridpath_engine::Workbook::open(&bytes).unwrap();
    wb2.evaluate();
    let mut diffs = 0usize; let mut checked = 0usize;
    for name in &sheets {
        let i1 = names.iter().position(|n| n == name).unwrap() as u32;
        let i2 = wb2.sheet_names().iter().position(|n| n == name).unwrap() as u32;
        let a: serde_json::Value = serde_json::from_str(&wb.sheet_cells(i1).unwrap()).unwrap();
        let b: serde_json::Value = serde_json::from_str(&wb2.sheet_cells(i2).unwrap()).unwrap();
        let bmap: std::collections::HashMap<(i64, i64), &serde_json::Value> = b["cells"].as_array().unwrap().iter().map(|c| ((c["r"].as_i64().unwrap(), c["c"].as_i64().unwrap()), c)).collect();
        for c in a["cells"].as_array().unwrap() {
            if c["f"].is_null() { continue; }
            checked += 1;
            let k = (c["r"].as_i64().unwrap(), c["c"].as_i64().unwrap());
            let same = bmap.get(&k).map_or(false, |d| d["v"] == c["v"] || (c["v"].is_null() && d["v"].as_str() == Some("")) || (d["v"].is_null() && c["v"].as_str() == Some("")));
            if !same { diffs += 1; if diffs <= 5 { eprintln!("  diff {name}!R{}C{}: {} → {}", k.0, k.1, c["v"], bmap.get(&k).map(|d| d["v"].to_string()).unwrap_or("missing".into())); } }
        }
    }
    eprintln!("verify: {checked} formula cells, {diffs} value differences after round trip");
}
