//! Save a scalar formula that needs array evaluation, reopen, and confirm the
//! engine evaluates it the same way after the round trip through the file.
fn main() {
    let path = std::env::args().nth(1).expect("file.xlsx");
    let base = std::fs::read(&path).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&base).unwrap();
    wb.evaluate();
    let names = wb.sheet_names();
    let sheet = names[0].clone();
    // Row-finder idiom the agent uses: MATCH(1, (rng="x")*(rng2<>"y"), 0).
    let f = "=IFERROR(MATCH(1,(A1:A200<>\"\")*(B1:B200<>\"\"),0),\"\")";
    wb.set_input(0, 300, 1, f).unwrap();
    wb.set_input(0, 300, 2, "=IF(A300=\"\",\"\",INDEX(A1:A200,A300))").unwrap();
    wb.evaluate_edited().unwrap();
    let live_a: serde_json::Value = serde_json::from_str(&wb.cell(0, 300, 1).unwrap()).unwrap();
    let live_b: serde_json::Value = serde_json::from_str(&wb.cell(0, 300, 2).unwrap()).unwrap();
    eprintln!("live: A300={} B300={}", live_a["value"], live_b["value"]);
    let patch = serde_json::json!({
        "version": 1,
        "sheets": [{ "name": sheet, "cells": [
            { "r": 299, "c": 0, "f": f.trim_start_matches('='), "v": { "t": "n", "n": live_a["value"].as_f64().unwrap_or(0.0) } },
            { "r": 299, "c": 1, "f": "IF(A300=\"\",\"\",INDEX(A1:A200,A300))" }
        ]}]
    });
    let out = gridpath_engine::patch(&base, &patch.to_string()).unwrap();
    eprintln!("patched: {} → {} bytes", base.len(), out.len());
    // What landed in the XML.
    // The zip stores sheet XML deflated, so inspect through the engine's own
    // import below instead of the raw bytes; metadata presence by name.
    let names_blob = String::from_utf8_lossy(&out);
    eprintln!("metadata part named in package: {}", names_blob.contains("xl/metadata.xml"));
    // Reopen and evaluate.
    let mut wb2 = gridpath_engine::Workbook::open(&out).unwrap();
    wb2.evaluate();
    let a: serde_json::Value = serde_json::from_str(&wb2.cell(0, 300, 1).unwrap()).unwrap();
    let b: serde_json::Value = serde_json::from_str(&wb2.cell(0, 300, 2).unwrap()).unwrap();
    eprintln!("reopened: A300={} ({}) B300={} ({})", a["value"], a["content"], b["value"], b["content"]);
    assert_eq!(a["value"], live_a["value"], "A300 differs after round trip");
    assert_eq!(b["value"], live_b["value"], "B300 differs after round trip");
    eprintln!("round trip OK");
}
