//! Incremental recalc against a full one: `incremental_check file.xlsx [trials] [seed]`.
//! For random numeric inputs: set it, `evaluate_edited`, snapshot; then a full
//! `evaluate`, snapshot; every cell must agree, and every cell whose value
//! moved must be among the cells `evaluate_edited` reported.
use std::collections::{HashMap, HashSet};
use std::time::Instant;

type Key = (u64, i64, i64);

fn values(wb: &gridpath_engine::Workbook) -> (HashMap<Key, serde_json::Value>, HashSet<Key>) {
    let snap: serde_json::Value = serde_json::from_str(&wb.snapshot().unwrap()).unwrap();
    let mut out = HashMap::new();
    let mut volatile = HashSet::new();
    for (si, sh) in snap["sheets"].as_array().unwrap().iter().enumerate() {
        for c in sh["cells"].as_array().unwrap() {
            let k = (si as u64, c["r"].as_i64().unwrap(), c["c"].as_i64().unwrap());
            if let Some(f) = c["f"].as_str() {
                let f = f.to_uppercase();
                if ["RAND", "NOW(", "TODAY(", "RANDBETWEEN"].iter().any(|v| f.contains(v)) {
                    volatile.insert(k);
                }
            }
            out.insert(k, c["v"].clone());
        }
    }
    (out, volatile)
}

fn same(a: Option<&serde_json::Value>, b: Option<&serde_json::Value>) -> bool {
    match (a, b) {
        (Some(x), Some(y)) => match (x.as_f64(), y.as_f64()) {
            (Some(p), Some(q)) => p == q || (p - q).abs() <= 1e-9 * p.abs().max(q.abs()).max(1.0),
            _ => x == y,
        },
        (None, None) => true,
        (Some(v), None) | (None, Some(v)) => v.is_null() || v.as_str() == Some(""),
    }
}

fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().expect("file");
    let trials: usize = a.next().map(|s| s.parse().unwrap()).unwrap_or(5);
    let mut seed: u64 = a.next().map(|s| s.parse().unwrap()).unwrap_or(7);
    let bytes = std::fs::read(&path).unwrap();
    let mut wb = gridpath_engine::Workbook::open(&bytes).unwrap();
    let (mut before, volatile) = values(&wb);
    // Numeric constants are the inputs to perturb.
    let snap: serde_json::Value = serde_json::from_str(&wb.snapshot().unwrap()).unwrap();
    let mut inputs: Vec<(u32, i32, i32, f64)> = Vec::new();
    for (si, sh) in snap["sheets"].as_array().unwrap().iter().enumerate() {
        for c in sh["cells"].as_array().unwrap() {
            if c.get("f").is_none() {
                if let Some(v) = c["v"].as_f64() {
                    inputs.push((si as u32, c["r"].as_i64().unwrap() as i32, c["c"].as_i64().unwrap() as i32, v));
                }
            }
        }
    }
    drop(snap);
    let mut failures = 0;
    for t in 0..trials {
        seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
        let (s, r, c, v) = inputs[(seed >> 33) as usize % inputs.len()];
        let nv = if v == 0.0 { 1.0 } else { v * 1.1 + 1.0 };
        wb.set_input(s, r, c, &nv.to_string()).unwrap();
        let ti = Instant::now();
        let out: serde_json::Value = serde_json::from_str(&wb.evaluate_edited().unwrap()).unwrap();
        let inc_ms = ti.elapsed().as_millis();
        let reported: HashSet<Key> = out["cells"]
            .as_array()
            .map(|cs| cs.iter().map(|x| (x["sheet"].as_u64().unwrap(), x["r"].as_i64().unwrap(), x["c"].as_i64().unwrap())).collect())
            .unwrap_or_default();
        let (inc, _) = values(&wb);
        // Cells that moved but were not reported (the UI would show stale values).
        let unreported: Vec<&Key> = inc.keys().chain(before.keys()).filter(|k| !volatile.contains(k) && !same(inc.get(k), before.get(k)) && !reported.contains(k)).collect::<HashSet<_>>().into_iter().collect();
        let tf = Instant::now();
        wb.evaluate();
        let full_ms = tf.elapsed().as_millis();
        let (full, _) = values(&wb);
        let wrong: Vec<&Key> = full.keys().chain(inc.keys()).filter(|k| !volatile.contains(k) && !same(full.get(k), inc.get(k))).collect::<HashSet<_>>().into_iter().collect();
        let moved = inc.keys().filter(|k| !same(inc.get(k), before.get(k))).count();
        println!(
            "trial {t}: input ({s},{r},{c}) {v} -> {nv}: incremental {inc_ms} ms, full {full_ms} ms, {moved} moved, {} reported, {} wrong, {} unreported",
            reported.len(), wrong.len(), unreported.len()
        );
        for k in wrong.iter().take(5) {
            println!("   wrong {:?}: incremental {:?} full {:?}", k, inc.get(k), full.get(k));
        }
        for k in unreported.iter().take(5) {
            println!("   unreported {:?}: before {:?} after {:?}", k, before.get(k), inc.get(k));
        }
        if !wrong.is_empty() || !unreported.is_empty() {
            failures += 1;
        }
        // Put the input back and resync to the full state for the next trial.
        wb.set_input(s, r, c, &v.to_string()).unwrap();
        wb.evaluate_edited().unwrap();
        wb.evaluate();
        before = values(&wb).0;
    }
    println!("{} of {trials} trials failed", failures);
    if failures > 0 {
        std::process::exit(1);
    }
}
