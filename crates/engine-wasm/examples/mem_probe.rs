//! Resident size of the engine process after each open-path stage:
//! `mem_probe file.xlsx "Sheet for a test edit"`
use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};
/// Counting allocator: exact live heap bytes, unaffected by pages the
/// system allocator keeps after a free (which made RSS deltas meaningless).
struct Counting;
static LIVE: AtomicUsize = AtomicUsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 { LIVE.fetch_add(l.size(), Ordering::Relaxed); System.alloc(l) }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) { LIVE.fetch_sub(l.size(), Ordering::Relaxed); System.dealloc(p, l) }
    unsafe fn realloc(&self, p: *mut u8, l: Layout, n: usize) -> *mut u8 { LIVE.fetch_add(n, Ordering::Relaxed); LIVE.fetch_sub(l.size(), Ordering::Relaxed); System.realloc(p, l, n) }
}
#[global_allocator]
static A: Counting = Counting;
fn rss_mb() -> f64 { LIVE.load(Ordering::Relaxed) as f64 / 1e6 }
fn main() {
    let mut a = std::env::args().skip(1);
    let path = a.next().unwrap(); let sheet = a.next().unwrap_or_default();
    let base = rss_mb();
    let bytes = std::fs::read(&path).unwrap();
    let wb = ironcalc::import::load_from_xlsx_bytes(&bytes, "workbook", "en", "UTC").unwrap();
    let mut model = ironcalc_base::Model::from_workbook(wb, "en").unwrap();
    drop(bytes);
    println!("{:>8.0} MB live  model loaded (parsed, unevaluated)", rss_mb() - base);
    let cached = gridpath_engine_core::build_cached(&model);
    println!("{:>8.0} MB  + file-cache map ({} entries)", rss_mb() - base, cached.len());
    model.evaluate();
    println!("{:>8.0} MB  + first full evaluate (cells + support maps)", rss_mb() - base);
    let json = gridpath_engine_core::snapshot_json(&model, &cached).unwrap();
    println!("{:>8.0} MB  + snapshot string ({} MB)", rss_mb() - base, json.len() / 1_000_000);
    drop(json);
    println!("{:>8.0} MB  snapshot string dropped", rss_mb() - base);
    let mut cache = gridpath_engine_core::index::IndexCache::default();
    let (idx, _) = cache.workbook_index_json(&model).unwrap();
    println!("{:>8.0} MB  + structural index cache ({} MB json)", rss_mb() - base, idx.len() / 1_000_000);
    // One incremental edit: builds the cached reader index.
    let names = model.workbook.get_worksheet_names();
    let si = names.iter().position(|n| *n == sheet).unwrap_or(0) as u32;
    model.set_user_input(si, 2000, 100, "=1+1".to_string()).unwrap();
    let _ = model.evaluate_edited(&[(si, 2000, 100)]);
    println!("{:>8.0} MB  + reader index (after first incremental edit)", rss_mb() - base);
    // New settle: compare live vs the file cache directly (no map), then prune the cache.
    let mut cached = cached;
    let changed = gridpath_engine_core::diff::changed_vs_cached(&model, &cached);
    println!("{:>8.0} MB  settle diff vs cache, no map ({} changed)", rss_mb() - base, changed.len());
    let pruned = gridpath_engine_core::diff::prune_cached(&model, &mut cached);
    println!("{:>8.0} MB  file-cache map pruned by {} to {} entries", rss_mb() - base, pruned, cached.len());
    // New sheet-op / structural diff: one compact before-map, streaming after side.
    let before = gridpath_engine_core::diff::formula_values(&model);
    println!("{:>8.0} MB  + compact before-map ({} entries)", rss_mb() - base, before.len());
    let changed = gridpath_engine_core::diff::changed_since(before, &model, Some);
    println!("{:>8.0} MB  streaming diff done ({} changed), before-map consumed", rss_mb() - base, changed.len());
    // For comparison, the old serde_json::Value map.
    let mut m1 = std::collections::HashMap::new();
    for (s, ws) in model.workbook.worksheets.iter().enumerate() { for (r, row) in &ws.sheet_data { for (c, cell) in row { if cell.get_formula().is_some() { m1.insert((s as u32, *r, *c), gridpath_engine_core::cell_value_json(model.get_cell_value_by_index(s as u32, *r, *c).unwrap())); } } } }
    println!("{:>8.0} MB  (old) one serde_json value map ({} entries)", rss_mb() - base, m1.len());
    drop(m1);
}
