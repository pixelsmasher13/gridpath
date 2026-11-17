//! Workspace: several workbooks loaded at once, linked by formulas the engine
//! evaluates.
//!
//! A formula in one workbook can read a sheet of another:
//! `='[NOW.xlsx]Model'!BV32`. Behind every sheet read that way the reading
//! workbook keeps a hidden copy of it, named exactly like the link
//! (`[NOW.xlsx]Model`, see `gridpath_engine_core::links`), and this module
//! keeps those copies current:
//!
//!   - the other workbook is found and loaded, with a tab or without one
//!     (a model without a tab lives in the same map under `bg:<path>`);
//!   - a value that changes in a workbook is carried to every workbook that
//!     reads it, and on to the readers of those;
//!   - rows and columns inserted or deleted, and sheets renamed, in a
//!     workbook are followed by the links that point there;
//!   - a save writes the links the way Excel stores them (`[1]Model!BV32`
//!     plus one external-link part per workbook read), and a file saved that
//!     way has its links attached again when it is opened.
//!
//! Two rules keep the locking simple. A file is never loaded while the model
//! map is locked, and the workspace lock is never taken while the model map
//! is locked. A model without a tab is always evaluated; only a tab's model
//! can still be waiting for its first evaluation.

use std::collections::{BTreeMap, HashMap, HashSet, VecDeque};
use std::panic::AssertUnwindSafe;
use std::path::{Component, Path, PathBuf};
use std::sync::{Arc, Mutex, MutexGuard};
use std::time::{Instant, SystemTime};

use gridpath_engine_core::diff::{changed_since, formula_values};
use gridpath_engine_core::links as core_links;
use ironcalc::base::new_empty::is_link_sheet_name;
use ironcalc::base::Model;
use ironcalc::import::load_from_xlsx;
use serde::Serialize;
use tauri::State;

use super::calc_engine::{
    deltas_for, lock_recover, next_seq, run_contained, settle_pending, CalcEditResult, CalcEngineState, CellDelta,
    EditCell, Loaded, Models,
};
use super::xlsx_patch::external_links;

type Key = (u32, i32, i32);

// ---------------------------------------------------------------------------
// Link text
// ---------------------------------------------------------------------------

/// One `[book]Sheet!` prefix in a formula. `start..end` are byte offsets and
/// cover the prefix with its `!`. `dir` is the folder written before the
/// bracket (quoted form only), else empty.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct FoundLink {
    pub start: usize,
    pub end: usize,
    pub dir: String,
    pub book: String,
    pub sheet: String,
}

fn name_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.'
}

/// Every link prefix in a formula, in order. String literals are skipped,
/// and so are the brackets of table references (`Sales[Amount]`,
/// `[@Amount]`): only `[book]Sheet!` counts. The same rules as
/// `packages/gridpath/src/workbook/links.ts`.
pub(super) fn find_links(formula: &str) -> Vec<FoundLink> {
    let chars: Vec<(usize, char)> = formula.char_indices().collect();
    let n = chars.len();
    let at = |i: usize| chars.get(i).map(|c| c.1);
    let byte = |i: usize| chars.get(i).map(|c| c.0).unwrap_or(formula.len());
    let mut out = Vec::new();
    let mut i = 0;
    while i < n {
        let ch = chars[i].1;
        if ch == '"' {
            i += 1;
            while i < n {
                if chars[i].1 == '"' {
                    if at(i + 1) == Some('"') {
                        i += 2;
                        continue;
                    }
                    break;
                }
                i += 1;
            }
            i += 1;
            continue;
        }
        if ch == '\'' {
            let mut j = i + 1;
            let mut name = String::new();
            let mut closed = false;
            while j < n {
                if chars[j].1 == '\'' {
                    if at(j + 1) == Some('\'') {
                        name.push('\'');
                        j += 2;
                        continue;
                    }
                    closed = true;
                    break;
                }
                name.push(chars[j].1);
                j += 1;
            }
            if !closed {
                break;
            }
            if at(j + 1) != Some('!') {
                i = j + 1;
                continue;
            }
            if let Some(close) = name.rfind(']') {
                if let Some(open) = name[..close].rfind('[') {
                    if close > open + 1 && close < name.len() - 1 {
                        out.push(FoundLink {
                            start: byte(i),
                            end: byte(j + 2),
                            dir: name[..open].to_string(),
                            book: name[open + 1..close].to_string(),
                            sheet: name[close + 1..].to_string(),
                        });
                    }
                }
            }
            i = j + 2;
            continue;
        }
        if ch == '[' {
            let prev = if i > 0 { Some(chars[i - 1].1) } else { None };
            let close = (i + 1..n).find(|&k| chars[k].1 == ']');
            if let Some(close) = close {
                // `Table[Column]` and `[[#Headers],[Column]]` are not links.
                if close > i + 1 && !prev.is_some_and(|p| name_char(p) || p == ']') {
                    let book: String = chars[i + 1..close].iter().map(|c| c.1).collect();
                    let first = book.chars().next();
                    if first != Some('@') && first != Some('#') && !book.contains(['[', '\'', '"', '!']) {
                        let mut k = close + 1;
                        while k < n && name_char(chars[k].1) {
                            k += 1;
                        }
                        if k > close + 1 && at(k) == Some('!') {
                            let sheet: String = chars[close + 1..k].iter().map(|c| c.1).collect();
                            out.push(FoundLink { start: byte(i), end: byte(k + 1), dir: String::new(), book, sheet });
                            i = k + 1;
                            continue;
                        }
                    }
                }
            }
        }
        i += 1;
    }
    out
}

/// The formula with each link prefix replaced by what `to` returns (`None`
/// keeps it as written). `to` sees the links in the order they are written.
pub(super) fn rewrite_links(formula: &str, mut to: impl FnMut(&FoundLink) -> Option<String>) -> String {
    let found = find_links(formula);
    let replacements: Vec<Option<String>> = found.iter().map(&mut to).collect();
    let mut out = formula.to_string();
    for (link, replacement) in found.iter().zip(replacements).rev() {
        if let Some(text) = replacement {
            out.replace_range(link.start..link.end, &text);
        }
    }
    out
}

/// `[NOW.xlsx]Model` → (`NOW.xlsx`, `Model`).
pub(super) fn split_link_sheet_name(name: &str) -> Option<(&str, &str)> {
    if !name.starts_with('[') {
        return None;
    }
    let close = name.find(']')?;
    if close < 2 || close == name.len() - 1 {
        return None;
    }
    Some((&name[1..close], &name[close + 1..]))
}

/// A sheet name as a formula prefix, always quoted: `'[NOW.xlsx]Model'!`.
pub(super) fn quoted_prefix(sheet_name: &str) -> String {
    format!("'{}'!", sheet_name.replace('\'', "''"))
}

/// A link as an .xlsx file holds it: the other workbook by its number in the
/// file, `[1]Model!`, quoted when the sheet name is more than letters,
/// digits, `_` and `.` (the rule the engine prints sheet names by).
pub(super) fn numbered_prefix(n: u32, sheet: &str) -> String {
    if !sheet.is_empty() && sheet.chars().all(name_char) {
        format!("[{n}]{sheet}!")
    } else {
        format!("'[{n}]{}'!", sheet.replace('\'', "''"))
    }
}

fn is_number(text: &str) -> bool {
    !text.is_empty() && text.bytes().all(|b| b.is_ascii_digit())
}

fn names_a_link(input: &str) -> bool {
    input.starts_with('=') && input.contains('[')
}

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

/// Paths are compared without case: the same file can be spelled both ways
/// on the file systems this runs on.
pub(super) fn same(a: &str, b: &str) -> bool {
    a.to_uppercase() == b.to_uppercase()
}

pub(super) fn basename(path: &str) -> String {
    path.rsplit(['/', '\\']).find(|s| !s.is_empty()).unwrap_or(path).to_string()
}

fn is_workbook_name(name: &str) -> bool {
    let lower = name.to_lowercase();
    lower.ends_with(".xlsx") || lower.ends_with(".xlsm")
}

fn is_absolute(path: &str) -> bool {
    !path.is_empty() && Path::new(path).is_absolute()
}

fn encode_segment(segment: &str) -> String {
    let mut out = String::new();
    for ch in segment.chars() {
        if ch.is_whitespace() || "%#?\"<>|\\^`{}".contains(ch) {
            let mut units = [0u16; 2];
            for unit in ch.encode_utf16(&mut units) {
                out.push_str(&format!("%{:02X}", unit));
            }
        } else {
            out.push(ch);
        }
    }
    out
}

/// How a saved file names the workbook `source`: relative when it sits in
/// `dir` (the folder the reader is saved into) or below it, else a `file:///`
/// address.
pub(super) fn link_target(source: &str, dir: &str) -> String {
    if !dir.is_empty() {
        if let Ok(rel) = Path::new(source).strip_prefix(Path::new(dir)) {
            let segments: Vec<String> = rel
                .components()
                .filter_map(|c| match c {
                    Component::Normal(part) => Some(encode_segment(&part.to_string_lossy())),
                    _ => None,
                })
                .collect();
            if !segments.is_empty() {
                return segments.join("/");
            }
        }
    }
    let segments: Vec<String> = source
        .split(['/', '\\'])
        .filter(|s| !s.is_empty())
        .enumerate()
        .map(|(i, s)| {
            let drive = i == 0 && s.len() == 2 && s.ends_with(':') && s.as_bytes()[0].is_ascii_alphabetic();
            if drive { s.to_string() } else { encode_segment(s) }
        })
        .collect();
    format!("file:///{}", segments.join("/"))
}

/// `%20` → space. `None` when the text is not a valid encoding.
fn percent_decode(text: &str) -> Option<String> {
    let bytes = text.as_bytes();
    let mut out: Vec<u8> = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            let hex = text.get(i + 1..i + 3)?;
            out.push(u8::from_str_radix(hex, 16).ok()?);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

/// `base` joined with `target` (or `target` alone when it is absolute), with
/// `.` and `..` folded away. Does not touch the disk.
fn resolve_lexical(base: &Path, target: &str) -> PathBuf {
    let joined = if Path::new(target).is_absolute() { PathBuf::from(target) } else { base.join(target) };
    let mut out = PathBuf::new();
    for part in joined.components() {
        match part {
            Component::CurDir => {}
            Component::ParentDir => {
                out.pop();
            }
            other => out.push(other.as_os_str()),
        }
    }
    out
}

/// The file called `name` in `dir`: spelled exactly so, or differing only in
/// case.
fn file_in_dir(dir: &Path, name: &str) -> Option<String> {
    let mut loose: Option<String> = None;
    for entry in std::fs::read_dir(dir).ok()?.flatten() {
        let found = entry.file_name().to_string_lossy().to_string();
        if found == name || (loose.is_none() && same(&found, name)) {
            let path = entry.path();
            if !path.is_file() {
                continue;
            }
            let text = path.to_string_lossy().to_string();
            if found == name {
                return Some(text);
            }
            loose = Some(text);
        }
    }
    loose
}

/// The workbook a typed book name means inside `dir`: `NOW.xlsx`, or `NOW`
/// for `NOW.xlsx` / `NOW.xlsm`.
fn book_in_dir(dir: &Path, book: &str) -> Option<String> {
    if is_workbook_name(book) {
        return file_in_dir(dir, book);
    }
    file_in_dir(dir, &format!("{book}.xlsx")).or_else(|| file_in_dir(dir, &format!("{book}.xlsm")))
}

/// The file a saved link names, seen from the workbook `reader` that holds
/// it: at the address as written, or a file of that name next to the reader
/// (a folder of models that was moved or copied as a whole).
pub(super) fn find_saved(reader: &str, target: &str) -> Option<String> {
    let mut t = percent_decode(target).unwrap_or_else(|| target.to_string());
    if t.len() >= 7 && t[..7].eq_ignore_ascii_case("file://") {
        t = t[7..].to_string();
        if t.len() >= 9 && t[..9].eq_ignore_ascii_case("localhost") {
            t = t[9..].to_string();
        }
    }
    let b = t.as_bytes();
    if b.len() >= 4 && b[0] == b'/' && b[1].is_ascii_alphabetic() && b[2] == b':' && (b[3] == b'/' || b[3] == b'\\') {
        t = t[1..].to_string();
    }
    if !is_workbook_name(&t) {
        return None;
    }
    let here = Path::new(reader).parent()?;
    let base = basename(&t);
    for candidate in [resolve_lexical(here, &t), here.join(&base)] {
        let (Some(dir), Some(name)) = (candidate.parent(), candidate.file_name()) else { continue };
        let Some(found) = file_in_dir(dir, &name.to_string_lossy()) else { continue };
        if same(&found, reader) {
            continue;
        }
        return Some(found);
    }
    None
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

/// One hidden copy a workbook keeps: `name` is the copy's sheet name (the
/// link as formulas spell it), `source` the file it mirrors and `sheet` the
/// sheet of that file.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct LinkEntry {
    pub name: String,
    pub source: String,
    pub sheet: String,
}

/// A file as it was when it was read: modified time and length.
pub(super) type Stamp = (SystemTime, u64);

pub(super) fn stamp_of(path: &str) -> Option<Stamp> {
    let meta = std::fs::metadata(path).ok()?;
    if !meta.is_file() {
        return None;
    }
    Some((meta.modified().ok()?, meta.len()))
}

/// What the workspace knows beside the loaded models.
#[derive(Default)]
pub struct WorkspaceState {
    /// The files the user put in the workspace, in the order they were added.
    members: Vec<String>,
    /// Files that could not be loaded, by upper-cased path, until they change.
    failed: HashMap<String, (Stamp, String)>,
    /// What a member that is not loaded reads, read from its file.
    tables: HashMap<String, (Stamp, Vec<String>)>,
}

pub(super) type Shared = Arc<Mutex<WorkspaceState>>;

fn lock_shared(shared: &Shared) -> MutexGuard<'_, WorkspaceState> {
    shared.lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

/// Values that changed in a workbook other than the one a command was for.
#[derive(Serialize, Debug)]
pub struct TabDeltas {
    pub tab: String,
    pub changed: Vec<CellDelta>,
}

/// What an edit did about the links it spelled.
#[derive(Serialize, Debug, Default)]
pub struct LinkReport {
    /// A workbook is read now that was not read before.
    pub changed: bool,
    /// Links that were left as typed, and why.
    pub notes: Vec<String>,
}

const BG: &str = "bg:";

fn bg_key(path: &str) -> String {
    format!("{BG}{path}")
}

pub(super) fn is_bg(key: &str) -> bool {
    key.starts_with(BG)
}

/// The model that stands for the file `path`: the tab that has it open (the
/// newest one, should several), else the model loaded without a tab.
pub(super) fn key_for_path(map: &HashMap<String, Loaded>, path: &str) -> Option<String> {
    if path.is_empty() {
        return None;
    }
    map.iter()
        .filter(|(_, l)| same(&l.path, path))
        .max_by_key(|(k, l)| (!is_bg(k), l.seq))
        .map(|(k, _)| k.clone())
}

/// The sheet called `name` (sheet names match without case), hidden copies
/// left out.
pub(super) fn real_sheet_index(m: &Model, name: &str) -> Option<u32> {
    if is_link_sheet_name(name) {
        return None;
    }
    let upper = name.to_uppercase();
    m.workbook.worksheets.iter().position(|ws| ws.name.to_uppercase() == upper).map(|i| i as u32)
}

/// Whether the workbook `from` reads `target`, directly or through the
/// workbooks it reads.
fn reads_transitively(map: &HashMap<String, Loaded>, from: &str, target: &str) -> bool {
    let mut seen: HashSet<String> = HashSet::new();
    let mut stack = vec![from.to_string()];
    while let Some(path) = stack.pop() {
        if !seen.insert(path.to_uppercase()) {
            continue;
        }
        let Some(key) = key_for_path(map, &path) else { continue };
        for entry in &map[&key].links {
            if same(&entry.source, target) {
                return true;
            }
            stack.push(entry.source.clone());
        }
    }
    false
}

/// The models that read the file `path`, the one under `except` left out.
/// Sorted, so the order things happen in does not depend on the map.
fn readers_of(map: &HashMap<String, Loaded>, path: &str, except: &str) -> Vec<String> {
    let mut keys: Vec<String> = map
        .iter()
        .filter(|(k, l)| k.as_str() != except && l.links.iter().any(|e| same(&e.source, path)))
        .map(|(k, _)| k.clone())
        .collect();
    keys.sort();
    keys
}

fn real_deltas(loaded: &Loaded, keys: impl IntoIterator<Item = Key>) -> Vec<CellDelta> {
    deltas_for(loaded, keys.into_iter().filter(|k| !core_links::is_link_sheet(&loaded.model, k.0)))
}

fn push_deltas(out: &mut Vec<TabDeltas>, tab: &str, changed: Vec<CellDelta>) {
    if changed.is_empty() {
        return;
    }
    match out.iter_mut().find(|t| t.tab == tab) {
        Some(existing) => existing.changed.extend(changed),
        None => out.push(TabDeltas { tab: tab.to_string(), changed }),
    }
}

pub(super) fn merge(out: &mut Vec<TabDeltas>, more: Vec<TabDeltas>) {
    for t in more {
        push_deltas(out, &t.tab, t.changed);
    }
}

fn delta_keys(deltas: &[CellDelta]) -> Vec<Key> {
    deltas.iter().map(|d| (d.sheet, d.row, d.column)).collect()
}

// ---------------------------------------------------------------------------
// Loading
// ---------------------------------------------------------------------------

/// Loads `path` and attaches the links the file was saved with. The model
/// comes back not yet evaluated. `loading` lists the files being loaded
/// further up: a saved link back to one of them is left unattached.
fn load_linked(inner: &Models, shared: &Shared, path: &str, loading: &mut Vec<String>) -> Result<Loaded, String> {
    let stamp = stamp_of(path);
    let model = load_from_xlsx(path, "en", "UTC", "en").map_err(|e| format!("ironcalc load: {e}"))?;
    let mut loaded = Loaded::new(model, path.to_string(), true);
    loaded.linkable = true;
    loaded.stamp = stamp;
    attach_saved_links(inner, shared, &mut loaded, loading);
    Ok(loaded)
}

/// A file as Excel (or a save from here) left it spells its links by number,
/// `[1]Model!B23`, and names the files in its external-link parts. Each such
/// link whose file is found gets its copy and is respelled with the file's
/// name, `[NOW.xlsx]Model`; the others stay as written and keep the values
/// the file was saved with.
fn attach_saved_links(inner: &Models, shared: &Shared, loaded: &mut Loaded, loading: &mut Vec<String>) {
    let names: Vec<String> = core_links::unresolved_links(&loaded.model)
        .into_iter()
        .filter(|n| split_link_sheet_name(n).is_some_and(|(book, _)| is_number(book)))
        .collect();
    if names.is_empty() || !is_absolute(&loaded.path) {
        return;
    }
    let Ok(bytes) = std::fs::read(&loaded.path) else { return };
    let Ok(saved) = external_links::read(&bytes) else { return };
    drop(bytes);
    let files = saved_link_files(&loaded.path, &saved);
    let mut wanted: Vec<String> = Vec::new();
    for name in &names {
        let Some(file) = split_link_sheet_name(name).and_then(|(book, _)| files.get(&book.parse::<u32>().ok()?)) else {
            continue;
        };
        if !wanted.iter().any(|w| same(w, file)) && !loading.iter().any(|l| same(l, file)) {
            wanted.push(file.clone());
        }
    }
    for file in &wanted {
        if let Err(e) = ensure_loaded(inner, shared, file, loading) {
            log::warn!("[workspace] {} reads {}, which could not be opened: {e}", loaded.path, file);
        }
    }
    let map = lock_recover(inner);
    for name in &names {
        let Some((book, sheet)) = split_link_sheet_name(name) else { continue };
        let Some(file) = book.parse::<u32>().ok().and_then(|n| files.get(&n)) else { continue };
        let base = basename(file);
        if base.contains(['[', ']']) || loading.iter().any(|l| same(l, file)) {
            continue;
        }
        let Some(source_key) = key_for_path(&map, file) else { continue };
        let source = &map[&source_key];
        if reads_transitively(&map, &source.path, &loaded.path) {
            continue;
        }
        let Some(source_sheet) = real_sheet_index(&source.model, sheet) else { continue };
        let from = source.model.workbook.worksheets[source_sheet as usize].name.clone();
        let linked = format!("[{base}]{from}");
        if loaded.links.iter().any(|e| same(&e.name, &linked)) {
            continue;
        }
        let Ok(copy) = loaded.model.add_link_sheet(name) else { continue };
        if let Err(e) =
            core_links::sync_link_sheet(&mut loaded.model, copy, &source.model, &source.cached, source_sheet)
        {
            log::warn!("[workspace] {name} of {}: {e}", loaded.path);
            let _ = loaded.model.delete_sheet(copy);
            continue;
        }
        // Under its number when the name is taken: the link works the same.
        let kept = if loaded.model.rename_link_sheet(copy, &linked).is_ok() { linked } else { name.clone() };
        loaded.links.push(LinkEntry { name: kept, source: source.path.clone(), sheet: from });
    }
    drop(map);
    loaded.sheets = loaded.model.workbook.get_worksheet_names();
}

/// The file behind each numbered link of a saved workbook. Two numbers for
/// one file: only the first counts.
fn saved_link_files(reader: &str, saved: &[external_links::ExternalLinkInfo]) -> HashMap<u32, String> {
    let mut taken: HashSet<String> = HashSet::new();
    let mut files: HashMap<u32, String> = HashMap::new();
    for link in saved {
        let Some(target) = &link.target else { continue };
        let Some(file) = find_saved(reader, target) else { continue };
        if taken.insert(file.to_uppercase()) {
            files.insert(link.index, file);
        }
    }
    files
}

/// Makes sure the file `path` has a model: nothing to do when a tab or an
/// earlier call loaded it, else it is loaded and evaluated here, without a
/// tab. Never call this with the model map locked.
pub(super) fn ensure_loaded(inner: &Models, shared: &Shared, path: &str, loading: &mut Vec<String>) -> Result<(), String> {
    if key_for_path(&lock_recover(inner), path).is_some() {
        return Ok(());
    }
    let stamp = stamp_of(path).ok_or_else(|| "the file was not found".to_string())?;
    if let Some((failed_at, error)) = lock_shared(shared).failed.get(&path.to_uppercase()) {
        if *failed_at == stamp {
            return Err(error.clone());
        }
    }
    let t = Instant::now();
    loading.push(path.to_string());
    // A panic while loading or evaluating this file must not take the
    // workbook that asked for it down with it.
    let result = std::panic::catch_unwind(AssertUnwindSafe(|| {
        let mut loaded = load_linked(inner, shared, path, loading)?;
        settle_pending(&mut loaded);
        Ok(loaded)
    }))
    .unwrap_or_else(|_| Err("the engine failed on this file".to_string()));
    loading.pop();
    match result {
        Ok(loaded) => {
            log::info!("[workspace] loaded {} without a tab: {} ms", path, t.elapsed().as_millis());
            let mut map = lock_recover(inner);
            if key_for_path(&map, path).is_none() {
                map.insert(bg_key(path), loaded);
            }
            Ok(())
        }
        Err(e) => {
            log::warn!("[workspace] {} could not be loaded: {e}", path);
            lock_shared(shared).failed.insert(path.to_uppercase(), (stamp, e.clone()));
            Err(e)
        }
    }
}

pub(super) struct Opened {
    pub sheets: Vec<String>,
    pub load_ms: u64,
    pub eval_ms: u64,
    pub others: Vec<TabDeltas>,
    pub needs_saving: bool,
}

/// Opens `path` for a tab, links attached. A model the workspace already
/// holds for the file is taken over when the file has not changed since.
pub(super) fn open_tab(inner: &Models, shared: &Shared, tab: &str, path: &str, defer: bool) -> Result<Opened, String> {
    let t = Instant::now();
    let stamp = stamp_of(path);
    let adopted = {
        let mut map = lock_recover(inner);
        let held = map
            .iter()
            .find(|(k, l)| is_bg(k) && same(&l.path, path) && stamp.is_some() && l.stamp == stamp)
            .map(|(k, _)| k.clone());
        held.and_then(|k| map.remove(&k))
    };
    let mut loaded = match adopted {
        Some(held) => held,
        None => load_linked(inner, shared, path, &mut vec![path.to_string()])?,
    };
    let load_ms = t.elapsed().as_millis() as u64;
    let t_eval = Instant::now();
    if !defer {
        settle_pending(&mut loaded);
    }
    let eval_ms = t_eval.elapsed().as_millis() as u64;
    log::info!(
        "[calc-engine] load {}: parse {} ms, evaluate {} ms{}, {} linked sheets",
        path,
        load_ms,
        eval_ms,
        if loaded.pending_eval { " (deferred)" } else { "" },
        loaded.links.len()
    );
    loaded.sheets = loaded.model.workbook.get_worksheet_names();
    loaded.seq = next_seq();
    let sheets = loaded.sheets.clone();
    let needs_saving = loaded.needs_saving;
    let others = install_tab(&mut lock_recover(inner), tab, loaded);
    Ok(Opened { sheets, load_ms, eval_ms, others, needs_saving })
}

/// Puts a freshly loaded model under its tab. From here on it stands for
/// its file: a model of the same file without a tab is dropped, and the
/// workbooks that read the file read this model.
pub(super) fn install_tab(map: &mut HashMap<String, Loaded>, tab: &str, loaded: Loaded) -> Vec<TabDeltas> {
    let path = loaded.path.clone();
    if !path.is_empty() {
        let stale: Vec<String> = map.iter().filter(|(k, l)| is_bg(k) && same(&l.path, &path)).map(|(k, _)| k.clone()).collect();
        for key in stale {
            map.remove(&key);
        }
    }
    map.insert(tab.to_string(), loaded);
    propagate(map, &path, Change::All)
}

// ---------------------------------------------------------------------------
// Edits
// ---------------------------------------------------------------------------

/// Which file each book name typed in a batch of edits means, by (folder as
/// written, upper-cased book name): the file, or why there is none.
type LinkPlan = HashMap<(String, String), Result<String, String>>;

fn book_matches(file_name: &str, book: &str) -> bool {
    if same(file_name, book) {
        return true;
    }
    if is_workbook_name(book) {
        return false;
    }
    file_name.rfind('.').is_some_and(|dot| same(&file_name[..dot], book))
}

/// The file a typed book name means for the workbook at `reader`.
fn resolve_book(reader: &str, read_already: &[String], known: &[String], dir: &str, book: &str) -> Result<String, String> {
    let here = Path::new(reader).parent().unwrap_or(Path::new(""));
    if !dir.is_empty() {
        let folder = resolve_lexical(here, dir);
        return book_in_dir(&folder, book).ok_or_else(|| format!("{book} was not found in {dir}"));
    }
    if book_matches(&basename(reader), book) {
        return Ok(reader.to_string());
    }
    if let Some(file) = read_already.iter().find(|p| book_matches(&basename(p), book)) {
        return Ok(file.clone());
    }
    let mut matches: Vec<&String> = Vec::new();
    for path in known {
        if book_matches(&basename(path), book) && !matches.iter().any(|m| same(m, path)) {
            matches.push(path);
        }
    }
    if matches.len() == 1 {
        return Ok(matches[0].clone());
    }
    if matches.len() > 1 {
        let beside: Vec<&&String> = matches.iter().filter(|m| Path::new(m.as_str()).parent() == Some(here)).collect();
        if beside.len() == 1 {
            return Ok((**beside[0]).clone());
        }
        return Err(format!("several workbooks are called {book}: write the folder in the link"));
    }
    book_in_dir(here, book)
        .ok_or_else(|| format!("{book} was not found: add it to the workspace, or put it next to this workbook"))
}

/// Finds and loads the workbooks the edits name. `None` when this workbook
/// cannot hold links (not opened with them, or not saved to a file yet).
fn prepare_links(inner: &Models, shared: &Shared, key: &str, edits: &[EditCell]) -> Option<LinkPlan> {
    let (reader, read_already, mut known) = {
        let map = lock_recover(inner);
        let loaded = map.get(key)?;
        if !loaded.linkable || !is_absolute(&loaded.path) {
            return None;
        }
        let mut known: Vec<String> = map.values().filter(|l| is_absolute(&l.path)).map(|l| l.path.clone()).collect();
        known.sort();
        (loaded.path.clone(), loaded.links.iter().map(|e| e.source.clone()).collect::<Vec<_>>(), known)
    };
    known.extend(lock_shared(shared).members.iter().cloned());
    let mut plan = LinkPlan::new();
    for edit in edits.iter().filter(|e| names_a_link(&e.value)) {
        for link in find_links(&edit.value) {
            if is_number(&link.book) {
                continue;
            }
            let slot = (link.dir.clone(), link.book.to_uppercase());
            if plan.contains_key(&slot) {
                continue;
            }
            let mut found = resolve_book(&reader, &read_already, &known, &link.dir, &link.book);
            if let Ok(file) = &found {
                if !same(file, &reader) {
                    if let Err(e) = ensure_loaded(inner, shared, file, &mut Vec::new()) {
                        found = Err(format!("{} could not be opened: {e}", basename(file)));
                    }
                }
            }
            plan.insert(slot, found);
        }
    }
    Some(plan)
}

fn note(notes: &mut Vec<String>, text: String) {
    if !notes.contains(&text) {
        notes.push(text);
    }
}

/// What a typed link prefix becomes: the prefix of the hidden copy, made
/// here when this workbook does not have it yet. `None` leaves the text as
/// typed, with the reason in `notes`.
fn link_prefix(
    map: &HashMap<String, Loaded>,
    reader: &mut Loaded,
    plan: &LinkPlan,
    link: &FoundLink,
    notes: &mut Vec<String>,
    created: &mut bool,
    sources: &mut Vec<String>,
) -> Option<String> {
    if is_number(&link.book) {
        return None;
    }
    let file = match plan.get(&(link.dir.clone(), link.book.to_uppercase()))? {
        Ok(file) => file,
        Err(why) => {
            note(notes, why.clone());
            // Kept as a link by name, in the spelling the engine reads
            // (a name with a space in it has to be quoted).
            return link.dir.is_empty().then(|| quoted_prefix(&format!("[{}]{}", link.book, link.sheet)));
        }
    };
    if same(file, &reader.path) {
        // A workbook's link to itself is an ordinary reference.
        return match real_sheet_index(&reader.model, &link.sheet) {
            Some(i) => Some(quoted_prefix(&reader.model.workbook.worksheets[i as usize].name)),
            None => {
                note(notes, format!("this workbook has no sheet called {}", link.sheet));
                None
            }
        };
    }
    let base = basename(file);
    if base.contains(['[', ']']) {
        note(notes, format!("{base} cannot be linked to: its name has a bracket in it"));
        return None;
    }
    let Some(source_key) = key_for_path(map, file) else {
        note(notes, format!("{base} could not be opened"));
        return None;
    };
    let source = &map[&source_key];
    if reads_transitively(map, &source.path, &reader.path) {
        note(notes, format!("{base} reads this workbook, so this workbook cannot read {base}"));
        return None;
    }
    let Some(source_sheet) = real_sheet_index(&source.model, &link.sheet) else {
        note(notes, format!("{base} has no sheet called {}", link.sheet));
        return None;
    };
    let from = source.model.workbook.worksheets[source_sheet as usize].name.clone();
    let name = format!("[{base}]{from}");
    if let Some(entry) = reader.links.iter().find(|e| same(&e.name, &name)) {
        if !same(&entry.source, &source.path) {
            note(notes, format!("another workbook called {base} is already linked here"));
            return None;
        }
        return Some(quoted_prefix(&entry.name));
    }
    let (copy, added) = match core_links::link_sheet_index(&reader.model, &name) {
        Some(copy) => (copy, false),
        None => match reader.model.add_link_sheet(&name) {
            Ok(copy) => (copy, true),
            Err(e) => {
                note(notes, format!("{name}: {e}"));
                return None;
            }
        },
    };
    if let Err(e) = core_links::sync_link_sheet(&mut reader.model, copy, &source.model, &source.cached, source_sheet) {
        if added {
            let _ = reader.model.delete_sheet(copy);
        }
        note(notes, format!("{name}: {e}"));
        return None;
    }
    // A new sheet, even one that had to be taken away again, leaves the
    // model in need of a full evaluation.
    *created = true;
    reader.links.push(LinkEntry { name: name.clone(), source: source.path.clone(), sheet: from });
    if !sources.iter().any(|s| same(s, &source.path)) {
        sources.push(source.path.clone());
    }
    Some(quoted_prefix(&name))
}

struct Edited {
    recalc_ms: u64,
    changed: Vec<CellDelta>,
    keys: Vec<Key>,
    sheets: Option<Vec<String>>,
    report: Option<LinkReport>,
    sources: Vec<String>,
}

/// The edits applied to `reader`, which is out of `map` meanwhile so the
/// workbooks it links to can be read next to it.
fn edit_locked(
    map: &HashMap<String, Loaded>,
    reader: &mut Loaded,
    edits: &mut [EditCell],
    plan: Option<&LinkPlan>,
) -> Result<Edited, String> {
    let settled = settle_pending(reader);
    let mut notes: Vec<String> = Vec::new();
    let mut created = false;
    let mut sources: Vec<String> = Vec::new();
    if let Some(plan) = plan {
        for edit in edits.iter_mut().filter(|e| names_a_link(&e.value)) {
            let text = rewrite_links(&edit.value, |link| {
                link_prefix(map, reader, plan, link, &mut notes, &mut created, &mut sources)
            });
            edit.value = text;
        }
    }
    let t = Instant::now();
    // A new copy is a new sheet: everything is evaluated again, and what
    // changed is found by comparing.
    let before = if created { Some(formula_values(&reader.model)) } else { None };
    let mut keys: Vec<Key> = Vec::with_capacity(edits.len());
    let mut failed: Option<String> = None;
    for edit in edits.iter() {
        match reader.model.set_user_input(edit.sheet, edit.row, edit.column, edit.value.clone()) {
            Ok(()) => keys.push((edit.sheet, edit.row, edit.column)),
            Err(e) => {
                failed = Some(e);
                break;
            }
        }
    }
    let dependents: Vec<Key> = match before {
        Some(before) => {
            reader.model.evaluate();
            reader.sheets = reader.model.workbook.get_worksheet_names();
            reader.index.mark_all_dirty();
            changed_since(before, &reader.model, Some).into_iter().map(|(k, _)| k).collect()
        }
        None => {
            if let Some(e) = failed {
                return Err(e);
            }
            let changed = reader.model.evaluate_edited(&keys);
            for k in keys.iter().chain(changed.iter()) {
                reader.index.mark_dirty(k.0);
            }
            changed
        }
    };
    if let Some(e) = failed {
        return Err(e);
    }
    let recalc_ms = t.elapsed().as_millis() as u64;
    let mut seen: HashSet<Key> = keys.iter().copied().collect();
    let mut all = keys;
    for k in dependents {
        if seen.insert(k) {
            all.push(k);
        }
    }
    let mut keys = delta_keys(&settled);
    let mut changed = settled;
    changed.extend(real_deltas(reader, all.iter().copied()));
    keys.extend(all);
    let report = if created || !notes.is_empty() { Some(LinkReport { changed: created, notes }) } else { None };
    Ok(Edited { recalc_ms, changed, keys, sheets: created.then(|| reader.sheets.clone()), report, sources })
}

/// Cell edits for the model under `key`: links the edits spell are found and
/// attached, the model is recalculated, and what changed is carried to the
/// workbooks that read this one. A link that cannot be made is left as
/// typed and reported; it never makes the edit fail.
pub(super) fn apply_edits(inner: &Models, shared: &Shared, key: &str, mut edits: Vec<EditCell>) -> Result<CalcEditResult, String> {
    let plan = if edits.iter().any(|e| names_a_link(&e.value)) { prepare_links(inner, shared, key, &edits) } else { None };
    let mut map = lock_recover(inner);
    let mut reader = map.remove(key).ok_or("no workbook loaded for tab")?;
    // A workbook that has no file yet cannot read another one (a link is
    // between files): the formula stays as typed, and the edit says why.
    let no_file = plan.is_none()
        && !is_absolute(&reader.path)
        && edits.iter().any(|e| names_a_link(&e.value) && find_links(&e.value).iter().any(|l| !is_number(&l.book)));
    let outcome = edit_locked(&map, &mut reader, &mut edits, plan.as_ref());
    let path = reader.path.clone();
    map.insert(key.to_string(), reader);
    let done = outcome?;
    if edits.len() > 1 {
        log::info!(
            "[calc-engine] edit batch: {} cells → recalc {} ms, {} values changed",
            edits.len(),
            done.recalc_ms,
            done.changed.len()
        );
    }
    let others = propagate(&mut map, &path, Change::Cells(done.keys));
    drop(map);
    if !done.sources.is_empty() {
        let mut ws = lock_shared(shared);
        for file in std::iter::once(&path).chain(done.sources.iter()) {
            if !ws.members.iter().any(|m| same(m, file)) {
                ws.members.push(file.clone());
            }
        }
    }
    let mut links = done.report;
    if no_file {
        note(
            &mut links.get_or_insert_with(LinkReport::default).notes,
            "this workbook has no file yet, so it cannot read another workbook: save it first, or start it with New workbook in the workspace panel".to_string(),
        );
    }
    Ok(CalcEditResult { recalc_ms: done.recalc_ms, changed: done.changed, others, sheets: done.sheets, links })
}

// ---------------------------------------------------------------------------
// Carrying changes to the workbooks that read a workbook
// ---------------------------------------------------------------------------

pub(super) enum Change {
    /// These cells changed.
    Cells(Vec<Key>),
    /// Anything may have changed: compare every sheet that is read.
    All,
}

/// Brings `reader`'s copies of `source`'s sheets up to date and recalculates
/// what reads them. Returns the formula cells of `reader` that changed.
fn follow_values(reader: &mut Loaded, source: &Loaded, change: &Change) -> Vec<Key> {
    let entries: Vec<LinkEntry> = reader.links.iter().filter(|e| same(&e.source, &source.path)).cloned().collect();
    let mut touched: Vec<Key> = Vec::new();
    for entry in &entries {
        let Some(source_sheet) = real_sheet_index(&source.model, &entry.sheet) else { continue };
        let Some(copy) = core_links::link_sheet_index(&reader.model, &entry.name) else { continue };
        let changed = match change {
            Change::All => {
                core_links::sync_link_sheet(&mut reader.model, copy, &source.model, &source.cached, source_sheet)
            }
            Change::Cells(keys) => {
                let cells: Vec<(i32, i32)> = keys.iter().filter(|k| k.0 == source_sheet).map(|k| (k.1, k.2)).collect();
                if cells.is_empty() {
                    continue;
                }
                core_links::push_link_cells(&mut reader.model, copy, &source.model, &source.cached, source_sheet, &cells)
            }
        };
        match changed {
            Ok(cells) => touched.extend(cells.into_iter().map(|(row, column)| (copy, row, column))),
            Err(e) => log::warn!("[workspace] {} in {}: {e}", entry.name, reader.path),
        }
    }
    // A model still waiting for its first evaluation picks the new values
    // up when that runs.
    if touched.is_empty() || reader.pending_eval {
        return Vec::new();
    }
    let changed = reader.model.evaluate_edited(&touched);
    for k in &changed {
        reader.index.mark_dirty(k.0);
    }
    changed
}

/// Carries a change in the workbook at `source_path` to every workbook that
/// reads it, and from those to their readers. Returns what changed in the
/// workbooks that are open in a tab.
pub(super) fn propagate(map: &mut HashMap<String, Loaded>, source_path: &str, change: Change) -> Vec<TabDeltas> {
    let mut out: Vec<TabDeltas> = Vec::new();
    if source_path.is_empty() {
        return out;
    }
    let mut queue: VecDeque<(String, Change)> = VecDeque::new();
    queue.push_back((source_path.to_string(), change));
    let mut steps = 0usize;
    while let Some((path, change)) = queue.pop_front() {
        steps += 1;
        if steps > 10_000 {
            log::error!("[workspace] stopped carrying changes from {source_path}: too many steps");
            break;
        }
        if let Change::Cells(keys) = &change {
            if keys.is_empty() {
                continue;
            }
        }
        let Some(source_key) = key_for_path(map, &path) else { continue };
        for reader_key in readers_of(map, &path, &source_key) {
            let Some(mut reader) = map.remove(&reader_key) else { continue };
            let source = &map[&source_key];
            match std::panic::catch_unwind(AssertUnwindSafe(|| follow_values(&mut reader, source, &change))) {
                Ok(changed) => {
                    if !changed.is_empty() {
                        if !is_bg(&reader_key) {
                            push_deltas(&mut out, &reader_key, real_deltas(&reader, changed.iter().copied()));
                        }
                        queue.push_back((reader.path.clone(), Change::Cells(changed)));
                    }
                    map.insert(reader_key, reader);
                }
                // Mid-change and not to be trusted: dropped, as a tab's own
                // model is when the engine fails under it.
                Err(_) => log::error!("[workspace] the engine failed in {}; its model was dropped", reader.path),
            }
        }
    }
    out
}

/// What a row or column change in a sheet of `source` means for `reader`:
/// the same change in its copy of that sheet, so its formulas keep pointing
/// at the same cells. `None` when it does not read that sheet. Returns the
/// cells of `reader` whose value or formula text changed.
fn follow_structural(
    reader: &mut Loaded,
    source: &Loaded,
    sheet_name: &str,
    op: &str,
    index: i32,
    count: i32,
) -> Option<Vec<Key>> {
    let entries: Vec<LinkEntry> = reader.links.iter().filter(|e| same(&e.source, &source.path)).cloned().collect();
    let entry = entries.iter().find(|e| same(&e.sheet, sheet_name))?;
    let copy = core_links::link_sheet_index(&reader.model, &entry.name)?;
    let mut keys = delta_keys(&settle_pending(reader));
    let before = formula_values(&reader.model);
    let moved = match op {
        "insert_rows" => reader.model.insert_rows(copy, index, count),
        "delete_rows" => reader.model.delete_rows(copy, index, count),
        "insert_columns" => reader.model.insert_columns(copy, index, count),
        "delete_columns" => reader.model.delete_columns(copy, index, count),
        other => Err(format!("unknown structural op: {other}")),
    };
    if let Err(e) = moved {
        log::warn!("[workspace] {op} not followed by {} in {}: {e}", entry.name, reader.path);
    }
    for e in &entries {
        let (Some(source_sheet), Some(copy)) =
            (real_sheet_index(&source.model, &e.sheet), core_links::link_sheet_index(&reader.model, &e.name))
        else {
            continue;
        };
        let _ = core_links::sync_link_sheet(&mut reader.model, copy, &source.model, &source.cached, source_sheet);
    }
    reader.model.evaluate();
    reader.index.mark_all_dirty();
    reader.needs_saving = true;
    reader.rewrite_links = true;
    keys.extend(changed_since(before, &reader.model, Some).into_iter().map(|(k, _)| k));
    keys.extend(core_links::link_formula_cells(&reader.model));
    keys.sort_unstable();
    keys.dedup();
    Some(keys)
}

/// What a renamed sheet of `source` means for `reader`: its copy takes the
/// new name and every link to it is respelled.
fn follow_rename(reader: &mut Loaded, source: &Loaded, old: &str, new: &str) -> Option<Vec<Key>> {
    let at = reader.links.iter().position(|e| same(&e.source, &source.path) && same(&e.sheet, old))?;
    let entry = reader.links[at].clone();
    let copy = core_links::link_sheet_index(&reader.model, &entry.name)?;
    let (book, _) = split_link_sheet_name(&entry.name)?;
    let renamed = format!("[{book}]{new}");
    let mut keys = delta_keys(&settle_pending(reader));
    let before = formula_values(&reader.model);
    match reader.model.rename_link_sheet(copy, &renamed) {
        Ok(()) => reader.links[at].name = renamed,
        Err(e) => log::warn!("[workspace] {} keeps its name in {}: {e}", entry.name, reader.path),
    }
    reader.links[at].sheet = new.to_string();
    reader.model.evaluate();
    reader.sheets = reader.model.workbook.get_worksheet_names();
    reader.index.mark_all_dirty();
    reader.needs_saving = true;
    reader.rewrite_links = true;
    keys.extend(changed_since(before, &reader.model, Some).into_iter().map(|(k, _)| k));
    keys.extend(core_links::link_formula_cells(&reader.model));
    keys.sort_unstable();
    keys.dedup();
    Some(keys)
}

/// Runs `follow` on every workbook that reads the model under `source_key`.
/// Returns what changed in open tabs, and the tabs whose formulas moved.
fn follow_in_readers(
    map: &mut HashMap<String, Loaded>,
    source_key: &str,
    follow: impl Fn(&mut Loaded, &Loaded) -> Option<Vec<Key>>,
) -> (Vec<TabDeltas>, Vec<String>) {
    let mut out: Vec<TabDeltas> = Vec::new();
    let mut moved: Vec<String> = Vec::new();
    let Some(path) = map.get(source_key).map(|l| l.path.clone()).filter(|p| !p.is_empty()) else {
        return (out, moved);
    };
    let mut onward: Vec<(String, Vec<Key>)> = Vec::new();
    for reader_key in readers_of(map, &path, source_key) {
        let Some(mut reader) = map.remove(&reader_key) else { continue };
        let source = &map[source_key];
        match std::panic::catch_unwind(AssertUnwindSafe(|| follow(&mut reader, source))) {
            Ok(Some(keys)) => {
                if !is_bg(&reader_key) {
                    push_deltas(&mut out, &reader_key, real_deltas(&reader, keys.iter().copied()));
                    moved.push(reader_key.clone());
                }
                onward.push((reader.path.clone(), keys));
                map.insert(reader_key, reader);
            }
            Ok(None) => {
                map.insert(reader_key, reader);
            }
            Err(_) => log::error!("[workspace] the engine failed in {}; its model was dropped", reader.path),
        }
    }
    for (path, keys) in onward {
        merge(&mut out, propagate(map, &path, Change::Cells(keys)));
    }
    (out, moved)
}

/// Rows or columns were inserted or deleted in sheet `sheet_name` of the
/// model under `source_key`: the links that point there follow.
pub(super) fn mirror_structural(
    map: &mut HashMap<String, Loaded>,
    source_key: &str,
    sheet_name: &str,
    op: &str,
    index: i32,
    count: i32,
) -> (Vec<TabDeltas>, Vec<String>) {
    follow_in_readers(map, source_key, |reader, source| follow_structural(reader, source, sheet_name, op, index, count))
}

/// Sheet `old` of the model under `source_key` is called `new` now.
pub(super) fn mirror_rename(map: &mut HashMap<String, Loaded>, source_key: &str, old: &str, new: &str) -> (Vec<TabDeltas>, Vec<String>) {
    follow_in_readers(map, source_key, |reader, source| follow_rename(reader, source, old, new))
}

// ---------------------------------------------------------------------------
// Closing and saving
// ---------------------------------------------------------------------------

/// Drops the models without a tab that nothing reads any more. One whose
/// formulas followed a change and were not saved yet is kept.
fn collect_unread(map: &mut HashMap<String, Loaded>) {
    loop {
        let unread: Vec<String> = map
            .iter()
            .filter(|(k, l)| is_bg(k) && !l.needs_saving && readers_of(map, &l.path, k).is_empty())
            .map(|(k, _)| k.clone())
            .collect();
        if unread.is_empty() {
            return;
        }
        for key in unread {
            map.remove(&key);
        }
    }
}

/// A tab was closed. A workbook that others read stays loaded without the
/// tab: as it is when it had nothing unsaved, as the file has it otherwise.
pub(super) fn close_tab(inner: &Models, shared: &Shared, tab: &str, dirty: bool) -> Vec<TabDeltas> {
    let mut map = lock_recover(inner);
    let Some(mut loaded) = map.remove(tab) else { return Vec::new() };
    let path = loaded.path.clone();
    let read = is_absolute(&path) && key_for_path(&map, &path).is_none() && !readers_of(&map, &path, tab).is_empty();
    let mut others = Vec::new();
    if read && !dirty {
        let keys = delta_keys(&settle_pending(&mut loaded));
        map.insert(bg_key(&path), loaded);
        others = propagate(&mut map, &path, Change::Cells(keys));
    } else if read {
        drop(loaded);
        drop(map);
        if let Err(e) = ensure_loaded(inner, shared, &path, &mut Vec::new()) {
            log::warn!("[workspace] {path} is read by open workbooks and could not be loaded again: {e}");
        }
        map = lock_recover(inner);
        others = propagate(&mut map, &path, Change::All);
    }
    collect_unread(&mut map);
    others
}

/// The tab's workbook was written to `path`.
pub(super) fn mark_saved(inner: &Models, shared: &Shared, tab: &str, path: &str) -> Vec<TabDeltas> {
    let mut map = lock_recover(inner);
    let Some(loaded) = map.get_mut(tab) else { return Vec::new() };
    let old = std::mem::replace(&mut loaded.path, path.to_string());
    loaded.stamp = stamp_of(path);
    loaded.needs_saving = false;
    if old.is_empty() || same(&old, path) {
        return Vec::new();
    }
    // Saved under another name. A model without a tab for the new file is
    // out of date now, and the old file is what its readers go on reading.
    let stale: Vec<String> = map.iter().filter(|(k, l)| is_bg(k) && same(&l.path, path)).map(|(k, _)| k.clone()).collect();
    for key in stale {
        map.remove(&key);
    }
    let mut others = propagate(&mut map, path, Change::All);
    let old_is_read = is_absolute(&old) && key_for_path(&map, &old).is_none() && !readers_of(&map, &old, "").is_empty();
    drop(map);
    if old_is_read {
        if let Err(e) = ensure_loaded(inner, shared, &old, &mut Vec::new()) {
            log::warn!("[workspace] {old} is read by open workbooks and could not be loaded: {e}");
        }
        merge(&mut others, propagate(&mut lock_recover(inner), &old, Change::All));
    }
    let mut ws = lock_shared(shared);
    if ws.members.iter().any(|m| same(m, &old)) && !ws.members.iter().any(|m| same(m, path)) {
        ws.members.push(path.to_string());
    }
    others
}

/// One formula cell a save writes with the file's spelling of its links.
#[derive(Serialize, Debug)]
pub struct PlanCell {
    /// Sheet name.
    pub sheet: String,
    /// 0-based.
    pub r: u32,
    pub c: u32,
    /// Formula text without the leading `=`.
    pub f: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub v: Option<serde_json::Value>,
}

#[derive(Serialize, Debug)]
pub struct PlanSheet {
    pub name: String,
    pub areas: Vec<[u32; 4]>,
    pub cells: Vec<core_links::LinkCacheCell>,
}

/// One external-link part of the saved file, in the save patch's shape.
#[derive(Serialize, Debug)]
pub struct PlanLink {
    pub index: u32,
    /// Only for a link the file does not have yet.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub target: Option<String>,
    pub sheets: Vec<PlanSheet>,
}

/// What a save has to add to its patch for the links of a workbook.
#[derive(Serialize, Debug, Default)]
#[serde(rename_all = "camelCase")]
pub struct LinkSavePlan {
    /// Every formula cell that names another workbook, as the file spells it.
    pub cells: Vec<PlanCell>,
    /// Write all of `cells`, not only the ones the session edited: rows
    /// moved or a sheet was renamed in a workbook this one reads.
    pub rewrite_all: bool,
    pub external_links: Vec<PlanLink>,
    /// The workbooks read.
    pub reads: Vec<String>,
}

/// Builds the save plan for `loaded`, to be written to `target_path` on top
/// of a file that already holds the links `saved`.
pub(super) fn save_plan(loaded: &Loaded, saved: &[external_links::ExternalLinkInfo], target_path: &str) -> LinkSavePlan {
    let model = &loaded.model;
    let cells = core_links::link_formula_cells(model);
    if cells.is_empty() {
        return LinkSavePlan::default();
    }
    let files = saved_link_files(&loaded.path, saved);
    let target_dir = Path::new(target_path).parent().map(|p| p.to_string_lossy().to_string()).unwrap_or_default();
    // The number of a workbook that is read: the one the file already uses
    // for it, else the next free one.
    let mut fresh: Vec<(String, u32, String)> = Vec::new(); // (what, number, target)
    let mut number_for = |what: &str, target: &dyn Fn() -> String| -> u32 {
        if let Some((_, n, _)) = fresh.iter().find(|(w, _, _)| same(w, what)) {
            return *n;
        }
        let n = (saved.len() + fresh.len() + 1) as u32;
        fresh.push((what.to_string(), n, target()));
        n
    };
    let mut number_of_copy = |name: &str| -> Option<u32> {
        let (book, _) = split_link_sheet_name(name)?;
        if is_number(book) {
            return book.parse().ok();
        }
        match loaded.links.iter().find(|e| same(&e.name, name)) {
            Some(entry) => Some(match files.iter().find(|(_, file)| same(file, &entry.source)) {
                Some((n, _)) => *n,
                None => number_for(&entry.source, &|| link_target(&entry.source, &target_dir)),
            }),
            // Named, and no workbook behind it: the name is all there is.
            None => Some(number_for(&format!("?{book}"), &|| encode_segment(book))),
        }
    };
    let mut kept: BTreeMap<u32, Vec<String>> = BTreeMap::new();
    let mut out_cells: Vec<PlanCell> = Vec::with_capacity(cells.len());
    for (sheet, row, column) in cells {
        let Ok(Some(text)) = model.get_cell_formula(sheet, row, column) else { continue };
        let body = text.strip_prefix('=').unwrap_or(&text);
        let f = rewrite_links(body, |link| {
            if !link.dir.is_empty() {
                return None;
            }
            let name = format!("[{}]{}", link.book, link.sheet);
            let numbered = is_number(&link.book);
            if numbered && core_links::link_sheet_index(model, &name).is_none() {
                // Saved with the file and not attached: left as it is.
                return None;
            }
            let n = number_of_copy(&name)?;
            let sheets = kept.entry(n).or_default();
            if !sheets.iter().any(|s| same(s, &link.sheet)) {
                sheets.push(link.sheet.clone());
            }
            if numbered { None } else { Some(numbered_prefix(n, &link.sheet)) }
        });
        out_cells.push(PlanCell {
            sheet: model.workbook.worksheets[sheet as usize].name.clone(),
            r: (row - 1) as u32,
            c: (column - 1) as u32,
            f,
            v: core_links::patch_value(model, sheet, row, column),
        });
    }
    let mut parts: BTreeMap<u32, Vec<PlanSheet>> = BTreeMap::new();
    for cache in core_links::link_cache(model) {
        let Some((_, sheet)) = split_link_sheet_name(&cache.sheet) else { continue };
        let Some(n) = number_of_copy(&cache.sheet) else { continue };
        parts.entry(n).or_default().push(PlanSheet { name: sheet.to_string(), areas: cache.areas, cells: cache.cells });
    }
    for (n, sheets) in kept {
        let part = parts.entry(n).or_default();
        for sheet in sheets {
            if !part.iter().any(|s| same(&s.name, &sheet)) {
                part.push(PlanSheet { name: sheet, areas: Vec::new(), cells: Vec::new() });
            }
        }
    }
    let external_links = parts
        .into_iter()
        .map(|(index, sheets)| PlanLink {
            index,
            target: fresh.iter().find(|(_, n, _)| *n == index).map(|(_, _, target)| target.clone()),
            sheets,
        })
        .collect();
    let mut reads: Vec<String> = Vec::new();
    for entry in &loaded.links {
        if !reads.iter().any(|r| same(r, &entry.source)) {
            reads.push(entry.source.clone());
        }
    }
    LinkSavePlan { cells: out_cells, rewrite_all: loaded.rewrite_links, external_links, reads }
}

// ---------------------------------------------------------------------------
// The workspace list
// ---------------------------------------------------------------------------

#[derive(Serialize, Debug)]
#[serde(rename_all = "camelCase")]
pub struct WorkspaceMember {
    pub path: String,
    pub name: String,
    /// `open` (in a tab), `loaded` (without a tab), `idle` (on disk, not
    /// loaded), `missing`, or `failed` (could not be loaded).
    pub state: &'static str,
    /// On the workspace list. A workbook that is only loaded because an open
    /// one reads it is shown too, with `false`.
    pub member: bool,
    /// The workbooks this one reads, and the ones that read it.
    pub reads: Vec<String>,
    pub read_by: Vec<String>,
    /// Links of this workbook whose sheet is gone from the workbook read.
    pub broken: Vec<String>,
    /// Its formulas followed a change in a workbook it reads and the file
    /// does not have that yet.
    pub needs_saving: bool,
    pub error: Option<String>,
}

/// What a file reads, from the links it was saved with.
fn reads_on_disk(shared: &Shared, path: &str, stamp: Stamp) -> Vec<String> {
    let slot = path.to_uppercase();
    if let Some((at, reads)) = lock_shared(shared).tables.get(&slot) {
        if *at == stamp {
            return reads.clone();
        }
    }
    let mut reads: Vec<String> = std::fs::read(path)
        .ok()
        .and_then(|bytes| external_links::read(&bytes).ok())
        .map(|saved| saved_link_files(path, &saved).into_values().collect())
        .unwrap_or_default();
    reads.sort();
    lock_shared(shared).tables.insert(slot, (stamp, reads.clone()));
    reads
}

pub(super) fn status(inner: &Models, shared: &Shared) -> Vec<WorkspaceMember> {
    let (members, failed) = {
        let ws = lock_shared(shared);
        (ws.members.clone(), ws.failed.clone())
    };
    let mut out: Vec<WorkspaceMember> = Vec::new();
    {
        let map = lock_recover(inner);
        let mut paths: Vec<String> = Vec::new();
        for path in members.iter().cloned().chain({
            let mut loaded: Vec<String> = map.values().filter(|l| is_absolute(&l.path)).map(|l| l.path.clone()).collect();
            loaded.sort();
            loaded
        }) {
            if !paths.iter().any(|p| same(p, &path)) {
                paths.push(path);
            }
        }
        for path in paths {
            let member = members.iter().any(|m| same(m, &path));
            let mut row = WorkspaceMember {
                name: basename(&path),
                state: "idle",
                member,
                reads: Vec::new(),
                read_by: Vec::new(),
                broken: Vec::new(),
                needs_saving: false,
                error: None,
                path,
            };
            if let Some(key) = key_for_path(&map, &row.path) {
                let loaded = &map[&key];
                row.state = if is_bg(&key) { "loaded" } else { "open" };
                row.needs_saving = loaded.needs_saving;
                for entry in &loaded.links {
                    if !row.reads.iter().any(|r| same(r, &entry.source)) {
                        row.reads.push(entry.source.clone());
                    }
                    let there = key_for_path(&map, &entry.source)
                        .is_some_and(|k| real_sheet_index(&map[&k].model, &entry.sheet).is_some());
                    if !there {
                        row.broken.push(entry.name.clone());
                    }
                }
            }
            out.push(row);
        }
    }
    // Files are read with the model map unlocked.
    for row in out.iter_mut().filter(|r| r.state == "idle") {
        match stamp_of(&row.path) {
            None => row.state = "missing",
            Some(stamp) => match failed.get(&row.path.to_uppercase()) {
                Some((at, error)) if *at == stamp => {
                    row.state = "failed";
                    row.error = Some(error.clone());
                }
                _ => row.reads = reads_on_disk(shared, &row.path, stamp),
            },
        }
    }
    let read: Vec<(String, Vec<String>)> = out.iter().map(|r| (r.path.clone(), r.reads.clone())).collect();
    for row in out.iter_mut() {
        row.read_by = read.iter().filter(|(_, reads)| reads.iter().any(|r| same(r, &row.path))).map(|(p, _)| p.clone()).collect();
    }
    out
}

/// Dropped files and folders as the workbooks they stand for: a folder
/// gives the workbooks directly in it.
pub(super) fn expand_paths(paths: &[String]) -> Vec<String> {
    let usable = |name: &str| is_workbook_name(name) && !name.starts_with("~$") && !name.starts_with('.');
    let mut out: Vec<String> = Vec::new();
    let mut add = |path: String| {
        if !out.iter().any(|p| same(p, &path)) {
            out.push(path);
        }
    };
    for path in paths {
        let p = Path::new(path);
        if p.is_dir() {
            let mut inside: Vec<String> = std::fs::read_dir(p)
                .map(|entries| {
                    entries
                        .flatten()
                        .filter(|e| usable(&e.file_name().to_string_lossy()) && e.path().is_file())
                        .map(|e| e.path().to_string_lossy().to_string())
                        .collect()
                })
                .unwrap_or_default();
            inside.sort();
            inside.into_iter().for_each(&mut add);
        } else if p.is_file() && usable(&basename(path)) {
            add(path.clone());
        }
    }
    out
}

/// A new, empty workbook file with the given sheets.
pub(super) fn create_workbook_bytes(sheets: &[String]) -> Result<Vec<u8>, String> {
    let mut model = Model::new_empty("workbook", "en", "UTC", "en")?;
    if let Some(first) = sheets.first() {
        model.rename_sheet_by_index(0, first)?;
    }
    for name in sheets.iter().skip(1) {
        model.add_sheet(name)?;
    }
    // What Excel gives a new workbook.
    if let Some(font) = model.workbook.styles.fonts.first_mut() {
        font.name = "Calibri".to_string();
        font.sz = 11;
    }
    model.evaluate();
    ironcalc::export::save_xlsx_to_writer(&model, std::io::Cursor::new(Vec::new()))
        .map(|cursor| cursor.into_inner())
        .map_err(|e| format!("{e:?}"))
}

// ---------------------------------------------------------------------------
// Commands
// ---------------------------------------------------------------------------

async fn blocking<T: Send + 'static>(f: impl FnOnce() -> T + Send + 'static) -> Result<T, String> {
    tauri::async_runtime::spawn_blocking(f).await.map_err(|e| format!("workspace task join error: {e}"))
}

/// Puts files on the workspace list. Returns the list.
#[tauri::command]
pub async fn workspace_add_members(state: State<'_, CalcEngineState>, paths: Vec<String>) -> Result<Vec<WorkspaceMember>, String> {
    let (inner, shared) = (state.inner.clone(), state.workspace.clone());
    blocking(move || {
        {
            let mut ws = lock_shared(&shared);
            for path in paths.into_iter().filter(|p| is_absolute(p) && is_workbook_name(p)) {
                if !ws.members.iter().any(|m| same(m, &path)) {
                    ws.members.push(path);
                }
            }
        }
        status(&inner, &shared)
    })
    .await
}

/// Takes a file off the workspace list. A workbook that an open one reads
/// stays loaded. Returns the list.
/// A workbook of the workspace as the agent reads it: the key the engine
/// commands (`calc_engine_workbook_index`, `calc_engine_sheet_window`,
/// `calc_engine_query_rows`) find its model under, and its sheets in engine
/// order (the position is the sheet index those commands take).
#[derive(Serialize, Debug)]
pub struct WorkspaceModel {
    pub key: String,
    pub sheets: Vec<WorkspaceModelSheet>,
}

#[derive(Serialize, Debug)]
pub struct WorkspaceModelSheet {
    pub name: String,
    /// The hidden copy of a sheet of another workbook: not this workbook's.
    pub link: bool,
}

/// The model for `path`, loaded without a tab when nothing holds one. Only
/// for files on the workspace list or already loaded: the agent names a
/// workbook, and that must not become a way to open any file on disk.
pub(super) fn model_for_reading(inner: &Models, shared: &Shared, path: &str) -> Result<WorkspaceModel, String> {
    if !is_absolute(path) || !is_workbook_name(path) {
        return Err("not a workbook path".to_string());
    }
    let listed = lock_shared(shared).members.iter().any(|m| same(m, path));
    if !listed && key_for_path(&lock_recover(inner), path).is_none() {
        return Err(format!("{} is not in the workspace", basename(path)));
    }
    ensure_loaded(inner, shared, path, &mut Vec::new())?;
    let map = lock_recover(inner);
    let key = key_for_path(&map, path).ok_or("the workbook could not be loaded")?;
    // A tab whose first evaluate is still queued is read with the file's
    // saved values until its grid has run it (settling it here would take
    // the changed values away from that grid).
    let loaded = map.get(&key).ok_or("the workbook could not be loaded")?;
    let sheets = loaded
        .model
        .workbook
        .get_worksheet_names()
        .into_iter()
        .map(|name| WorkspaceModelSheet { link: is_link_sheet_name(&name), name })
        .collect();
    Ok(WorkspaceModel { key, sheets })
}

#[tauri::command]
pub async fn workspace_model(state: State<'_, CalcEngineState>, path: String) -> Result<WorkspaceModel, String> {
    let (inner, shared) = (state.inner.clone(), state.workspace.clone());
    blocking(move || model_for_reading(&inner, &shared, &path)).await?
}

/// Replaces the list: another saved workspace is the one shown now. Open
/// tabs stay as they are, and so does every model an open workbook reads.
#[tauri::command]
pub async fn workspace_set_members(state: State<'_, CalcEngineState>, paths: Vec<String>) -> Result<Vec<WorkspaceMember>, String> {
    let (inner, shared) = (state.inner.clone(), state.workspace.clone());
    blocking(move || {
        {
            let mut ws = lock_shared(&shared);
            ws.members.clear();
            for path in paths.into_iter().filter(|p| is_absolute(p) && is_workbook_name(p)) {
                if !ws.members.iter().any(|m| same(m, &path)) {
                    ws.members.push(path);
                }
            }
        }
        collect_unread(&mut lock_recover(&inner));
        status(&inner, &shared)
    })
    .await
}

#[tauri::command]
pub async fn workspace_remove_member(state: State<'_, CalcEngineState>, path: String) -> Result<Vec<WorkspaceMember>, String> {
    let (inner, shared) = (state.inner.clone(), state.workspace.clone());
    blocking(move || {
        lock_shared(&shared).members.retain(|m| !same(m, &path));
        collect_unread(&mut lock_recover(&inner));
        status(&inner, &shared)
    })
    .await
}

#[tauri::command]
pub async fn workspace_status(state: State<'_, CalcEngineState>) -> Result<Vec<WorkspaceMember>, String> {
    let (inner, shared) = (state.inner.clone(), state.workspace.clone());
    blocking(move || status(&inner, &shared)).await
}

#[tauri::command]
pub async fn workspace_expand_paths(paths: Vec<String>) -> Result<Vec<String>, String> {
    blocking(move || expand_paths(&paths)).await
}

/// Writes a new, empty workbook to `path` (the save dialog already asked
/// about replacing a file that is there).
#[tauri::command]
pub async fn workspace_create_workbook(path: String, sheets: Option<Vec<String>>) -> Result<(), String> {
    blocking(move || {
        if !is_absolute(&path) {
            return Err("a new workbook needs a full path".to_string());
        }
        let sheets = sheets.filter(|s| !s.is_empty()).unwrap_or_else(|| vec!["Sheet1".to_string()]);
        let bytes = create_workbook_bytes(&sheets)?;
        super::commands::write_atomic(Path::new(&path), &bytes)
    })
    .await?
}

/// The links part of a save: see `LinkSavePlan`. `target_path` is the file
/// the save writes (the tab's own file, or the new one of a Save As).
#[tauri::command]
pub async fn calc_engine_link_save_plan(
    state: State<'_, CalcEngineState>,
    tab: String,
    target_path: String,
) -> Result<LinkSavePlan, String> {
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| {
        let path = match lock_recover(inner).get(&key) {
            Some(loaded) if is_absolute(&loaded.path) => loaded.path.clone(),
            _ => return Ok(LinkSavePlan::default()),
        };
        // The patch is applied to the bytes the session started from.
        let saved = super::commands::patch_base_bytes(Path::new(&path))
            .ok()
            .and_then(|bytes| external_links::read(&bytes).ok())
            .unwrap_or_default();
        let map = lock_recover(inner);
        let loaded = map.get(&key).ok_or("no workbook loaded for tab")?;
        Ok(save_plan(loaded, &saved, &target_path))
    })
    .await
}

/// The tab's workbook was saved to `path`.
#[tauri::command]
pub async fn calc_engine_mark_saved(state: State<'_, CalcEngineState>, tab: String, path: String) -> Result<Vec<TabDeltas>, String> {
    let shared = state.workspace.clone();
    let key = tab.clone();
    run_contained(state.inner.clone(), tab, move |inner| Ok(mark_saved(inner, &shared, &key, &path))).await
}

#[cfg(test)]
mod tests {
    use super::super::calc_engine::{cell_value_to_json, sheet_op, structural};
    use super::*;
    use serde_json::json;

    /// A fresh folder for one test.
    fn folder(name: &str) -> PathBuf {
        let dir = std::env::temp_dir().join(format!("gridpath-workspace-{name}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    /// Writes a workbook file: its sheets by name, and cells as
    /// (sheet, row, column, what is typed).
    fn book(dir: &Path, file: &str, sheets: &[&str], cells: &[(u32, i32, i32, &str)]) -> String {
        let mut m = Model::new_empty("book", "en", "UTC", "en").unwrap();
        if sheets[0] != "Sheet1" {
            m.rename_sheet_by_index(0, sheets[0]).unwrap();
        }
        for name in &sheets[1..] {
            m.add_sheet(name).unwrap();
        }
        for (sheet, row, column, input) in cells {
            m.set_user_input(*sheet, *row, *column, input.to_string()).unwrap();
        }
        m.evaluate();
        let path = dir.join(file).to_string_lossy().to_string();
        ironcalc::export::save_to_xlsx(&m, &path).unwrap();
        path
    }

    struct Ws {
        inner: Models,
        shared: Shared,
    }

    impl Ws {
        fn new() -> Self {
            Ws { inner: Default::default(), shared: Default::default() }
        }
        fn open(&self, tab: &str, path: &str) -> Opened {
            open_tab(&self.inner, &self.shared, tab, path, false).unwrap()
        }
        fn edit(&self, tab: &str, sheet: u32, row: i32, column: i32, value: &str) -> CalcEditResult {
            apply_edits(&self.inner, &self.shared, tab, vec![EditCell { sheet, row, column, value: value.to_string() }])
                .unwrap()
        }
        fn value(&self, key: &str, sheet: u32, row: i32, column: i32) -> serde_json::Value {
            lock_recover(&self.inner)[key].model.get_cell_value_by_index(sheet, row, column).map(cell_value_to_json).unwrap()
        }
        fn formula(&self, key: &str, sheet: u32, row: i32, column: i32) -> String {
            lock_recover(&self.inner)[key].model.get_cell_formula(sheet, row, column).unwrap().unwrap_or_default()
        }
        fn keys(&self) -> Vec<String> {
            let mut keys: Vec<String> = lock_recover(&self.inner).keys().cloned().collect();
            keys.sort();
            keys
        }
        fn plan(&self, tab: &str, saved: &[external_links::ExternalLinkInfo], target: &str) -> LinkSavePlan {
            save_plan(&lock_recover(&self.inner)[tab], saved, target)
        }
    }

    fn delta(others: &[TabDeltas], tab: &str, row: i32, column: i32) -> Option<serde_json::Value> {
        others
            .iter()
            .find(|t| t.tab == tab)?
            .changed
            .iter()
            .find(|d| d.sheet == 0 && d.row == row && d.column == column)
            .map(|d| d.value.clone())
    }

    /// What a save sends to the patcher: the plan's cells on one sheet and
    /// its external links.
    fn patch_json(sheet: &str, plan: &LinkSavePlan) -> String {
        let cells: Vec<serde_json::Value> =
            plan.cells.iter().map(|c| json!({ "r": c.r, "c": c.c, "f": c.f, "v": c.v })).collect();
        json!({ "version": 1, "sheets": [{ "name": sheet, "cells": cells }], "externalLinks": plan.external_links })
            .to_string()
    }

    #[test]
    fn links_are_found_in_formula_text() {
        let found = find_links("SUM('[NOW.xlsx]Model'!A1:B2)+[1]Sheet1!A1");
        assert_eq!(found.len(), 2);
        assert_eq!((found[0].book.as_str(), found[0].sheet.as_str(), found[0].dir.as_str()), ("NOW.xlsx", "Model", ""));
        assert_eq!(&"SUM('[NOW.xlsx]Model'!A1:B2)+[1]Sheet1!A1"[found[0].start..found[0].end], "'[NOW.xlsx]Model'!");
        assert_eq!((found[1].book.as_str(), found[1].sheet.as_str()), ("1", "Sheet1"));

        let found = find_links(r"'C:\models\[Q3 plan.xlsx]Income ''24'!B2");
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].dir, r"C:\models\");
        assert_eq!(found[0].book, "Q3 plan.xlsx");
        assert_eq!(found[0].sheet, "Income '24");

        // Not links: table columns, text, an ordinary quoted sheet.
        assert!(find_links("SUM(Sales[Amount])+Sales[[#Headers],[Amount]]").is_empty());
        assert!(find_links(r#"IF(A1="[x]y!","[1]a!",'My Sheet'!A1)"#).is_empty());
        assert!(find_links("[@Amount]*2").is_empty());
        // Letters outside ASCII are part of a sheet name.
        assert_eq!(find_links("[Übersicht.xlsx]Zahlen!A1")[0].sheet, "Zahlen");

        assert_eq!(
            rewrite_links("[A.xlsx]S!A1+'[B.xlsx]My Sheet'!A1+[2]T!A1", |l| {
                if is_number(&l.book) { None } else { Some(numbered_prefix(7, &l.sheet)) }
            }),
            "[7]S!A1+'[7]My Sheet'!A1+[2]T!A1"
        );
        assert_eq!(numbered_prefix(1, "Model"), "[1]Model!");
        assert_eq!(numbered_prefix(2, "Income Statement"), "'[2]Income Statement'!");
        assert_eq!(quoted_prefix("[A.xlsx]It's"), "'[A.xlsx]It''s'!");
        assert_eq!(split_link_sheet_name("[NOW.xlsx]Model"), Some(("NOW.xlsx", "Model")));
        assert_eq!(split_link_sheet_name("[]Model"), None);
        assert_eq!(split_link_sheet_name("Model"), None);
    }

    #[test]
    fn a_saved_link_names_its_file_and_the_file_is_found_again() {
        assert_eq!(link_target("/a/b/sub/NOW model.xlsx", "/a/b"), "sub/NOW%20model.xlsx");
        assert_eq!(link_target("/x/y z.xlsx", "/a/b"), "file:///x/y%20z.xlsx");
        assert_eq!(link_target(r"C:\m\a#1.xlsx", r"D:\x"), "file:///C:/m/a%231.xlsx");

        let dir = folder("find-saved");
        let a = book(&dir, "A model.xlsx", &["Sheet1"], &[]);
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        assert_eq!(find_saved(&r, "A%20model.xlsx"), Some(a.clone()));
        assert_eq!(find_saved(&r, "a MODEL.XLSX"), Some(a.clone()));
        // Saved somewhere else and moved here with its reader.
        assert_eq!(find_saved(&r, "file:///C:/Users/someone/models/A%20model.xlsx"), Some(a.clone()));
        assert_eq!(find_saved(&r, "file:///nowhere/at/all/A%20model.xlsx"), Some(a));
        assert_eq!(find_saved(&r, "R.xlsx"), None);
        assert_eq!(find_saved(&r, "B.xlsx"), None);
        assert_eq!(find_saved(&r, "notes.txt"), None);
    }

    #[test]
    fn a_typed_link_is_found_loaded_and_evaluated() {
        let dir = folder("typed");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[(0, 1, 1, "10"), (0, 1, 2, "=A1*2")]);
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("r", &r);

        let res = ws.edit("r", 0, 1, 1, "=[A.xlsx]Sheet1!B1+1");
        assert_eq!(ws.value("r", 0, 1, 1), json!(21.0));
        let links = res.links.expect("the edit reports the link");
        assert!(links.changed);
        assert!(links.notes.is_empty(), "{:?}", links.notes);
        assert_eq!(res.sheets, Some(vec!["Sheet1".to_string(), "[A.xlsx]Sheet1".to_string()]));
        assert!(ws.formula("r", 0, 1, 1).contains("[A.xlsx]Sheet1"), "{}", ws.formula("r", 0, 1, 1));
        assert!(res.changed.iter().any(|d| (d.sheet, d.row, d.column) == (0, 1, 1) && d.value == json!(21.0)));
        assert!(res.changed.iter().all(|d| d.sheet == 0), "the hidden copy is not reported to the grid");
        // The workbook read is loaded, without a tab.
        assert_eq!(ws.keys(), vec![bg_key(&a), "r".to_string()]);

        // The same workbook again, named without its extension: no new copy.
        let res = ws.edit("r", 0, 2, 1, "=[a]Sheet1!A1");
        assert_eq!(ws.value("r", 0, 2, 1), json!(10.0));
        assert!(res.links.is_none() && res.sheets.is_none());
        assert_eq!(lock_recover(&ws.inner)["r"].links.len(), 1);

        // What cannot be linked is left as typed, and said; the edit goes through.
        let res = ws.edit("r", 0, 3, 1, "=[Nope.xlsx]Sheet1!A1");
        let links = res.links.expect("a note");
        assert!(!links.changed);
        assert!(links.notes.len() == 1 && links.notes[0].contains("Nope.xlsx"), "{:?}", links.notes);
        let res = ws.edit("r", 0, 4, 1, "=[A.xlsx]Other!A1+[A.xlsx]Other!A2");
        let notes = res.links.expect("a note").notes;
        assert_eq!(notes, vec!["A.xlsx has no sheet called Other".to_string()]);

        // A link to itself is an ordinary reference.
        ws.edit("r", 0, 5, 1, "=[R.xlsx]Sheet1!A2*3");
        assert_eq!(ws.value("r", 0, 5, 1), json!(30.0));
        assert!(!ws.formula("r", 0, 5, 1).contains('['), "{}", ws.formula("r", 0, 5, 1));

        // Both are on the workspace list now, and the list says who reads whom.
        let list = status(&ws.inner, &ws.shared);
        let row = |path: &str| list.iter().find(|m| m.path == path).unwrap();
        assert_eq!((row(&r).state, row(&r).member, row(&r).reads.clone()), ("open", true, vec![a.clone()]));
        assert_eq!((row(&a).state, row(&a).member, row(&a).read_by.clone()), ("loaded", true, vec![r.clone()]));
        assert!(row(&r).broken.is_empty());
    }

    #[test]
    fn a_change_reaches_every_workbook_down_the_chain() {
        let dir = folder("chain");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[(0, 1, 1, "1"), (0, 2, 1, "=A1+100")]);
        let b = book(&dir, "B.xlsx", &["Sheet1"], &[]);
        let c = book(&dir, "C.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("a", &a);
        ws.open("b", &b);
        ws.open("c", &c);
        ws.edit("b", 0, 1, 1, "=[A.xlsx]Sheet1!A2*2");
        ws.edit("c", 0, 1, 1, "=[B.xlsx]Sheet1!A1+1");
        assert_eq!(ws.value("c", 0, 1, 1), json!(203.0));

        let res = ws.edit("a", 0, 1, 1, "5");
        assert_eq!(delta(&res.others, "b", 1, 1), Some(json!(210.0)));
        assert_eq!(delta(&res.others, "c", 1, 1), Some(json!(211.0)));
        assert_eq!(ws.value("c", 0, 1, 1), json!(211.0));

        // A loop is refused: A cannot read C, which reads it through B.
        let res = ws.edit("a", 0, 5, 1, "=[C.xlsx]Sheet1!A1");
        let links = res.links.expect("a note");
        assert!(!links.changed && links.notes[0].contains("reads this workbook"), "{:?}", links.notes);
        assert!(lock_recover(&ws.inner)["a"].links.is_empty());
    }

    #[test]
    fn rows_inserted_in_a_workbook_are_followed_by_its_readers() {
        let dir = folder("rows");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[(0, 3, 1, "7"), (0, 4, 1, "1")]);
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        let c = book(&dir, "C.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("a", &a);
        ws.open("r", &r);
        ws.open("c", &c);
        ws.edit("r", 0, 1, 1, "=[A.xlsx]Sheet1!A3");
        ws.edit("r", 0, 2, 1, "=SUM([A.xlsx]Sheet1!A1:A4)");
        ws.edit("c", 0, 1, 1, "=[R.xlsx]Sheet1!A2");
        assert_eq!(ws.value("c", 0, 1, 1), json!(8.0));

        // R's tab is closed with nothing unsaved; C still reads it.
        close_tab(&ws.inner, &ws.shared, "r", false);
        let held = bg_key(&r);
        assert_eq!(ws.keys(), vec!["a".to_string(), held.clone(), "c".to_string()]);

        let res = structural(&ws.inner, "a", "insert_rows", 0, 1, 2).unwrap();
        assert!(res.moved.is_empty(), "R has no tab");
        assert!(ws.formula(&held, 0, 1, 1).contains("A5"), "{}", ws.formula(&held, 0, 1, 1));
        assert!(ws.formula(&held, 0, 2, 1).contains("A3:A6"), "{}", ws.formula(&held, 0, 2, 1));
        assert_eq!(ws.value(&held, 0, 1, 1), json!(7.0));
        assert_eq!(ws.value(&held, 0, 2, 1), json!(8.0));
        assert_eq!(ws.value("c", 0, 1, 1), json!(8.0));
        {
            let map = lock_recover(&ws.inner);
            assert!(map[&held].needs_saving && map[&held].rewrite_links);
        }

        // A value typed into the moved cell arrives under the new address.
        let res = ws.edit("a", 0, 5, 1, "70");
        assert_eq!(ws.value(&held, 0, 1, 1), json!(70.0));
        assert_eq!(delta(&res.others, "c", 1, 1), Some(json!(71.0)));

        // Opened again, R's model is the one that followed, and it says it
        // has something to save.
        let opened = ws.open("r", &r);
        assert!(opened.needs_saving);
        assert_eq!(ws.keys(), vec!["a".to_string(), "c".to_string(), "r".to_string()]);
        assert!(ws.plan("r", &[], &r).rewrite_all);

        // With a tab, the reader is named so the app can mark it changed.
        let res = structural(&ws.inner, "a", "delete_rows", 0, 1, 2).unwrap();
        assert_eq!(res.moved, vec!["r".to_string()]);
        assert!(ws.formula("r", 0, 1, 1).contains("A3"), "{}", ws.formula("r", 0, 1, 1));
        assert_eq!(delta(&res.others, "r", 1, 1), Some(json!(70.0)));
        assert_eq!(ws.value("c", 0, 1, 1), json!(71.0));
    }

    #[test]
    fn a_renamed_sheet_is_renamed_in_the_links() {
        let dir = folder("rename");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[(0, 1, 1, "4")]);
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("a", &a);
        ws.open("r", &r);
        ws.edit("r", 0, 1, 1, "=[A.xlsx]Sheet1!A1");

        let res = sheet_op(&ws.inner, "a", "rename", "Sheet1", Some("Data 2")).unwrap();
        assert_eq!(res.moved, vec!["r".to_string()]);
        assert!(ws.formula("r", 0, 1, 1).contains("[A.xlsx]Data 2"), "{}", ws.formula("r", 0, 1, 1));
        assert_eq!(ws.value("r", 0, 1, 1), json!(4.0));
        {
            let map = lock_recover(&ws.inner);
            assert_eq!(map["r"].links, vec![LinkEntry { name: "[A.xlsx]Data 2".into(), source: a.clone(), sheet: "Data 2".into() }]);
            assert_eq!(map["r"].sheets, vec!["Sheet1".to_string(), "[A.xlsx]Data 2".to_string()]);
            assert!(map["r"].needs_saving && map["r"].rewrite_links);
        }
        let res = ws.edit("a", 0, 1, 1, "9");
        assert_eq!(delta(&res.others, "r", 1, 1), Some(json!(9.0)));

        // A deleted sheet keeps its last values in the reader, and the list
        // shows the link as broken.
        sheet_op(&ws.inner, "a", "add", "Other", None).unwrap();
        sheet_op(&ws.inner, "a", "delete", "Data 2", None).unwrap();
        assert_eq!(ws.value("r", 0, 1, 1), json!(9.0));
        let list = status(&ws.inner, &ws.shared);
        assert_eq!(list.iter().find(|m| m.path == r).unwrap().broken, vec!["[A.xlsx]Data 2".to_string()]);
    }

    #[test]
    fn links_are_saved_into_the_file_and_attached_when_it_is_opened() {
        let dir = folder("save");
        let a = book(
            &dir,
            "A.xlsx",
            &["Sheet1", "Income Statement"],
            &[(0, 1, 1, "10"), (0, 1, 2, "=A1*2"), (1, 2, 2, "300")],
        );
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[(0, 9, 1, "1")]);
        let ws = Ws::new();
        ws.open("r", &r);
        assert!(ws.plan("r", &[], &r).cells.is_empty(), "nothing to add for a workbook without links");
        ws.edit("r", 0, 1, 1, "=[A.xlsx]Sheet1!B1+1");
        ws.edit("r", 0, 2, 1, "='[A.xlsx]Income Statement'!B2");

        let plan = ws.plan("r", &[], &r);
        let texts: Vec<&str> = plan.cells.iter().map(|c| c.f.as_str()).collect();
        assert_eq!(texts, vec!["[1]Sheet1!B1+1", "'[1]Income Statement'!B2"]);
        assert_eq!(plan.cells[0].v, Some(json!({ "t": "n", "n": 21.0 })));
        assert_eq!((plan.cells[1].sheet.as_str(), plan.cells[1].r, plan.cells[1].c), ("Sheet1", 1, 0));
        assert_eq!(plan.external_links.len(), 1);
        assert_eq!(plan.external_links[0].index, 1);
        assert_eq!(plan.external_links[0].target.as_deref(), Some("A.xlsx"));
        let mut sheets: Vec<&str> = plan.external_links[0].sheets.iter().map(|s| s.name.as_str()).collect();
        sheets.sort();
        assert_eq!(sheets, vec!["Income Statement", "Sheet1"]);
        assert!(!plan.rewrite_all);
        assert_eq!(plan.reads, vec![a.clone()]);
        // Saved into another folder, the workbook is named by its full address.
        let away = ws.plan("r", &[], "/somewhere/else/R.xlsx").external_links[0].target.clone().unwrap();
        assert!(away.starts_with("file:///") && away.ends_with("/A.xlsx"), "{away}");

        let out = xlsx_patch::apply_patch_json(&std::fs::read(&r).unwrap(), &patch_json("Sheet1", &plan)).unwrap();
        std::fs::write(&r, out).unwrap();
        let saved = external_links::read(&std::fs::read(&r).unwrap()).unwrap();
        assert_eq!(saved.len(), 1);
        assert_eq!(saved[0].target.as_deref(), Some("A.xlsx"));

        // Another run of the app: the file's links are attached again and
        // the workbook it reads is loaded with it.
        let ws = Ws::new();
        ws.open("r", &r);
        assert_eq!(ws.keys(), vec![bg_key(&a), "r".to_string()]);
        assert_eq!(ws.value("r", 0, 1, 1), json!(21.0));
        assert_eq!(ws.value("r", 0, 2, 1), json!(300.0));
        assert!(ws.formula("r", 0, 1, 1).contains("[A.xlsx]Sheet1"), "{}", ws.formula("r", 0, 1, 1));
        assert_eq!(lock_recover(&ws.inner)["r"].links.len(), 2);
        // Saved again, it keeps the number the file has for the workbook.
        let again = ws.plan("r", &saved, &r);
        assert_eq!(again.cells.iter().map(|c| c.f.as_str()).collect::<Vec<_>>(), texts);
        assert_eq!((again.external_links[0].index, again.external_links[0].target.clone()), (1, None));
        // A second workbook read gets the next number.
        let b = book(&dir, "B.xlsx", &["Sheet1"], &[(0, 1, 1, "5")]);
        ws.edit("r", 0, 3, 1, "=[B.xlsx]Sheet1!A1");
        let two = ws.plan("r", &saved, &r);
        assert_eq!(two.cells[2].f, "[2]Sheet1!A1");
        assert_eq!((two.external_links[1].index, two.external_links[1].target.as_deref()), (2, Some("B.xlsx")));
        assert_eq!(two.reads, vec![a.clone(), b]);

        // Closed, nothing stays loaded.
        close_tab(&ws.inner, &ws.shared, "r", true);
        assert!(ws.keys().is_empty());
    }

    #[test]
    fn a_link_typed_in_a_workbook_without_a_file_says_why_it_is_left() {
        let ws = Ws::new();
        let mut model = Model::new_empty("untitled", "en", "UTC", "en").unwrap();
        model.evaluate();
        lock_recover(&ws.inner).insert("draft".to_string(), Loaded::new(model, String::new(), false));

        let typed = ws.edit("draft", 0, 1, 1, "=[NOW.xlsx]Model!B2");
        let notes = typed.links.expect("the edit reports on its link").notes;
        assert_eq!(notes.len(), 1);
        assert!(notes[0].contains("has no file yet"), "{notes:?}");
        assert_eq!(ws.formula("draft", 0, 1, 1), "=[NOW.xlsx]Model!B2");
        assert_eq!(ws.keys(), vec!["draft"]);

        // Brackets that are not a link, and plain edits, say nothing.
        assert!(ws.edit("draft", 0, 2, 1, "=SUM(1,2)").links.is_none());
        assert!(ws.edit("draft", 0, 3, 1, "=[1]Sheet1!A1").links.is_none());
    }

    #[test]
    fn a_link_to_a_file_that_is_not_there_is_saved_by_name() {
        let dir = folder("unresolved");
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("r", &r);
        ws.edit("r", 0, 1, 1, "=[Not here.xlsx]Sheet1!A1+1");
        let plan = ws.plan("r", &[], &r);
        assert_eq!(plan.cells[0].f, "[1]Sheet1!A1+1");
        assert_eq!(plan.external_links.len(), 1);
        assert_eq!(plan.external_links[0].target.as_deref(), Some("Not%20here.xlsx"));
        assert_eq!(plan.external_links[0].sheets.len(), 1);
        assert!(plan.external_links[0].sheets[0].cells.is_empty());

        let out = xlsx_patch::apply_patch_json(&std::fs::read(&r).unwrap(), &patch_json("Sheet1", &plan)).unwrap();
        std::fs::write(&r, out).unwrap();
        let saved = external_links::read(&std::fs::read(&r).unwrap()).unwrap();
        assert_eq!(saved[0].target.as_deref(), Some("Not%20here.xlsx"));

        // Opened again the link stays as the file has it, and a save leaves
        // it alone.
        let ws = Ws::new();
        ws.open("r", &r);
        assert_eq!(ws.keys(), vec!["r".to_string()]);
        assert!(ws.formula("r", 0, 1, 1).contains("[1]Sheet1"), "{}", ws.formula("r", 0, 1, 1));
        let again = ws.plan("r", &saved, &r);
        assert_eq!(again.cells[0].f, "[1]Sheet1!A1+1");
        assert!(again.external_links.is_empty());

        // The file turns up: the link is attached at the next open.
        book(&dir, "Not here.xlsx", &["Sheet1"], &[(0, 1, 1, "41")]);
        let ws = Ws::new();
        ws.open("r", &r);
        assert_eq!(ws.value("r", 0, 1, 1), json!(42.0));
        assert!(ws.formula("r", 0, 1, 1).contains("[Not here.xlsx]Sheet1"), "{}", ws.formula("r", 0, 1, 1));
    }

    #[test]
    fn a_workbook_that_is_read_outlives_its_tab() {
        let dir = folder("close");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[(0, 1, 1, "1")]);
        let r = book(&dir, "R.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("a", &a);
        ws.open("r", &r);
        ws.edit("r", 0, 1, 1, "=[A.xlsx]Sheet1!A1");
        ws.edit("a", 0, 1, 1, "99");
        assert_eq!(ws.value("r", 0, 1, 1), json!(99.0));

        // Closed without saving: its readers go back to what the file has.
        let others = close_tab(&ws.inner, &ws.shared, "a", true);
        assert_eq!(delta(&others, "r", 1, 1), Some(json!(1.0)));
        assert_eq!(ws.keys(), vec![bg_key(&a), "r".to_string()]);

        // Opened again, the model held for it is taken over.
        let seq = lock_recover(&ws.inner)[&bg_key(&a)].seq;
        ws.open("a", &a);
        assert_eq!(ws.keys(), vec!["a".to_string(), "r".to_string()]);
        assert!(lock_recover(&ws.inner)["a"].seq > seq);
        let res = ws.edit("a", 0, 1, 1, "50");
        assert_eq!(delta(&res.others, "r", 1, 1), Some(json!(50.0)));

        // Closed with nothing unsaved: kept as it is.
        assert!(close_tab(&ws.inner, &ws.shared, "a", false).is_empty());
        assert_eq!(ws.keys(), vec![bg_key(&a), "r".to_string()]);
        assert_eq!(ws.value("r", 0, 1, 1), json!(50.0));

        // Saved under another name: the workspace list follows.
        let r2 = dir.join("R2.xlsx").to_string_lossy().to_string();
        std::fs::copy(&r, &r2).unwrap();
        mark_saved(&ws.inner, &ws.shared, "r", &r2);
        assert_eq!(lock_recover(&ws.inner)["r"].path, r2);
        assert!(lock_shared(&ws.shared).members.contains(&r2));

        close_tab(&ws.inner, &ws.shared, "r", false);
        assert!(ws.keys().is_empty());
    }

    #[test]
    fn the_agent_reads_a_workspace_workbook_that_has_no_tab() {
        let dir = folder("agent-read");
        let snow = book(&dir, "SNOW.xlsx", &["Model"], &[(0, 2, 2, "3400"), (0, 3, 2, "=B2*2")]);
        let snps = book(&dir, "SNPS.xlsx", &["Model"], &[(0, 1, 1, "1")]);
        let other = book(&dir, "Other.xlsx", &["Sheet1"], &[]);
        let ws = Ws::new();
        ws.open("t-snps", &snps);

        // Not on the list and not loaded: the agent cannot make the app open it.
        let refused = model_for_reading(&ws.inner, &ws.shared, &other).unwrap_err();
        assert!(refused.contains("not in the workspace"), "{refused}");
        assert!(model_for_reading(&ws.inner, &ws.shared, "SNOW.xlsx").is_err());
        assert!(model_for_reading(&ws.inner, &ws.shared, &dir.join("notes.txt").to_string_lossy()).is_err());
        assert_eq!(ws.keys(), vec!["t-snps"]);

        // On the list: loaded without a tab, evaluated, and found under its key.
        lock_shared(&ws.shared).members.push(snow.clone());
        let model = model_for_reading(&ws.inner, &ws.shared, &snow).unwrap();
        assert_eq!(model.key, bg_key(&snow));
        assert_eq!(model.sheets.iter().map(|s| (s.name.as_str(), s.link)).collect::<Vec<_>>(), vec![("Model", false)]);
        assert_eq!(ws.value(&model.key, 0, 3, 2), json!(6800.0));
        // Asked again, it is the same model.
        assert_eq!(model_for_reading(&ws.inner, &ws.shared, &snow).unwrap().key, model.key);

        // The workbook in a tab is read as the tab has it, unsaved edits and all,
        // and a link the agent then writes shows up as a hidden copy, flagged.
        ws.edit("t-snps", 0, 1, 1, "42");
        ws.edit("t-snps", 0, 2, 1, "='[SNOW.xlsx]Model'!B3");
        let own = model_for_reading(&ws.inner, &ws.shared, &snps).unwrap();
        assert_eq!(own.key, "t-snps");
        assert_eq!(ws.value("t-snps", 0, 1, 1), json!(42.0));
        assert_eq!(ws.value("t-snps", 0, 2, 1), json!(6800.0));
        assert_eq!(own.sheets.iter().map(|s| (s.name.as_str(), s.link)).collect::<Vec<_>>(), vec![("Model", false), ("[SNOW.xlsx]Model", true)]);
    }

    #[test]
    fn dropped_folders_give_their_workbooks_and_a_new_workbook_opens() {
        let dir = folder("drop");
        let a = book(&dir, "A.xlsx", &["Sheet1"], &[]);
        let b = book(&dir, "B.xlsx", &["Sheet1"], &[]);
        std::fs::write(dir.join("~$A.xlsx"), b"lock").unwrap();
        std::fs::write(dir.join("notes.txt"), b"x").unwrap();
        std::fs::create_dir_all(dir.join("deeper")).unwrap();
        book(&dir.join("deeper"), "C.xlsx", &["Sheet1"], &[]);
        let dropped = vec![dir.to_string_lossy().to_string(), a.clone(), dir.join("notes.txt").to_string_lossy().to_string()];
        assert_eq!(expand_paths(&dropped), vec![a, b]);

        let made = dir.join("Comps.xlsx").to_string_lossy().to_string();
        std::fs::write(&made, create_workbook_bytes(&["Comps".to_string(), "Notes".to_string()]).unwrap()).unwrap();
        let ws = Ws::new();
        assert_eq!(ws.open("t", &made).sheets, vec!["Comps".to_string(), "Notes".to_string()]);
    }
}
