# Surgical Save: design plan

**Goal:** a user opens a real .xlsx (bank model, FactSet-linked, charts, macros-adjacent parts),
edits cells in GridPath, saves — and the file opens clean in Excel with everything we didn't
touch preserved **byte-identical**.

**Why:** the current save path exports the whole workbook through ExcelJS, which re-serializes
the entire package from its object model. Measured on a real GS model, a zero-edit
open→save drops: `xl/externalLinks/*` (breaks `[n]` formulas → Excel repair dialog),
`xl/webextensions/*` (FactSet task pane), `customXml/*`, `docProps/custom.xml`,
11 `customProperty*.bin`, both comment parts + VML anchors, one drawing part, and all
printer settings. 918KB → 586KB. Preservation can never be retrofitted into ExcelJS —
it only writes what it models.

**Principle:** the original file bytes are the source of truth. Save = apply a minimal
patch to the original zip package in Rust. Every zip entry we don't have a reason to
touch is raw-copied, exact bytes, original compression.

---

## Architecture

```
Frontend (unchanged roles)                Rust (new)
─────────────────────────                 ──────────────────────────────
ExcelJS  = read/parse only                save_workbook_patched(path, baseHash, patchJson)
Univer   = view + edit model                └─ engine/workbook/xlsx_patch/
SaveMirror = op log (formats,                  reads base bytes → applies patch →
  merges, structure ops…)                      atomic write (existing temp+rename)
+ NEW: cell diff at save time
```

- **Open** (unchanged): `read_workbook_file` → frontend parses with ExcelJS → Univer.
  Frontend additionally keeps the **load-time canonical cell model** (it already builds
  `cellData` during import — retain a frozen copy) and a **hash of the base bytes**.
- **Save**: frontend computes a *cell diff* (Univer snapshot vs load-time copy), folds in
  the existing `SaveMirror` ops, sends one patch JSON to Rust. Rust re-reads the file at
  `path`, verifies the hash (detects out-of-band modification → surface to user), applies
  the patch, atomic-writes. After save, the new bytes/hash become the base.
- **Untitled / brand-new workbooks**: no base to preserve — keep the current ExcelJS
  fresh-export path. Nothing to butcher.
- **Failure escape hatch**: if patch application fails, fall back to the ExcelJS full
  export **with a visible warning** ("saved with reduced fidelity") — losing fidelity
  beats losing the user's edits. Log the patch + error for repro.

### Cell diff rules (frontend)

Diff `(f, v)` per cell against the load-time copy:

- Cell **changed** iff its formula differs, or it has no formula and its value differs.
- A cell whose `f` is unchanged but whose `v` (Univer recalc result) drifted is **not**
  a change — we never trust Univer's numbers in the file; Excel recalcs (see
  `fullCalcOnLoad` below).
- External-pinned cells (cached-value pinning from `UniverGrid.tsx`) fall out naturally:
  no `f`, `v` equals the pinned cache → not in the diff → sheet XML untouched →
  `xl/externalLinks/*` survives because it is never re-serialized. The WeakMap
  skip-on-save logic can be deleted once the ExcelJS write path retires.

Diff-at-save (rather than instrumenting every edit path) catches typing, paste, fills,
and agent batches uniformly, and costs one pass over sheets that Univer marks dirty.

---

## Patch format (versioned JSON)

```jsonc
{
  "version": 1,
  "baseHash": "xxh64:9f2c…",         // bytes delivered at open / last save
  "fullCalcOnLoad": true,             // set once if any formula/value changed
  "sheets": [{
    "name": "NTM Ests.",
    "cells": [
      { "r": 4, "c": 2, "v": { "t": "n", "n": 123.4 } },       // number
      { "r": 4, "c": 3, "v": { "t": "s", "s": "1Q27E" } },     // string
      { "r": 4, "c": 4, "v": { "t": "b", "b": true } },        // bool
      { "r": 5, "c": 2, "f": "SUM(A1:A3)" },                   // formula, no cached v
      { "r": 6, "c": 2, "clear": "contents" }                  // keep style
    ],
    "colWidths":  [{ "c": 3, "chars": 12.5 }],
    "rowHeights": [{ "r": 7, "pts": 18 }],
    "merges":     [{ "range": "B2:D2", "merge": true }],
    "freeze":     { "rows": 1, "cols": 2 },
    "autoFilter": "A1:F40",
    "rowColOps":  []                   // phase 3
  }],
  "styles":       [],                  // phase 2: {sheet, r, c, format{…}}
  "definedNames": [{ "name": "WACC", "ref": "'DCF'!$C$4" }],
  "sheetOps":     []                   // phase 3: create/rename/delete
}
```

Everything except `cells` maps 1:1 from today's `SaveMirror`; `cells` comes from the diff.

---

## Rust module layout

```
src-tauri/src/engine/workbook/xlsx_patch/
├── mod.rs             pub fn apply_patch(base: &[u8], patch: &Patch) -> Result<Vec<u8>>
├── patch.rs           serde types (versioned; reject unknown version)
├── package.rs         zip walk; raw-copy untouched entries; part add/remove;
│                      [Content_Types].xml and .rels edits
├── sheet_xml.rs       streaming worksheet rewriter (quick-xml event copy + splice)
├── shared_strings.rs  phase 2: append-only sharedStrings (phase 1 uses inlineStr)
├── styles.rs          phase 2: append-only interning of xf/font/fill/border/numFmt
├── workbook_xml.rs    calcPr fullCalcOnLoad, definedNames; sheetOps (phase 3)
├── refs.rs            A1↔(r,c), range parse, formula ref shifting (phases 1*/3)
└── tests/             corpus harness + unit tests
```

New crates: `zip` (use `raw_copy_file` so untouched entries keep exact bytes and
compression), `quick-xml`, `xxhash-rust`. New Tauri command
`save_workbook_patched(path, base_hash, patch_json)` alongside the existing
`write_workbook_file` (which remains for the fallback and new-file paths).

\* `refs.rs` is needed in phase 1 only for shared-formula materialization (below).

---

## Which parts get touched, per edit type

| Edit | Parts rewritten | Notes |
|---|---|---|
| — nothing (empty patch) | **none — output byte-identical to input** | the core invariant |
| cell value/formula | that `sheetN.xml`; `calcChain.xml` **deleted** (+ its Content-Types override); `workbook.xml` gets `<calcPr fullCalcOnLoad="1"/>` once | Excel rebuilds calcChain silently and recalcs on open |
| new string value | none beyond the sheet — phase 1 writes `t="inlineStr"` | avoids touching `sharedStrings.xml` entirely; Excel normalizes on its next save. Phase 2: append-only sharedStrings if inlineStr proves problematic |
| formula write | sheet only; emit `<f>…</f>` with **no `<v>`** | never persist Univer's computed result |
| cell format | sheet (`s=` attr) + `styles.xml` (append-only: intern new xf/font/fill/border, never renumber existing) | phase 2 |
| col width / row height / merge / freeze / autofilter | that sheet's `cols`/`row`/`mergeCells`/`sheetViews`/`autoFilter` elements | phase 2 |
| defined names | `workbook.xml` | phase 2 |
| sheet create | new `sheetN.xml` + `workbook.xml` + `workbook.xml.rels` + `[Content_Types].xml` | phase 3 |
| sheet rename | `workbook.xml` + **every** sheet's formulas referencing the old name + definedNames | phase 3 |
| sheet delete | remove part + rels + content-type + workbook entry; leave shared media alone | phase 3; refs to it become `#REF!` (matches Excel) |
| row/col insert/delete | that sheet's rows/cells shifted **and formula references adjusted across all sheets** (plus merges, conditional-formatting ranges, dataValidation, hyperlinks, autoFilter, definedNames touching the range) | phase 3 — the genuinely hard one; see below |

### sheet_xml.rs mechanics

Stream the original sheet XML event-by-event (quick-xml), copying raw events verbatim,
splicing only inside `<sheetData>`:

- Rows sorted by `r`; merge patched cells into existing `<row>` elements or emit new rows
  in order. On a touched cell, preserve its existing `s=` (style) attribute unless a style
  patch targets it; on untouched cells/rows, bytes pass through unmodified.
- **Shared formulas**: if a patched cell is inside a shared-formula group, materialize the
  group first — the master's `<f t="shared" ref si>` and each dependent's bare
  `<f t="shared" si/>` are rewritten as plain translated formulas (relative refs shifted
  by row/col delta — `refs.rs`). This also permanently fixes the current import bug where
  dependents' formulas read as `=<master address>`.
- `dimension` updated only if the used range grew; stale dimensions are tolerated by Excel.
- No pretty-printing, no attribute reordering, no namespace rewriting — raw event copy.

---

## Phases

**Phase 0 — enablers (small):**
frontend keeps load-time cell model + base hash; `xlsx_patch` crate skeleton with
`apply_patch` handling the empty patch; Tauri command; corpus test harness.
Gate: *empty patch → byte-identical file* on the whole corpus.

**Phase 1 — cell values & formulas (the 80% case):**
sheetData splice, inlineStr, formula-without-value, calcChain removal, fullCalcOnLoad,
shared-formula materialization. Behind a setting/flag; ExcelJS export stays as fallback.
Gate: corpus round-trips with cell edits open clean in Excel + LibreOffice; untouched
parts byte-identical.

**Phase 2 — formatting & sheet furniture:**
styles.xml interning (append-only), widths/heights, merges, freeze, autofilter,
definedNames, sharedStrings append if needed. Retire the ExcelJS write path for
existing files; delete the external-pin save-skip logic.

**Phase 3 — structure ops:**
sheet create/rename/delete, then row/col insert/delete with cross-sheet reference
adjustment. Until it ships, a save whose patch contains `rowColOps` uses the ExcelJS
fallback (with warning) rather than writing a file with wrong references. Honest note:
this phase approaches "reimplement a chunk of Excel's ref engine" — sequence it last,
and consider whether the product actually needs in-place row insertion on *preserved*
files before paying for it.

**Phase 4 — cleanup:**
new-file creation moved to a minimal Rust writer or kept on ExcelJS (fine either way);
ExcelJS remains read-only.

---

## Test harness

- `xlsx_patch/tests/corpus/` — synthetic fixtures committed; real proprietary files
  (the GS model etc.) in a **gitignored** local corpus dir, path via env var.
- Invariants per corpus file:
  1. empty patch → output bytes == input bytes;
  2. any patch → every untouched zip entry byte-identical, entry order/compression preserved;
  3. output opens without repair in LibreOffice headless (`soffice --convert-to xlsx` exit
     status as a CI-friendly validity smoke) — Excel spot-checks manually;
  4. reopening the patched file in GridPath yields the expected cell model.
- Unit tests: sheetData splice cases (new row between rows / append past last row / new
  cell mid-row / overwrite / clear), shared-formula translation, inlineStr escaping,
  A1 edge cases (`XFD1048576`, absolute/mixed refs), hash-mismatch rejection.

## Open decisions

- **inlineStr vs sharedStrings append in phase 1** — plan says inlineStr (zero risk to an
  existing part); revisit only if a target tool mishandles it.
- **Where the diff runs** — plan says frontend (it owns both models). Alternative of
  shipping full snapshots to Rust and diffing against a parsed base doubles the parsing
  work in Rust for no fidelity gain.
- **Concurrent modification UX** — on base-hash mismatch, offer "overwrite / save a copy",
  don't silently clobber.
