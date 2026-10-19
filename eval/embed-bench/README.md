# Embed benchmark

GridPath (`gridpath`) against the embeddable spreadsheets a developer
would otherwise pick, on real workbooks, in headless Chromium with production
builds.

| Product | Setup |
|---|---|
| GridPath | `createSpreadsheet` from the packed `packages/gridpath` tarball |
| SpreadJS 19.2 | `spread-sheets` + `-io`, `-charts`, `-shapes`, `-pivot-addon`, `-slicers`; evaluation mode (no key). Also run with `openMode: lazy` |
| Handsontable 18.1 | No xlsx support of its own: ExcelJS 4.4 reads and writes, HyperFormula 3.5 calculates (its Excel-compatibility settings), Handsontable displays. The glue is `src/hot.ts` |

Univer is not included: its xlsx import/export is a paid feature that needs a
conversion server.

## Results, 2026-10-19

Apple M2 Pro, 16 GB, macOS; Chromium 151 (Playwright 1.62.1, headless);
Node 22; production builds; one run per cell, edits the median of five.
Memory is the resident memory of the browser's whole process tree after
open; on macOS it moves by about 20% between runs (memory compression), so
read it as a range. SpreadJS ran in evaluation mode (no license key).

The two real models are private and not published. What they are:

- **Small fixture**: `eval/fixtures/rich-model.xlsx` (8 KB, in the repo): two
  sheets, a chart, a drawing, custom XML, printer settings.
- **Analyst model**: a 1.4 MB sell-side equity model; 12 sheets, 55,000
  formulas, 3,418 defined names, INDIRECT-heavy summary sheets and about
  1,000 data-provider add-in formulas.
- **Data-intensive model**: over 10 MB; 2.3 million cells and 930,000
  formulas, most of them SUMIFS over 7,000-row columns.

### Analyst model (1.4 MB, 55,000 formulas)

| | GridPath 0.2.0 | SpreadJS 19.2 | Handsontable 18.1 + HyperFormula 3.5 + ExcelJS 4.4 |
|---|---|---|---|
| Open → grid usable | 0.6 s | 1.5 s | 1.9 s |
| Open → fully recalculated | 0.8 s | 3.6 s | 1.9 s |
| Memory after open | 348 MB | 448 MB | 447 MB |
| Edit an input that moves 2,256 formulas → dependent updated | 99 ms | 164 ms | never updates |
| Agrees with Excel's own results | 95.5% | 94.0% | 56.9% |
| Saved file: parts byte-identical | 75 of 76 | 3 of 76 | 2 of 76 |
| Lost on save | nothing | custom XML (6 parts), printer settings (6) | custom XML, printer settings, 3 of 8 comment parts, 3,335 of 3,418 defined names |

### Data-intensive model (over 10 MB, 2.3 million cells)

| | GridPath 0.2.0 | SpreadJS 19.2 | SpreadJS 19.2, lazy open | Handsontable stack |
|---|---|---|---|---|
| Open → grid usable | 4.2 s | tab crashed (at 3.5 GB) | 6.5 s | 10.8 min |
| Open → fully recalculated | 10.9 s | — | 3.7 min | 10.8 min |
| Memory after open | 0.8–1.0 GB | — | 2.7 GB | 2.4 GB |
| Edit an input that moves 919 formulas → dependent updated | 2.7 s | — | tab crashed | 64 s |
| Agrees with Excel's own results | 99.97% | — | — | 68.7% |
| Saved file: parts byte-identical | 38 of 43 | — | — | 1 of 43 |
| Lost on save | nothing | — | — | nothing |

### Small fixture (chart, drawing, custom XML)

| | GridPath 0.2.0 | SpreadJS 19.2 | Handsontable stack |
|---|---|---|---|
| Saved file: parts byte-identical | 15 of 16 | 0 of 16 | could not open the file (ExcelJS fails on the chart) |
| Lost on save | nothing | custom XML, printer settings (the chart is kept, rewritten) | — |
| Edit → dependent updated | 11 ms | 0.4 ms | — |

Notes:

- Where GridPath and SpreadJS disagree with Excel it is mostly on the same
  cells: formulas that depend on today's date, and add-in calls no engine can
  evaluate.
- SpreadJS's crashes are Chrome's per-tab JavaScript memory limit (about
  4 GB), which every user's browser has.
- Handsontable's edit never registers on the analyst model because the
  watched cell evaluates to empty in HyperFormula: the model's INDIRECT chains
  and the defined names ExcelJS does not read.
- Where GridPath is slower: a single edit on a tiny workbook costs about
  11 ms (a round trip to the engine worker) against SpreadJS's 0.4 ms, and the
  heaviest edit measured on the data-intensive model takes 2.7 s.

## Measured

- **Open → grid usable / fully recalculated.** From the file's bytes in memory.
  SpreadJS keeps the file's cached results on import, so it is then made to
  recalculate everything (`calculate(rebuild)`), as the others do on open.
- **Memory.** RSS of the whole browser process tree (page, workers, GPU),
  minus an idle tab; after open and GC, and the peak.
- **Edit → dependent updated.** `prep.mjs` picks a numeric input that moves
  many formulas and one dependent to watch (on another sheet when possible);
  median of 5 alternating edits, until the dependent reads back changed.
- **Agrees with Excel.** Every formula cell that has a cached result in the
  file, compared with the product's value after a full recalculation
  (relative tolerance 1e-9; errors match any error). GridPath is scored in
  Node with the same engine build. Volatile functions (`TODAY()`) and add-in
  calls count against every product equally.
- **Save.** Part-by-part comparison with the original package, losses by
  feature kind (products that renumber parts are not penalised for the
  rename), defined names, and formulas in the saved file reopened in the
  GridPath engine.

## Run

```sh
crates/engine-wasm/build.sh && (cd packages/gridpath && npm install && npm pack)
cd eval/embed-bench && npm install
node --max-old-space-size=14000 prep.mjs      # edit targets + Excel's cached results
npx vite build
node --max-old-space-size=12000 run.mjs       # or product:file ...
node summary.mjs                              # results/summary.md
```

Workbooks are listed in `files.mjs`. The two real models are private files
given by `ANALYST_MODEL=` and `LARGE_MODEL=` (paths to your own copies; a
workbook whose variable is unset is skipped); they and
everything under `results/` stay out of git.

One run per cell on one machine (Apple-silicon Mac); treat differences under
~20% as noise.
