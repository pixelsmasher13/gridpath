# Local recording demo

From this directory, run `npm run recording`, then open
<http://127.0.0.1:4173/?record=1>. Recording uses a production build of local
GridPath source and its WASM engine, including native line charts. These results
describe the local build, not the published npm package. To rebuild the engine,
first run `bash ../../crates/engine-wasm/build.sh` (Rust and wasm-bindgen required).
`npm run check:recording` checks against the same local package types used by
the recording build; the ordinary demo installs the published package.

Timing reports, save validations and screenshots from recording runs go to
`outputs/` at the repo root, which is not committed: the evidence paths below
are on the recording machine.

## Larkspur Orbital income statement

The active workbook is `public/larkspur-income-statement.xlsx`: 431 formulas,
812 populated cells, 27,642 bytes, three linked sheets and three native line charts.
Larkspur Orbital is fictional (it was SPCX in earlier takes). The 2025–2026 columns are explicitly labeled **modeled base**
(M), not company-reported historical results. The five forecast years (F) are
2027–2031. All operating, financial and share-count assumptions are synthetic.

- **Income statement:** six revenue streams, segment costs, gross profit,
  R&D, SG&A, EBITDA, D&A, EBIT, interest, taxes, net income, margins and diluted
  EPS. Seven annual periods run across the columns. The net-income chart appears
  beside the top of the statement, with revenue/EBITDA and gross/EBITDA margins
  below it. Header rows and line-item labels remain frozen.
- **Operating model:** external commercial/government standard launches,
  heavy-lift missions, pricing and launch costs; subscribers, churn, ARPU,
  terminal sales and network costs; government programs and corporate expenses.
- **Capital & debt:** capital asset and debt roll-forwards, capex, disposals,
  depreciation, average debt interest and assumed cash interest. These feed
  D&A and net interest into the income statement.

The main demo input is **Income statement!F5**, 2027 paid standard launches:
120 → 180. Later standard launches grow from this value using each year's
growth rate on Operating model. Heavy-lift launches are a separate assumption.
Internal satellite deployments are excluded from external launch revenue.
The change raises modeled 2027 revenue from $22.240B to $26.680B and net income
from $3.883B to $5.536B. The two modeled base years remain unchanged.

Other editable forecasts are blue on amber: launch mix and prices, subscriber
additions and churn, ARPU, terminal economics, R&D/SG&A, taxes and diluted shares
on Operating model; capex, debt and interest assumptions on Capital & debt.
The 2027 launch-growth percentage is derived from the main launch-count input;
growth assumptions from 2028 onward drive the forecast directly.

Subscription revenue uses average beginning/ending subscribers × monthly ARPU
× 12. New terminal sales use gross additions. D&A uses opening net assets plus
a half-year charge on capex, capped at assets available after disposals.
Interest uses average debt balances; cash is an explicit interest-income
assumption. Taxes apply to positive pretax income without loss carryforwards.
This is an income statement model with supporting schedules, not a complete
balance sheet or cash-flow model.

## One continuous take

The workbook opens automatically inside a compact Mission Control app header
and a bordered document panel. The default view has a plain document title,
**View chart**, **Download .xlsx**, and **Timings**. The chart shortcut brings the
native net-income chart into view in a narrow pane and toggles back to the statement.
Download saves the current workbook, including
manual edits. Timing collection runs in the background; its panel starts closed.

1. Open **Timings**, begin recording, then click **Reset workbook** for a fresh
   measured load. Reset discards edits. Keep loading and calculation at normal speed.
2. Adjust sheet zoom for the recording viewport; 80% fits the seven statement
   columns in a narrow pane. Use a wider pane to show the charts alongside it.
3. Click **Model 180 paid launches**. The button writes F5 through the public
   component API, checks twelve revenue/EBITDA/net-income results including
   unchanged base years, checks all 35 chart values, and waits two animation frames.
4. Scroll through earnings and explore Operating model / Capital & debt.
5. In **Timings**, click **Save & compare**, then **Export timings** and
   **Download saved copy**. Close the panel to return to the document view.

**Undo scenario** restores and checks the original model and charts. It
clears the previous edit/save timings. **Select launch input** selects F5 for
manual typing in the cell or formula bar. Custom edits clear the scripted
verification and edit/save timings; the model and charts continue calculating,
and saving remains available. Reset the workbook to repeat the scripted scenario.

**Reset workbook** remounts a fresh component and engine worker from the
original workbook. Browser persistence is disabled. Reload after a rebuild to
use current assets; old hashed assets remain available for already-open tabs.
Keep the browser foreground during measurements.

## Measurements and verification — October 22, 2026

In the local in-app Chromium 155 browser, before the document-view interface
cleanup and the net-income chart addition (424 formulas, two charts, original
click-to-open flow):

| Grid ready | Fully calculated, total | Edit → verified statement/charts | Save |
| ---: | ---: | ---: | ---: |
| 0.18 s | 0.23 s | 0.07 s | 0.10 s |

The edit changed 152 cells. Engine evaluation took 3.7 ms, serialization 2.9 ms,
and reading both charts 1.4 ms. Repeating after undo took 0.04 s. These are
single-run observations for this workbook, not a large-file benchmark or a
comparison with another library. Browser HTTP cache is not cleared.

- **Grid ready** now starts when the verified workbook is automatically mounted,
  or when **Reset workbook** remounts it, and ends at component `onReady`, including
  worker/WASM startup and import. File fetch, formula counting and hash checks
  happen before timing begins; this is not page-navigation time. Exported reports
  record whether the open was automatic or a reset. The historical table above
  measured from the original open-button click.
- **Fully calculated** uses the same start through `whenCalculated`; it is total
  elapsed time, not extra time after the grid is ready.
- **Edit** includes the assumption write, calculation, independent result and
  chart verification, and two animation frames. No viewport scroll is timed.
- **Save** is the public `save()` call, excluding ZIP comparison and download.

The saved file retains all 15 original ZIP parts, ten byte-identical. Independent
checks verified every one of the 424 formula caches and unchanged formula text,
all 28 chart values/categories, well-formed XML, and chart XML unchanged outside
its caches. Reopening in the WASM engine verified every formula and chart again.
Additional checks covered edit/restore/re-edit, zero standard launches, and an
ARPU edit flowing through revenue, net income and charts. No cached fallback
formulas were used in the measured browser run.

The hash and scenario checks are in `public/larkspur-income-statement.json`.
Raw evidence is under `outputs/reddit-demo-20261022/`: `spcx-income-timings.json`,
`spcx-income-save-validation.json`, and the original and saved income statement
workbooks. Prior compact-calculator reports (`spcx-*-timings.json`) describe the
29-formula `larkspur-launch-model.xlsx`, which remains available but is no longer active.

### Net-income chart addition — October 22, 2026

The active revision adds a native net-income line chart linked to the existing
statement. Browser checks verified all 35 chart values after the 120 → 180 launch
edit and after undo. The saved workbook retained all 16 ZIP parts (ten unchanged),
all 431 formula caches verified, and all three charts reopened successfully in
the calculation engine. The operating and capital sheets remain byte-identical
to the prior revision; only the new chart helper column changes statement cells.

One automatic-open browser run measured 0.12 s to grid ready, 0.15 s total through
calculation, 0.04 s for the verified edit, and 0.03 s for save. These timings describe
this revision only; they are single-run observations with the same exclusions above.
Evidence: `spcx-income-net-income-chart-timings.json` and
`spcx-income-net-income-chart-save-validation.json` in the output directory.

## Historical AI-bill benchmarks

The previous 25,000-request workbook remains available as `public/the-ai-bill.xlsx`
with its original metadata. It is no longer the recording demo's active file.
Its measurements below are retained as evidence of that heavier workload;
the income-statement redesign is not a fix for the AI model's calculation bottleneck.

### Earlier runs before native chart display — October 22, 2026

This workbook contains 125,127 formulas and 275,226 populated cells in a
2,385,835-byte XLSX. Its modeled bill changes from $2,019.65 to $877.02
(56.6% lower) with the 20% → 80% cached-input assumption. This is a scenario
result using fictional rates, not a claim about achievable provider savings.

Two consecutive runs in the in-app Chromium 155 browser on an Apple M2 Pro
with 16 GB RAM, with the production local build described above:

| Run | Grid ready | Fully calculated, total | Edit → updated result | Save |
| --- | ---: | ---: | ---: | ---: |
| First | 1.30 s | 1.75 s | 5.71 s | 1.36 s |
| Start another take | 1.16 s | 1.63 s | 3.77 s | 1.39 s |

Both runs verified the baseline and changed totals with no cached fallback
formulas. A separate check of the first saved file validated all 125,127
formula caches, well-formed XML, no removed ZIP parts, and byte-identical
chart XML. Ten of fourteen original ZIP parts were byte-identical. Excel's
full-recalculation flag is set; chart caches themselves are preserved unchanged.
The raw timing reports and save validation are in `outputs/reddit-demo-20261022/`
at the repository root (`ai-measured-timings.json`,
`ai-measured-timings-take2.json`, and `ai-save-validation.json`).

### Native line chart run — October 22, 2026

The current local source, same workbook/hash and machine, in Chromium 155:

| Grid ready | Fully calculated, total | Edit → verified result and chart | Save |
| ---: | ---: | ---: | ---: |
| 1.23 s | 1.73 s | 5.68 s | 1.30 s |

The edit changed 75,087 cells. Engine evaluation took 5.30 s, change tracking
4.5 ms, result serialization 276.5 ms, JSON decoding 19.3 ms, and applying
the result on the page 26 ms. Reading the 31 chart values was below the
browser timer's resolution. Evaluation accounts for about 93% of the observed
edit time; the incremental calculation path is the next performance target.
These numbers are one local run, not a SheetJS benchmark.

The saved file has all 14 original parts: nine byte-identical, five rewritten.
Independent validation checked all 125,127 formula caches, the 31 chart values
and category labels, and byte-identical chart XML outside its data caches.
The Excel chart remains a native object with its original ranges and formatting.
Raw evidence: `outputs/reddit-demo-20261022/ai-chart-timings.json` and
`ai-chart-save-validation.json` at the repository root.

Undo restored the original bill and chart. Repeating the edit in the same
session then took **12.21 s**, including **10.47 s** in evaluation and **1.05 s**
in serialization. Both the bill and chart still verified. This slower repeat
is retained in `ai-chart-after-undo-timings.json`; its open/calculate/save
fields are the earlier measurements from the same session, not a new take.
Do not advertise this workload as consistently fast until that variation and
the incremental evaluation cost are resolved.
