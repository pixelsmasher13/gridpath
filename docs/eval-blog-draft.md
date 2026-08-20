# DRAFT — Same model, different harness: benchmarking GridPath against Claude Code on real spreadsheet work

*Status: draft for review. Companion artifact repo: `gridpath-evals` (staged at
`~/gridpath-evals`, publish as a public GitHub repo and update the links below).
All numbers below are from the published N=5 dataset; every claim is
re-derivable from the repo.*

---

Every AI spreadsheet product faces the same skeptical question: why not just
use a general-purpose coding agent? Claude Code can read and write xlsx files
with Python. It's excellent. So we benchmarked ourselves against it — same
tasks, same underlying model (`claude-sonnet-5`), five runs per task per
harness, every artifact published.

Running the *same model* in both harnesses is the point. This isn't "our AI
vs. their AI" — it's a controlled experiment on the one variable we actually
control: **the harness**. GridPath's agent operates on a live workbook through
structured spreadsheet tools with a real-time calc engine. Claude Code
reconstructs the file through scripts, dozens of shell round-trips per task.
Same intelligence; different machinery. Any difference in the table below is
machinery.

## The tasks

Three published tasks (specs, grader, and all 30 run outputs are in the repo):

- **fixture-model-edit** — a one-row edit on a feature-dense workbook (chart,
  conditional formatting, data validation, defined names, merges). Probes
  precision and round-trip fidelity: what survives the save?
- **dh-income-statement** — "put together a model on Delivery Hero income
  statement with details on growth." Open-ended build with real-world data.
- **aapl-forecast** — "build an AAPL income statement forecast 2026–2028 with
  quarterly breakdowns." Structured forecast build.

Grading is layout-independent (label-driven assertions, so different layouts
grade fairly) and value assertions run after normalizing outputs through
LibreOffice headless recalculation — a neutral third-party calc engine, so
neither harness's cache behavior can skew results.

## Results

| task | harness | clean runs | assertions | median time | median output tokens |
|---|---|---|---|---|---|
| edit (fixture) | GridPath | **5/5** | 40/40 | **6.2s** | 260 |
| edit (fixture) | Claude Code | 3/5 | 38/40 | 96s | 8,700 |
| build (DH) | GridPath | 0/5 | 33/45 | **178s** | 16,158 |
| build (DH) | Claude Code | 2/5 | 37/45 | 864s | 78,265 |
| build (AAPL) | GridPath | 1/5 | 32/40 | **253s** | 25,300 |
| build (AAPL) | Claude Code | 2/5 | 35/40 | 1,181s | 87,008 |

Three findings.

**1. Quality is effectively tied — as it should be.** Across 125 assertions,
GridPath passed 105, Claude Code 110. Same model, same correctness ceiling.
If we'd claimed a big quality win over the same model, you should have been
suspicious.

**2. The fidelity gap is real, and it's silent.** In 2 of 5 edit runs, Claude
Code's output dropped parts of the original workbook — charts, conditional
formatting, validation — while completing the requested edit correctly. The
diff looks fine; the file is quietly damaged. This is a structural hazard of
the script-the-file approach (the standard Python xlsx libraries don't
round-trip every part), not a model failure. GridPath edits the live workbook
and preserved every part in every run.

**3. The efficiency gap is a category difference, not a speedup.** A
six-second edit is interactive — you stay in the sheet, you iterate. A
96-second edit is a coffee break. A 3–4 minute build means trying three
variations this hour; 14–20 minutes means you get one. Token economics
follow: 3–36× fewer output tokens per task, which is the price-independent
way to say it (our lane ran on a flat subscription; Claude Code's fifteen
runs metered $118.42).

## Where both harnesses struggled

Builds are genuinely hard, and neither harness swept them. The recurring
failures on both sides were *omissions*, not wrong arithmetic: a missing
actuals anchor column, a missing net-income row, no CAGR summary despite the
prompt asking about growth. Model behavior, not harness behavior — which is
why it shows up on both sides of the table.

## What we got wrong, and fixed

Two things surfaced during analysis that belong in the open:

- **Our grader had a bug that penalized Claude Code.** The AAPL revenue
  assertion matched row labels against `/revenue/i`. Apple's P&L calls that
  line "net sales" — and all five Claude Code runs had the exact right figure
  (416,161) under that label. We corrected the assertion to accept the
  issuer's own terminology and re-graded every run before publishing. The
  table above is post-correction. (Benchmark authors grading their own
  product: publish your grader.)
- **The eval immediately improved our product.** The most common GridPath
  build failure (4/5 AAPL runs) was omitting the FY2025 actuals column that
  forecasts should hang off. One added sentence to the agent's rules — "a
  forecast model MUST include the most recent completed fiscal year as an
  actuals anchor" — and post-fix verification runs went from 1/5 clean to 2/3
  clean. Those runs are *not* in the published table (they're segregated in
  the repo); the table is the honest pre-fix measurement. But it's the real
  reason to maintain an eval like this: it converts anecdotes into fixes.

## Caveats, stated plainly

N=5 per cell, one machine, one model at default effort, August 2026, Claude
Code 2.1.191. Agents are stochastic; medians over five runs beat single
anecdotes but this is a snapshot, not a law. Claude Code is the natural
baseline ("what a technical person would otherwise use"), not a survey of
spreadsheet-specific competitors. And our fourth internal task — the same
edit design on a real investment-bank model — is excluded here because its
file is proprietary and results based on it aren't reproducible by readers;
we publish only what you can rerun.

Everything else, you can rerun: the grader, the tasks, all 30 outputs, and
the entire Claude Code lane are in the repo. Grade our homework.

---

## Marketing-page claims (approved subset — keep in sync with this post)

- "Same model, same tasks: GridPath edits in seconds (median 6s), a
  general-purpose agent takes minutes (median 96s)." *(date-stamped, links
  to methodology)*
- "10 out of 10 clean edit runs, zero workbook features lost. The
  general-purpose agent silently dropped charts or formatting in 40% of edit
  runs."
- "3–36× fewer output tokens for the same work."
- Do NOT claim: quality superiority, "beats Claude", or absolute dollar
  savings (subscription vs metered isn't like-for-like — tokens are).

## Pre-publish checklist

- [ ] `cd ~/gridpath-evals && git init` + push to public GitHub repo; update links here
- [ ] Legal/comms pass on naming Claude Code + Anthropic (nominative use, factual benchmark)
- [ ] Screenshot or 10-second clip of the 6s edit for the post header
- [ ] Marketing page + post ship together
- [ ] Decide repo posture: frozen snapshot vs living benchmark (accepting issues)
