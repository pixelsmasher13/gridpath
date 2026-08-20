# GridPath agent evals — head-to-head vs Claude Code

Same task, same model, two harnesses. Assertions are layout-independent
(label search across sheets) so different layouts for the same task grade
fairly. Every assertion is pass / fail / warn — warn means "can't be graded
fairly" (e.g. formula cells saved without cached values).

## Layout

```
eval/
  tasks/*.json        task specs: prompt, optional start file, assertions
  corpus/             real workbooks for edit tasks (GITIGNORED — proprietary)
  runs/<task>/<ts>-<harness>/   output.xlsx, meta.json, *.grade.json (GITIGNORED)
  grade.mjs           grader
  run-claude-code.mjs headless Claude Code runner (runs + grades)
  run-gridpath.mjs    self-driving GridPath runner (runs + grades)
```

## Running the Claude Code side (automated)

```bash
node eval/run-claude-code.mjs dh-income-statement --model claude-opus-4-8 --yolo
```

`--yolo` = `--dangerously-skip-permissions`, needed for unattended runs
(Claude Code uses Bash/python for xlsx work). The workdir is isolated under
`eval/runs/`; only use it with prompts you wrote. Timing, turns, tokens and
cost land in `meta.json`; the grade report prints and is saved next to the
output.

## Running the GridPath side (automated)

```bash
npx vite build && (cd src-tauri && cargo build --release --features custom-protocol)
node eval/run-gridpath.mjs dh-income-statement             # one task, one command
```

The `custom-protocol` feature is REQUIRED: it's what makes the binary serve
the embedded dist/ (it's passed automatically by `tauri build`, but not by
raw cargo builds — without it even a release binary loads the vite dev URL
and renders a blank window standalone). Assets embed at compile time, so
the vite build must run BEFORE the cargo build.

The wrapper pre-creates `output.xlsx` in the run dir (blank workbook for
build tasks, corpus copy for edit tasks) and launches the app binary with
`GRIDPATH_EVAL_*` env vars. The app then drives the real product path
itself: opens that file, submits the prompt verbatim, auto-accepts every
batch, saves in place, writes `meta.json` (model, effort, duration, batches,
tokens) and exits — the wrapper grades the result. `--timeout <min>`
(default 20) kills a hung run; `--binary <path>` overrides binary discovery
(release preferred, then debug).

**First-time setup (once per machine):** webview storage is per-origin, so
the standalone binary (`tauri://`) does NOT share the Firebase login or the
onboarding flag with your `tauri dev` sessions (`localhost:5173`) — an
un-setup binary boots to the login screen and the driver never engages.
Launch the release binary once by hand, log in, click through onboarding,
quit. The Claude/Codex credential itself lives in the shared SQLite, so it
carries over automatically.

Caveats: GridPath must NOT already be running (single-instance lock);
credentials and the model/effort under test come from the app's own settings
DB — set them in the app first (model parity with the CC runner is on you).
The app window opens during the run; don't touch it — but don't fully
occlude it either: macOS throttles timers in unattended windows, and one
observed run spent ~8 minutes in file load that normally takes seconds
(wall-clock pollution; the agent's own time was normal). Leave the window
visible somewhere on screen. A LOCKED SCREEN is fatal, not just slow:
at the lock screen macOS suspends webview rendering entirely, the eval
driver never engages (app.log stops right after read_workbook_file, no
session_upsert), and every run times out. Observed 2026-08-17: three
consecutive runs lost this way. Don't run the GridPath lane unattended
with auto-lock enabled; the CC lane is headless and immune. Pass --binary to override binary discovery.

Manual grading of any output still works:

```bash
node eval/grade.mjs dh-income-statement eval/runs/dh-income-statement/<ts>-gridpath/output.xlsx
# edit tasks additionally compare against the original:
node eval/grade.mjs gs-model-edit .../output.xlsx --original eval/runs/.../original.xlsx
```

## Recalc grading (`--recalc`)

Both runners grade with `--recalc` (manual grades: pass the flag yourself).
It normalizes the output through LibreOffice headless before value grading:
a throwaway profile is seeded with "always recalculate on load", the file is
converted xlsx→xlsx (which evaluates every formula and writes fresh cached
values into a `*.recalc.xlsx` copy next to the output), and VALUE-domain
assertions (`label_value`, `check_rows_zero`, `no_error_cells`) read that
copy. Structure-domain assertions (`label_formula`, `parts_preserved`,
`min_formula_cells`, …) always read the raw output — the LibreOffice
round-trip must never touch the fidelity axis. This makes value grading
pass/fail regardless of either harness's cache behavior (GridPath's surgical
save strips caches; openpyxl can't compute them and may leave stale ones),
and it's neutral: both harnesses go through the same third-party engine.
Requires LibreOffice (`SOFFICE_PATH` to override discovery); if missing, the
grader warns and falls back to raw values. Caveat: a formula using a
function LibreOffice can't evaluate recalcs to an error — if `no_error_cells`
fails on the recalc copy only, check the raw file before blaming the model.

## Grading caveats
- **`parts_preserved`** is the round-trip fidelity axis. Expect GridPath to
  pass and openpyxl-based flows to fail on rich files — that asymmetry is a
  real product difference, not a grading artifact.
- Comprehension tasks (explain-the-model) need an LLM judge and aren't in
  this phase.

## Seed tasks

| task | type | what it probes |
|---|---|---|
| `dh-income-statement` | build | research triangulation, tie-outs/check rows (rules 21/22), growth detail |
| `aapl-forecast` | build | forecast structure, quarterly layout, formula-driven EPS |
| `fixture-model-edit` | edit | trivial edit; grades WHAT SURVIVES the save on the COMMITTED fixture (chart, CF, validation, defined names, merges) — publishable, reproducible by anyone |
| `gs-model-edit` | edit | same design on a real bank model (external links, webextensions, pivots, comments) — LOCAL-ONLY: the corpus file is proprietary, never committed; results based on it aren't publishable |

The fixture is authored by `fixtures/make_rich_model.py` (openpyxl) and
committed; regenerate it after changing the script. Prefer building new
edit tasks on committed fixtures — corpus files are a private bonus axis,
not something the suite may depend on.

## Roadmap

- Hermetic build tasks (data embedded in start file) so repeatability
  doesn't depend on live web sources.
- LLM-judge assertions for comprehension tasks and qualitative layout.
