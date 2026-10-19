# SpreadsheetBench lane

Runs GridPath's real product path against [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench)
and scores it with a port of the official comparator. Built on the same
self-driving eval mode as `eval/run-gridpath.mjs`, plus a manifest mode
that runs many tasks per app launch.

Reference points (others' published numbers, single attempt):
UniverAgent 68.86% on the full 912 set; on the 400-task Verified set,
Shortcut 86%, Decide 82.5%, GPT for Work 92.5%. Verified is what current
systems report on, and it removes the ambiguous tasks — run it first.

## Layout

```
eval/spreadsheetbench/
  run.mjs                          runner (stages inputs, launches app in batches, grades)
  grade.py                         grader (openpyxl, official comparator port + LibreOffice recalc)
  spreadsheetbench_verified_400/   dataset (GITIGNORED — download below)
    dataset.json                   id, instruction, instruction_type, answer_position, answer_sheet
    spreadsheet/<id>/1_<id>_init.xlsx, 1_<id>_golden.xlsx, prompt.txt
    _recalc/                       cached LibreOffice recalcs of the goldens
  runs/<run>/                      GITIGNORED
    run.json                       config + invocations
    <id>/output.xlsx               the app's saved answer (input staged here, saved in place)
    <id>/output.recalc.xlsx        LibreOffice-recalculated copy used for value grading
    <id>/meta.json                 model, effort, duration, tokens, batches, error
    <id>/task.json                 the dataset record
    _manifests/, _logs/            per-batch task lists and app logs
    results.json                   per-task verdicts + summary
    eval_official_format.json      official per-task record shape
```

## Setup

```bash
cd eval/spreadsheetbench
curl -sL -o v.tar.gz https://raw.githubusercontent.com/RUCKBReasoning/SpreadsheetBench/main/data/spreadsheetbench_verified_400.tar.gz
tar xzf v.tar.gz && rm v.tar.gz        # → spreadsheetbench_verified_400/
# full set: data/spreadsheetbench_912_v0.1.tar.gz (naming <k>_<id>_input/answer.xlsx, 3 cases per task)

pip3 install openpyxl                  # grader
# LibreOffice at /Applications/LibreOffice.app (or SOFFICE_PATH) for recalc

# app: same prereqs as eval/README.md
npx vite build && (cd src-tauri && cargo build --release --features custom-protocol)
# launch the release binary once by hand → log in → onboarding → quit
# set the model/effort under test in the app's settings; quit the app
```

## Run

```bash
node eval/spreadsheetbench/run.mjs --run smoke --ids 10452,13-1,17-35   # 3 tasks
node eval/spreadsheetbench/run.mjs --run v400-opus-high                  # all 400
node eval/spreadsheetbench/run.mjs --run v400-opus-high                  # again = resume
node eval/spreadsheetbench/run.mjs --run v400-opus-high --retry-failed   # re-run crashed/errored tasks
python3 eval/spreadsheetbench/grade.py v400-opus-high -v                 # re-grade only
```

Options: `--limit/--offset`, `--shard i/n`, `--batch n` (tasks per launch,
default 25), `--task-timeout m` (default 10; the in-app driver stops the
turn, the wrapper kills a stalled app 3 min later), `--prompt-suffix "…"`
(default none — the instruction goes in verbatim; record any suffix you
use when reporting), `--dry-run`, `--no-grade`.

The app window opens and must stay visible and unlocked for the whole run
(see the caveats in `eval/README.md` — a locked screen suspends the webview
and every task times out). Budget roughly 1–3 minutes per task.

## Scoring

`grade.py` ports `evaluation.py` verbatim: `openpyxl` `data_only` values at
`answer_position`, numbers rounded to 2 dp, datetimes to serial days,
numeric strings coerced, `""` == `None`, type mismatch = miss, all ranges
must match. Both output and golden are first recalculated through
LibreOffice headless (the official pipeline does the same with
`open_spreadsheet.py`). Deviations, each logged with `-v`: `answer_sheet` is
honoured for ranges without a sheet prefix (official takes the first sheet;
differs on one Verified task), commas inside quoted sheet names don't split
the range list, and `A:G` ranges expand to the golden's used rows.

Reporting: quote `results.json` → `summary.pass_rate`, the model/effort, the
dataset name, and whether a prompt suffix was used.
