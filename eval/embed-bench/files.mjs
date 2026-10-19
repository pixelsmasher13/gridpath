// The benchmark's workbooks. The two real models are private: point the
// environment variables at your own copies (a workbook whose variable is
// unset is skipped). Nothing here copies them into the repo.
const fixture = new URL("../fixtures/rich-model.xlsx", import.meta.url).pathname;

export const FILES = Object.fromEntries(
  Object.entries({
    "rich-model": { path: fixture, editBudgetMs: 5_000, label: "small fixture (8 KB; chart, drawing, custom XML)" },
    analyst: { path: process.env.ANALYST_MODEL, editBudgetMs: 60_000, label: "analyst model (1.4 MB, 55k formulas)" },
    large: { path: process.env.LARGE_MODEL, editBudgetMs: 120_000, label: "data-intensive model (over 10 MB, 2.3M cells)" },
  }).filter(([, f]) => f.path),
);
