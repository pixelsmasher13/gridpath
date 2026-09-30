import { describe, expect, it } from "vitest";
import { colIndexToLetters } from "../formulaRuns";
import {
  buildWorkbookIndex,
  describeWorkbookPayload,
  findRowsInIndex,
  firstLabelIn,
  getWorkbookIndex,
  isLabelText,
  snapshotContentKey,
} from "../agent/workbookIndex";

/**
 * Snapshot builder in Univer's IWorkbookData shape (the subset the index
 * reads): sheets keyed by id with name/cellData/hidden, a workbook-level
 * style map, optional resources.
 */
function snap(
  sheets: Record<string, { name: string; cellData: any; hidden?: number }>,
  styles: any = {},
  resources?: any[],
): any {
  return {
    sheetOrder: Object.keys(sheets),
    sheets,
    styles,
    ...(resources ? { resources } : {}),
  };
}

/** cellData row from a sparse col → cell map. */
const row = (cells: Record<number, any>) => cells;
const v = (value: any, s?: string) => ({ v: value, ...(s ? { s } : {}) });
const f = (formula: string, value?: any) => ({ f: formula, ...(value !== undefined ? { v: value } : {}) });

describe("isLabelText / firstLabelIn", () => {
  it("accepts real labels, rejects flags, years, numbers, formula-ish text", () => {
    expect(isLabelText("Total revenue")).toBe(true);
    expect(isLabelText("EPS")).toBe(true);
    expect(isLabelText("Y")).toBe(false); // single-char flag column
    expect(isLabelText("2024A")).toBe(false); // year header
    expect(isLabelText("$1,234")).toBe(false);
    expect(isLabelText("=SUM(A1)")).toBe(false);
    expect(isLabelText("(loss)")).toBe(false);
    expect(isLabelText(1234)).toBe(false);
    expect(isLabelText(null)).toBe(false);
  });

  it("picks the first label in column order and clips to 60 chars", () => {
    expect(firstLabelIn(["Y", 42, "Subscriber detail", "other"])).toBe("Subscriber detail");
    expect(firstLabelIn([null, undefined, 3])).toBe(null);
    const long = "x".repeat(100);
    expect(firstLabelIn([long])).toHaveLength(60);
  });
});

describe("buildWorkbookIndex", () => {
  it("computes used range from data extent, not grid capacity", () => {
    const s = snap({
      s1: {
        name: "Model",
        cellData: { 0: row({ 0: v("Title") }), 9: row({ 3: v(42) }) },
      },
    });
    const idx = buildWorkbookIndex(s);
    expect(idx.sheets[0].usedRange).toBe("A1:D10");
  });

  it("infers row labels, counts, and sample columns", () => {
    const s = snap({
      s1: {
        name: "Model",
        cellData: {
          4: row({ 0: v("Revenue"), 1: v(100), 2: f("=B5*1.1", 110), 3: f("=C5*1.1", 121) }),
        },
      },
    });
    const r5 = buildWorkbookIndex(s).sheets[0].rows[0];
    expect(r5.row).toBe(5);
    expect(r5.label).toBe("Revenue");
    expect(r5.values).toBe(2); // label + 100
    expect(r5.formulas).toBe(2);
    expect(r5.sampleCol).toBe(3); // rightmost formula cell
  });

  it("takes a formula cell's cached string value as a label", () => {
    const s = snap({
      s1: { name: "M", cellData: { 2: row({ 0: f('=Assumptions!A1', "Gross margin"), 1: v(0.4) }) } },
    });
    expect(buildWorkbookIndex(s).sheets[0].rows[0].label).toBe("Gross margin");
  });

  it("detects the header row (majority text periods) and maps col letters", () => {
    const s = snap({
      s1: {
        name: "Model",
        cellData: {
          0: row({ 0: v("NFLX Model") }), // title row: single cell, not a header
          3: row({ 0: v("$ in millions"), 2: v("FY2019A"), 3: v("FY2020A"), 4: v("FY2021E") }),
          4: row({ 0: v("Revenue"), 2: v(20156), 3: v(24996), 4: f("=D5*1.15", 28745) }),
        },
      },
    });
    const sheet = buildWorkbookIndex(s).sheets[0];
    expect(sheet.headerRow).toBe(4);
    expect(sheet.headers).toEqual({ C: "FY2019A", D: "FY2020A", E: "FY2021E" });
  });

  it("accepts integer years and date-formatted serials as header cells", () => {
    const s = snap(
      {
        s1: {
          name: "M",
          cellData: {
            1: row({ 0: v("Periods"), 1: v(2024), 2: v(2025), 3: v(45000, "dateStyle") }),
          },
        },
      },
      { dateStyle: { n: { pattern: "mmm-yy" } } },
    );
    const sheet = buildWorkbookIndex(s).sheets[0];
    expect(sheet.headerRow).toBe(2);
    expect(sheet.headers.B).toBe("2024");
    expect(sheet.headers.D).toBe("2023-03-15"); // serial 45000
  });

  it("does not mistake a numeric data row for a header row", () => {
    const s = snap({
      s1: {
        name: "M",
        cellData: {
          0: row({ 0: v("Revenue"), 1: v(10.5), 2: v(11.2), 3: v(12.9) }),
          1: row({ 0: v("COGS"), 1: v(4.1), 2: v(4.4), 3: v(4.9) }),
        },
      },
    });
    expect(buildWorkbookIndex(s).sheets[0].headerRow).toBe(null);
  });

  it("splits sections on gaps of 2+ blank rows and titles them by first label", () => {
    const s = snap({
      s1: {
        name: "M",
        cellData: {
          4: row({ 0: v("Revenue build"), 1: v(1) }),
          5: row({ 0: v("Streaming"), 1: v(2) }),
          6: row({ 0: v("DVD"), 1: v(3) }),
          // rows 8-9 blank (2 blank rows) → new section
          9: row({ 0: v("Subscriber detail"), 1: v(4) }),
          10: row({ 0: v("US subs"), 1: v(5) }),
        },
      },
    });
    const secs = buildWorkbookIndex(s).sheets[0].sections;
    expect(secs).toEqual([
      { startRow: 5, endRow: 7, title: "Revenue build" },
      { startRow: 10, endRow: 11, title: "Subscriber detail" },
    ]);
  });

  it("does not split on a single blank row", () => {
    const s = snap({
      s1: {
        name: "M",
        cellData: {
          0: row({ 0: v("Alpha"), 1: v(1) }),
          2: row({ 0: v("Beta"), 1: v(2) }), // one blank row between
        },
      },
    });
    expect(buildWorkbookIndex(s).sheets[0].sections).toHaveLength(1);
  });

  it("starts a new section at a bold label-only row, but not at a bold total row", () => {
    const s = snap(
      {
        s1: {
          name: "M",
          cellData: {
            0: row({ 0: v("Income statement", "bold") }), // bold header, no data
            1: row({ 0: v("Revenue"), 1: v(100) }),
            2: row({ 0: v("Total revenue", "bold"), 1: f("=B2", 100) }), // bold WITH data: no split
            3: row({ 0: v("Margins", "bold") }), // bold header again → split
            4: row({ 0: v("Gross margin"), 1: v(0.4) }),
          },
        },
      },
      { bold: { bl: 1 } },
    );
    const secs = buildWorkbookIndex(s).sheets[0].sections;
    expect(secs).toEqual([
      { startRow: 1, endRow: 3, title: "Income statement" },
      { startRow: 4, endRow: 5, title: "Margins" },
    ]);
  });

  it("compresses column-translated formula runs and keeps distinct formulas separate", () => {
    const cells: Record<number, any> = { 0: v("Revenue") };
    for (let c = 2; c <= 51; c++) {
      // =B5*1.1, =C5*1.1, … exact column translations
      cells[c] = f(`=${colIndexToLetters(c - 1)}5*1.1`, 100 + c);
    }
    cells[60] = f("=SUM(C5:AZ5)", 9999); // separate formula, own run
    const s = snap({ s1: { name: "M", cellData: { 4: row(cells) } } });
    const runs = buildWorkbookIndex(s).sheets[0].formulaRuns;
    expect(runs).toHaveLength(2);
    expect(runs[0]).toEqual({ row: 5, startCol: 2, endCol: 51, formula: "=B5*1.1" });
    expect(runs[1]).toEqual({ row: 5, startCol: 60, endCol: 60, formula: "=SUM(C5:AZ5)" });
  });

  it("flags hidden sheets", () => {
    const s = snap({
      s1: { name: "Visible", cellData: { 0: row({ 0: v("x1") }) } },
      s2: { name: "Helper", cellData: { 0: row({ 0: v("x2") }) }, hidden: 1 },
    });
    const idx = buildWorkbookIndex(s);
    expect(idx.sheets.map((sh) => sh.hidden)).toEqual([false, true]);
  });

  it("parses defined names from the snapshot resources", () => {
    const s = snap(
      { s1: { name: "M", cellData: {} } },
      {},
      [
        {
          name: "SHEET_DEFINED_NAME_PLUGIN",
          data: JSON.stringify({
            id1: { id: "id1", name: "GrowthRate", formulaOrRefString: "Model!$B$22" },
            id2: { id: "id2", name: "TaxRate", formulaOrRefString: "Model!$B$23" },
          }),
        },
      ],
    );
    expect(buildWorkbookIndex(s).definedNames).toEqual([
      { name: "GrowthRate", ref: "Model!$B$22" },
      { name: "TaxRate", ref: "Model!$B$23" },
    ]);
  });
});

describe("getWorkbookIndex memoization", () => {
  it("returns the same object while content is unchanged, rebuilds on edit", () => {
    const s1 = snap({ s1: { name: "M", cellData: { 0: row({ 0: v("Revenue"), 1: v(1) }) } } });
    const a = getWorkbookIndex("/tmp/wb.xlsx", s1);
    const b = getWorkbookIndex("/tmp/wb.xlsx", s1);
    expect(b).toBe(a);
    // Same content, fresh snapshot object → still memoized (content-keyed).
    const s2 = snap({ s1: { name: "M", cellData: { 0: row({ 0: v("Revenue"), 1: v(1) }) } } });
    expect(getWorkbookIndex("/tmp/wb.xlsx", s2)).toBe(a);
    // Content change → rebuild.
    const s3 = snap({ s1: { name: "M", cellData: { 0: row({ 0: v("Revenue"), 1: v(2) }) } } });
    const c = getWorkbookIndex("/tmp/wb.xlsx", s3);
    expect(c).not.toBe(a);
    expect(snapshotContentKey(s3)).not.toBe(snapshotContentKey(s1));
  });

  it("keeps the key stable when only a formula cell's computed value moves (recalc jitter)", () => {
    // The engine clears/rewrites `v` on formula cells during recalc; the key
    // must depend on the formula TEXT only, or a no-edit turn re-bills the
    // whole cached context block (the "thank you" 181K cache-write).
    const settling = snap({
      s1: { name: "M", cellData: { 0: row({ 0: v(10), 1: f("=A1*2") }) } },
    });
    const settled = snap({
      s1: { name: "M", cellData: { 0: row({ 0: v(10), 1: f("=A1*2", 20) }) } },
    });
    expect(snapshotContentKey(settled)).toBe(snapshotContentKey(settling));
    // But an authored change — formula text or a literal — still re-keys.
    const editedFormula = snap({
      s1: { name: "M", cellData: { 0: row({ 0: v(10), 1: f("=A1*3", 20) }) } },
    });
    const editedLiteral = snap({
      s1: { name: "M", cellData: { 0: row({ 0: v(11), 1: f("=A1*2", 20) }) } },
    });
    expect(snapshotContentKey(editedFormula)).not.toBe(snapshotContentKey(settled));
    expect(snapshotContentKey(editedLiteral)).not.toBe(snapshotContentKey(settled));
  });
});

describe("findRowsInIndex", () => {
  const model = snap({
    s1: {
      name: "Model",
      cellData: {
        3: row({ 0: v("$ in millions"), 2: v("FY2019A"), 3: v("FY2020A"), 4: v("FY2021E") }),
        8: row({ 0: v("Revenue build"), 1: v(1) }),
        9: row({ 0: v("Streaming revenue"), 2: v(100), 3: f("=C10*1.2", 120) }),
        208: row({ 0: v("Total revenue"), 3: f("=D10", 120) }),
      },
    },
    s2: {
      name: "Assumptions",
      cellData: { 0: row({ 0: v("Revenue growth"), 1: v(0.15) }) },
    },
  });
  const idx = buildWorkbookIndex(model);

  it("finds rows by case-insensitive substring across sheets", () => {
    const { matches } = findRowsInIndex(idx, "revenue", null, 20);
    const labels = matches.map((m) => `${m.sheet}!${m.label}`);
    expect(labels).toContain("Model!Total revenue");
    expect(labels).toContain("Model!Streaming revenue");
    expect(labels).toContain("Assumptions!Revenue growth");
  });

  it("ranks an exact match first and reports the containing section", () => {
    const { matches } = findRowsInIndex(idx, "total revenue", null, 20);
    expect(matches[0].label).toBe("Total revenue");
    expect(matches[0].row).toBe(209);
    expect(matches[0].sampleCol).toBe(3);
  });

  it("matches by word-initial fuzzy tokens", () => {
    const { matches } = findRowsInIndex(idx, "tot rev", null, 20);
    expect(matches.some((m) => m.label === "Total revenue")).toBe(true);
  });

  it("matches header cells and reports their column letter", () => {
    const { matches } = findRowsInIndex(idx, "FY2020A", null, 20);
    const header = matches.find((m) => m.column !== undefined);
    expect(header).toMatchObject({ sheet: "Model", row: 4, column: "D", label: "FY2020A" });
  });

  it("respects sheet filter and max_results with a total count", () => {
    const filtered = findRowsInIndex(idx, "revenue", "Assumptions", 20);
    expect(filtered.matches.every((m) => m.sheet === "Assumptions")).toBe(true);
    const capped = findRowsInIndex(idx, "revenue", null, 1);
    expect(capped.matches).toHaveLength(1);
    expect(capped.total).toBeGreaterThan(1);
  });
});

describe("describeWorkbookPayload", () => {
  const idx = buildWorkbookIndex(
    snap({
      s1: {
        name: "Model",
        cellData: {
          3: row({ 0: v("$ in millions"), 2: v("FY2019A"), 3: v("FY2020A"), 4: v("FY2021E") }),
          5: row({ 0: v("Revenue build"), 1: v(1) }),
          6: row({ 0: v("Streaming"), 1: v(2) }),
        },
      },
      s2: { name: "Empty", cellData: {} },
    }),
  );

  it("ships name, used range, header row/headers, and section list per sheet", () => {
    const p: any = describeWorkbookPayload(idx, null);
    expect(p.sheets).toHaveLength(2);
    const model = p.sheets[0];
    expect(model.name).toBe("Model");
    expect(model.used_range).toBe("A1:E7");
    expect(model.header_row).toBe(4);
    expect(model.headers.C).toBe("FY2019A");
    // One blank row between header row (4) and content → one section; the
    // header row's period text doesn't title it, the first real label does.
    expect(model.sections).toEqual([{ rows: "4-7", title: "Revenue build" }]);
    expect(p.sheets[1].used_range).toBe(null);
  });

  it("filters by sheet and errors helpfully on an unknown name", () => {
    const p: any = describeWorkbookPayload(idx, "Model");
    expect(p.sheets).toHaveLength(1);
    const err: any = describeWorkbookPayload(idx, "Nope");
    expect(err.error).toContain("Nope");
    expect(err.sheets).toEqual(["Model", "Empty"]);
  });
});
