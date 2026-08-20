import { describe, expect, it } from "vitest";
import { groupBatchForReview, describeMutation } from "../agent/reviewGroups";
import type { UniverMutation } from "../types";

function cell(
  sheet: string,
  row: number,
  col: number,
  oldV: any,
  newV: any,
  newFormula: string | null = null,
): UniverMutation {
  return {
    type: "set_cell",
    address: { sheet, row, col },
    old_value: oldV,
    new_value: newV,
    old_formula: null,
    new_formula: newFormula,
  };
}

describe("groupBatchForReview", () => {
  it("groups a contiguous column run into one region", () => {
    const muts = [cell("Model", 72, 72, null, 5431.7, "=BP73*(1+BU75)"), cell("Model", 73, 72, null, 0.1), cell("Model", 74, 72, null, 4033.5)];
    const g = groupBatchForReview(muts);
    expect(g.regions).toHaveLength(1);
    expect(g.regions[0].sheet).toBe("Model");
    expect(g.regions[0].range).toBe("BU73:BU75");
    expect(g.regions[0].cells).toHaveLength(3);
    expect(g.totalCells).toBe(3);
  });

  it("splits regions on row gaps larger than the threshold", () => {
    const muts = [
      cell("Model", 10, 5, null, 1),
      cell("Model", 12, 5, null, 2),
      // 40-row gap → new region
      cell("Model", 52, 5, null, 3),
      cell("Model", 53, 5, null, 4),
    ];
    const g = groupBatchForReview(muts);
    expect(g.regions).toHaveLength(2);
    expect(g.regions[0].range).toBe("F11:F13");
    expect(g.regions[1].range).toBe("F53:F54");
  });

  it("keeps a rectangular block (set_range shape) as one region", () => {
    const muts: UniverMutation[] = [];
    for (let r = 4; r <= 8; r++) for (let c = 1; c <= 3; c++) muts.push(cell("Sheet1", r, c, null, r * c));
    const g = groupBatchForReview(muts);
    expect(g.regions).toHaveLength(1);
    expect(g.regions[0].range).toBe("B5:D9");
    expect(g.regions[0].cells).toHaveLength(15);
  });

  it("separates sheets into distinct regions and counts", () => {
    const muts = [cell("A", 0, 0, null, 1), cell("B", 0, 0, null, 2)];
    const g = groupBatchForReview(muts);
    expect(g.regions).toHaveLength(2);
    expect(g.countsBySheet).toEqual({ A: 1, B: 1 });
  });

  it("uses the labelOf callback for the region's first row", () => {
    const muts = [cell("Model", 20, 10, null, 1), cell("Model", 21, 10, null, 2)];
    const g = groupBatchForReview(muts, (sheet, row) => (sheet === "Model" && row === 20 ? "Revenue build" : null));
    expect(g.regions[0].label).toBe("Revenue build");
  });

  it("attaches per-cell row labels and column headers, memoized per row/col", () => {
    // One row filled across three columns — labelOf hit once, headerOf per column.
    const muts = [cell("Model", 74, 72, null, 0.09), cell("Model", 74, 73, null, 0.09), cell("Model", 74, 74, null, 0.08)];
    const labelCalls: string[] = [];
    const headerCalls: string[] = [];
    const labels: Record<number, string> = { 74: "ARPU growth %" };
    const headers: Record<number, string> = { 72: "FY2026E", 73: "FY2027E", 74: "FY2028E" };
    const g = groupBatchForReview(
      muts,
      (_s, row) => { labelCalls.push(`r${row}`); return labels[row] ?? null; },
      (_s, col) => { headerCalls.push(`c${col}`); return headers[col] ?? null; },
    );
    const cells = g.regions[0].cells;
    expect(cells.map((c) => c.rowLabel)).toEqual(["ARPU growth %", "ARPU growth %", "ARPU growth %"]);
    expect(cells.map((c) => c.header)).toEqual(["FY2026E", "FY2027E", "FY2028E"]);
    expect(labelCalls).toEqual(["r74"]);
    expect(headerCalls).toEqual(["c72", "c73", "c74"]);
  });

  it("routes non-cell mutations to the others list with descriptions", () => {
    const muts: UniverMutation[] = [
      cell("Model", 0, 0, null, 1),
      { type: "insert_rows", sheet: "Model", before: 10, count: 2 },
      {
        type: "set_format",
        sheet: "Model",
        range: "A1:B2",
        cells: [{ row: 0, col: 0 }],
        old_format: [{ row: 0, col: 0, format: null }],
        new_format: { bold: true },
      },
      { type: "create_sheet", name: "Summary" },
    ];
    const g = groupBatchForReview(muts);
    expect(g.totalCells).toBe(1);
    expect(g.others).toHaveLength(3);
    expect(g.others[0].desc).toContain("Insert 2 rows before row 11");
    expect(g.others[1].desc).toContain("Format Model!A1:B2");
    expect(g.others[2].desc).toContain('New sheet "Summary"');
  });

  it("prefers formulas over values in the diff strings", () => {
    const g = groupBatchForReview([cell("M", 0, 0, "old", 42, "=SUM(A1:A9)")]);
    expect(g.regions[0].cells[0].newValue).toBe("=SUM(A1:A9)");
    expect(g.regions[0].cells[0].isFormula).toBe(true);
  });

  it("collapses a single cell to a bare A1 address", () => {
    const g = groupBatchForReview([cell("M", 4, 2, null, 1)]);
    expect(g.regions[0].range).toBe("C5");
  });
});

describe("describeMutation", () => {
  it("describes structural mutations", () => {
    expect(
      describeMutation({ type: "rename_sheet", old_name: "Old", new_name: "New" }),
    ).toBe('Rename sheet "Old" → "New"');
    expect(
      describeMutation({ type: "delete_columns", sheet: "M", start: 26, count: 3 }),
    ).toBe("Delete 3 columns from AA (M)");
    expect(
      describeMutation({ type: "hide_rows", sheet: "M", rows: [6, 7, 8] }),
    ).toBe("Hide rows 7, 8, 9 (M)");
  });
});
