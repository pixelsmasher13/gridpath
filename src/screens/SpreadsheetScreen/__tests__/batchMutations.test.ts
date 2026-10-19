import { describe, expect, it } from "vitest";
import { isLiteralOnlyBatch, validateBatchLayout } from "../agent/batchMutations";
import type { UniverMutation } from "../types";

const lit = (row: number, value: string | number): UniverMutation => ({
  type: "set_cell",
  address: { sheet: "M", row, col: 0 },
  old_value: null,
  new_value: value,
  new_formula: null,
});

const formula = (row: number, f: string): UniverMutation => ({
  type: "set_cell",
  address: { sheet: "M", row, col: 0 },
  old_value: null,
  new_value: f,
  new_formula: f,
});

describe("isLiteralOnlyBatch", () => {
  it("true for a pure data lay", () => {
    expect(isLiteralOnlyBatch([lit(0, "Revenue"), lit(1, 97690)])).toBe(true);
  });

  it("false as soon as any cell writes a formula", () => {
    expect(isLiteralOnlyBatch([lit(0, "Revenue"), formula(1, "=A1*2")])).toBe(false);
  });

  it("false for format / clear / structural mutations", () => {
    expect(
      isLiteralOnlyBatch([
        lit(0, 1),
        { type: "set_format", sheet: "M", range: "A1", cells: [{ row: 0, col: 0 }], old_format: [], new_format: { bold: true } },
      ]),
    ).toBe(false);
    expect(
      isLiteralOnlyBatch([
        { type: "clear_range", sheet: "M", range: "A1", cells: [{ row: 0, col: 0, old_value: null, old_formula: null }] },
      ]),
    ).toBe(false);
    expect(isLiteralOnlyBatch([{ type: "insert_rows", sheet: "M", before: 0, count: 1 }])).toBe(false);
  });

  it("false for an empty batch (nothing to skip for; keep the settle)", () => {
    expect(isLiteralOnlyBatch([])).toBe(false);
  });
});

describe("validateBatchLayout — cross-sheet references (regression)", () => {
  const gridWith = (cells: Record<string, { value: any; formula: string | null }>) => ({
    getCell: (sheet: string, row: number, col: number) => cells[`${sheet}!${row},${col}`] ?? null,
  }) as any;

  it("does NOT flag a sheet-qualified reference to the same address on another sheet", () => {
    // Assumptions!H8 = ='Income Statement'!H8/'Income Statement'!C8-1
    // Same A1 address, different sheet — an ordinary cross-sheet link.
    const grid = gridWith({
      "Assumptions!7,7": { value: 0.899, formula: "='Income Statement'!H8/'Income Statement'!C8-1" },
    });
    const errors = validateBatchLayout(grid, [
      { type: "set_cell", address: { sheet: "Assumptions", row: 7, col: 7 }, old_value: null, new_value: null, new_formula: "='Income Statement'!H8/'Income Statement'!C8-1" } as any,
    ]);
    expect(errors).toEqual([]);
  });

  it("still flags a genuine unqualified self-reference", () => {
    const grid = gridWith({ "Model!7,7": { value: 0, formula: "=H8+1" } });
    const errors = validateBatchLayout(grid, [
      { type: "set_cell", address: { sheet: "Model", row: 7, col: 7 }, old_value: null, new_value: null, new_formula: "=H8+1" } as any,
    ]);
    expect(errors).toHaveLength(1);
    expect(errors[0].kind).toBe("self_ref");
  });

  it("still flags a self-reference inside a range", () => {
    const grid = gridWith({ "Model!7,7": { value: 0, formula: "=SUM(H1:H8)" } });
    const errors = validateBatchLayout(grid, [
      { type: "set_cell", address: { sheet: "Model", row: 7, col: 7 }, old_value: null, new_value: null, new_formula: "=SUM(H1:H8)" } as any,
    ]);
    expect(errors).toHaveLength(1);
  });
});
