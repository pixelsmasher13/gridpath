import { describe, expect, it } from "vitest";

import { cellDataToInput, extractEditedCells, parseFormattedNumber, valuesAgree } from "../ironcalcShadow";

describe("cellDataToInput", () => {
  it("prefers the formula text", () => {
    expect(cellDataToInput({ f: "=SUM(A1:B2)", v: 42 })).toBe("=SUM(A1:B2)");
  });
  it("stringifies plain values", () => {
    expect(cellDataToInput({ v: 42 })).toBe("42");
    expect(cellDataToInput({ v: "hello" })).toBe("hello");
  });
  it("clears on null cell data", () => {
    expect(cellDataToInput(null)).toBe("");
  });
  it("rejects shared-formula refs without text", () => {
    expect(cellDataToInput({ si: "0" })).toBeNull();
  });
  it("ignores style-only patches", () => {
    expect(cellDataToInput({ s: "xyz" })).toBeNull();
  });
});

describe("extractEditedCells", () => {
  it("decodes a row/col matrix (typed cell edit)", () => {
    const cells = extractEditedCells({
      subUnitId: "sheet-1",
      value: { "4": { "2": { v: 123 } } },
    });
    expect(cells).toEqual([{ row0: 4, col0: 2, input: "123" }]);
  });

  it("decodes a single cell value over a range", () => {
    const cells = extractEditedCells({
      range: { startRow: 1, startColumn: 1, endRow: 1, endColumn: 3 },
      value: { v: 7 },
    });
    expect(cells).toEqual([
      { row0: 1, col0: 1, input: "7" },
      { row0: 1, col0: 2, input: "7" },
      { row0: 1, col0: 3, input: "7" },
    ]);
  });

  it("returns null for oversized ranges (desync instead of flooding IPC)", () => {
    const cells = extractEditedCells({
      range: { startRow: 0, startColumn: 0, endRow: 999, endColumn: 99 },
      value: { v: 1 },
    });
    expect(cells).toBeNull();
  });

  it("honors a caller-supplied cap (ENGINE mode mirrors far more than shadow samples)", () => {
    // 500-cell matrix: over shadow's 400 default, fine for engine mode. A
    // shadow-sized cap on the engine path degraded it on any 400+-cell
    // agent write.
    const matrix: Record<string, Record<string, { v: number }>> = {};
    for (let r = 0; r < 50; r++) {
      matrix[String(r)] = {};
      for (let c = 0; c < 10; c++) matrix[String(r)][String(c)] = { v: r * 10 + c };
    }
    expect(extractEditedCells({ value: matrix })).toBeNull();
    expect(extractEditedCells({ value: matrix }, 50_000)).toHaveLength(500);
  });

  it("returns null on unrecognized shapes", () => {
    expect(extractEditedCells({ value: "weird" })).toBeNull();
    expect(extractEditedCells({})).toBeNull();
  });
});

describe("valuesAgree", () => {
  it("compares numbers with relative tolerance", () => {
    expect(valuesAgree(1000000.0000001, 1000000.0000002)).toBe(true);
    expect(valuesAgree(1, 2)).toBe(false);
  });
  it("treats empty and null as equivalent", () => {
    expect(valuesAgree(null, "")).toBe(true);
    expect(valuesAgree("", null)).toBe(true);
  });
  it("compares strings exactly", () => {
    expect(valuesAgree("abc", "abc")).toBe(true);
    expect(valuesAgree("abc", "abd")).toBe(false);
  });
  it("coerces number-formatted display strings loosely", () => {
    expect(valuesAgree("4,624.1", 4624.0997)).toBe(true);
    expect(valuesAgree("$70.00", 70)).toBe(true);
    expect(valuesAgree("(65.2%)", -0.65239)).toBe(true);
    expect(valuesAgree("29.3%", 0.29312)).toBe(true);
    expect(valuesAgree("2,560.0", 3200)).toBe(false);
  });
});

describe("parseFormattedNumber", () => {
  it("handles currency, thousands separators, percents and parens", () => {
    expect(parseFormattedNumber("4,624.1")).toBeCloseTo(4624.1);
    expect(parseFormattedNumber("$70.00")).toBe(70);
    expect(parseFormattedNumber("(65.2%)")).toBeCloseTo(-0.652);
    expect(parseFormattedNumber("29.3%")).toBeCloseTo(0.293);
    expect(parseFormattedNumber(12.5)).toBe(12.5);
    expect(parseFormattedNumber("n/a")).toBeNull();
    expect(parseFormattedNumber("")).toBeNull();
  });
});
