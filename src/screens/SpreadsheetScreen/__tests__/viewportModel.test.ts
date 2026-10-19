import { describe, expect, it } from "vitest";
import {
  DEFAULT_COL_PX,
  DEFAULT_ROW_PX,
  HEADER_H,
  SheetGeometry,
  TILE_COLS,
  TILE_ROWS,
  TileCache,
  cellText,
  colLetters,
  colWidthPx,
  parseA1,
  rowHeightPx,
  tileKey,
  type EngineLayoutSheet,
} from "../viewportModel";
import { contextFromIndex } from "../agent/captureContext";

const sheet = (over: Partial<EngineLayoutSheet> = {}): EngineLayoutSheet => ({
  name: "S",
  hidden: false,
  color: null,
  frozenRows: 0,
  frozenColumns: 0,
  showGridLines: true,
  merges: [],
  cols: [],
  rows: [],
  cellCount: 0,
  maxRow: 0,
  maxCol: 0,
  ...over,
});

describe("SheetGeometry", () => {
  it("sizes the grid past the used range and places rows/cols by prefix sums", () => {
    const g = new SheetGeometry(sheet({ maxRow: 500, maxCol: 30 }));
    expect(g.rowCount).toBe(540);
    expect(g.colCount).toBe(38);
    expect(g.rowTop[0]).toBe(0);
    expect(g.rowTop[10]).toBe(10 * DEFAULT_ROW_PX);
    expect(g.colLeft[3]).toBe(3 * DEFAULT_COL_PX);
    expect(g.totalHeight).toBe(540 * DEFAULT_ROW_PX);
    expect(g.rowAt(0)).toBe(0);
    expect(g.rowAt(10 * DEFAULT_ROW_PX + 1)).toBe(10);
    expect(g.colAt(3 * DEFAULT_COL_PX - 1)).toBe(2);
    // Past the end clamps to the last row/col.
    expect(g.rowAt(1e9)).toBe(g.rowCount - 1);
  });

  it("applies custom widths, custom heights and hidden rows/cols", () => {
    const g = new SheetGeometry(
      sheet({
        maxRow: 10,
        maxCol: 5,
        cols: [{ min: 2, max: 3, width: 20, hidden: false }, { min: 4, max: 4, width: 8, hidden: true }],
        rows: [{ r: 3, height: 30, hidden: false }, { r: 4, height: 15, hidden: true }],
      }),
    );
    expect(g.colWidth(0)).toBe(DEFAULT_COL_PX);
    expect(g.colWidth(1)).toBe(colWidthPx(20));
    expect(g.colWidth(2)).toBe(colWidthPx(20));
    expect(g.colWidth(3)).toBe(0);
    expect(g.rowHeight(2)).toBe(rowHeightPx(30));
    expect(g.rowHeight(3)).toBe(0);
    // A hidden row takes no space: the next row starts where it would have.
    expect(g.rowTop[4]).toBe(g.rowTop[3]);
    // Hit-testing at the boundary never lands on a zero-height row.
    expect(g.rowHeight(g.rowAt(g.rowTop[4]))).toBeGreaterThan(0);
  });

  it("resolves merges to their anchor rectangle", () => {
    const g = new SheetGeometry(sheet({ maxRow: 10, maxCol: 10, merges: ["B2:D3"], frozenRows: 1, frozenColumns: 2 }));
    const m = g.mergeOf(1, 1);
    expect(m).toEqual({ r0: 1, c0: 1, r1: 2, c1: 3 });
    expect(g.mergeOf(2, 3)).toBe(m);
    expect(g.mergeOf(0, 0)).toBeUndefined();
    expect(g.frozenRows).toBe(1);
    expect(g.frozenCols).toBe(2);
  });
});

describe("TileCache", () => {
  it("stores windows as tiles, reports missing tiles, and evicts least-recently-used", () => {
    const t = new TileCache(2);
    const k0 = tileKey(0, 0, 0);
    t.putTile(k0, [{ r: 1, c: 1, v: 5, d: "5.0" }, { r: 2, c: 1, f: "=A1*2", v: 10 }]);
    expect(t.cell(0, 0, 0)).toMatchObject({ v: 5, d: "5.0" });
    expect(t.cell(0, 1, 0)).toMatchObject({ v: 10, f: "=A1*2" });
    expect(t.cell(0, 5, 5)).toBeUndefined();
    expect(t.missing(0, 0, TILE_ROWS * 2 - 1, 0, TILE_COLS - 1).map((m) => m.key)).toEqual([tileKey(0, 1, 0)]);
    t.markPending(tileKey(0, 1, 0));
    expect(t.missing(0, 0, TILE_ROWS * 2 - 1, 0, TILE_COLS - 1)).toEqual([]);
    t.putTile(tileKey(0, 1, 0), []);
    t.cell(0, 0, 0); // touch k0 so it is the most recent
    t.putTile(tileKey(0, 2, 0), []);
    expect(t.has(k0)).toBe(true);
    expect(t.has(tileKey(0, 1, 0))).toBe(false);
    expect(t.size).toBe(2);
  });

  it("applies engine values in place and takes optimistic local writes", () => {
    const t = new TileCache();
    t.putTile(tileKey(1, 0, 0), [{ r: 1, c: 1, v: 1 }]);
    t.applyValues(1, [{ row0: 0, col0: 0, v: 7 }, { row0: 0, col0: 1, v: "x" }]);
    expect(t.cell(1, 0, 0)?.v).toBe(7);
    expect(t.cell(1, 0, 1)?.v).toBe("x");
    t.setLocal(1, 0, 0, "=SUM(A1)");
    expect(t.cell(1, 0, 0)?.f).toBe("=SUM(A1)");
    t.setLocal(1, 0, 0, "12");
    expect(t.cell(1, 0, 0)).toMatchObject({ v: 12 });
    expect(t.cell(1, 0, 0)?.f).toBeUndefined();
    t.dropTile(tileKey(1, 0, 0));
    expect(t.cell(1, 0, 0)).toBeUndefined();
  });
});

describe("helpers", () => {
  it("letters, A1 and display text", () => {
    expect(colLetters(0)).toBe("A");
    expect(colLetters(25)).toBe("Z");
    expect(colLetters(26)).toBe("AA");
    expect(parseA1("BU12")).toEqual({ row0: 11, col0: 72 });
    expect(cellText({ v: 1234.5 })).toBe("1234.5");
    expect(cellText({ v: 0.5, d: "50%" })).toBe("50%");
    expect(cellText({ v: null })).toBe("");
    expect(cellText(undefined)).toBe("");
    expect(HEADER_H).toBeGreaterThan(0);
  });
});

describe("contextFromIndex", () => {
  it("turns the engine index into a stable structural context", () => {
    const ctx = contextFromIndex("/w.xlsx", {
      definedNames: [],
      sheets: [
        {
          name: "Model",
          hidden: false,
          usedRange: "A1:BU120",
          headerRow: 3,
          headers: { B: "FY2024A", C: "FY2025E" },
          rows: [
            { row: 5, label: "Revenue", values: 2, formulas: 10, bold: true, sampleCol: 10 } as any,
            { row: 6, label: null, values: 0, formulas: 0, bold: false, sampleCol: null } as any,
            { row: 7, label: "COGS", values: 0, formulas: 12, bold: false, sampleCol: 10 } as any,
          ],
          sections: [],
          formulaRuns: [],
          formulaRunsTruncated: false,
        },
        { name: "Scratch", hidden: true, usedRange: null, headerRow: null, headers: {}, rows: [], sections: [], formulaRuns: [], formulaRunsTruncated: false },
      ],
    });
    expect(ctx.sheets.map((s) => s.name)).toEqual(["Model", "Scratch"]);
    expect(ctx.sheets[0].row_count).toBe(120);
    expect(ctx.sheets[0].column_count).toBe(73);
    expect(ctx.sheets[0].used_range).toBe("A1:BU120");
    expect(ctx.sheets[0].cells_preview).toContain('headers (row 3): B="FY2024A", C="FY2025E"');
    expect(ctx.sheets[0].cells_preview).toContain('row 5: "Revenue" [2 values, 10 formulas, bold]');
    expect(ctx.sheets[0].cells_preview).not.toContain("row 6");
    expect(ctx.sheets[1].cells_preview.startsWith("(HIDDEN sheet.")).toBe(true);
    expect(ctx.sheets[1].row_count).toBe(100);
    // Same input → identical bytes (prompt cache stability).
    expect(JSON.stringify(contextFromIndex("/w.xlsx", { definedNames: [], sheets: [] }))).toBe(JSON.stringify(contextFromIndex("/w.xlsx", { definedNames: [], sheets: [] })));
  });
});
