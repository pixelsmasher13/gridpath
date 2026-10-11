import { describe, it, expect } from "vitest";
import { mapCell, mapSnapshot, newStyleCache, stripImplicitIntersection, styleToUniver } from "./snapshot";

describe("stripImplicitIntersection", () => {
  it("removes the engine's @ markers outside string literals", () => {
    expect(stripImplicitIntersection("=@IF(Name=\"x\",1,@INDEX(A:A,2))")).toBe("=IF(Name=\"x\",1,INDEX(A:A,2))");
    expect(stripImplicitIntersection("=@A1:A5*2")).toBe("=A1:A5*2");
    expect(stripImplicitIntersection("=@'My Sheet'!B2")).toBe("='My Sheet'!B2");
  });
  it("keeps @ inside strings and leaves plain formulas untouched", () => {
    expect(stripImplicitIntersection('="mail@example.com"&@A1')).toBe('="mail@example.com"&A1');
    expect(stripImplicitIntersection("=SUM(A1:A3)")).toBe("=SUM(A1:A3)");
  });
  it("is applied by mapCell", () => {
    const { cell } = mapCell({ r: 1, c: 1, f: "=@IF(A1,1,0)", v: 1 }, {}, newStyleCache());
    expect(cell?.f).toBe("=IF(A1,1,0)");
  });
});

describe("styleToUniver borders", () => {
  it("maps IronCalc border sides to Univer bd codes with resolved colours", () => {
    const u = styleToUniver({
      num_fmt: "General",
      fill: {},
      font: { sz: 11, name: "Calibri" },
      border: { top: { style: "thin", color: "#FF0000" }, bottom: { style: "medium" } },
      quote_prefix: false,
    } as any);
    expect(u?.bd).toEqual({ t: { s: 1, cl: { rgb: "#FF0000" } }, b: { s: 8, cl: { rgb: "#000000" } } });
  });
});

describe("mapSnapshot lazy sheets", () => {
  it("flags lazy sheets, keeps their skeleton and sizes them from the engine extent", () => {
    const sheet = (name: string, extra: Record<string, unknown>) => ({
      name, hidden: false, color: null, frozenRows: 0, frozenColumns: 0, showGridLines: true, merges: [], cols: [], rows: [], comments: [], ...extra,
    });
    const m = mapSnapshot({
      sheets: [
        sheet("Front", { cells: [{ r: 1, c: 1, v: "x" }] }) as any,
        sheet("Big", { lazy: true, cellCount: 500000, maxRow: 20000, maxCol: 110, cells: [{ r: 1, c: 1, v: "Label" }, { r: 7000, c: 1, v: "last" }] }) as any,
      ],
      styles: {},
      definedNames: [],
    }, "wb");
    expect([...m.lazySheets]).toEqual(["Big"]);
    const big = m.workbook.sheets[m.workbook.sheetOrder[1]];
    expect(big.rowCount).toBe(19999 + 100); // engine extent, not the skeleton's
    expect(big.columnCount).toBe(200);
    expect(big.cellData[0][0]).toEqual({ v: "Label" });
    expect(m.workbook.sheets[m.workbook.sheetOrder[0]].rowCount).toBe(10000);
  });
});
