import { describe, expect, it } from "vitest";
import { expandCopy, type CopyGridReader } from "../agent/copyRange";
import { interpretToolCall, parseColumnList, parseRowList } from "../agent/toolToMutation";
import type { CellFormat } from "../types";

describe("column/row list parsing", () => {
  it("handles letters, commas, and colon runs", () => {
    expect(parseColumnList("A")).toEqual([0]);
    expect(parseColumnList("B,AO")).toEqual([1, 40]);
    expect(parseColumnList("C:F")).toEqual([2, 3, 4, 5]);
    expect(parseColumnList("A,C:E,C")).toEqual([0, 2, 3, 4]);
  });
  it("handles rows, commas, and colon runs (1-indexed → 0-indexed)", () => {
    expect(parseRowList("1")).toEqual([0]);
    expect(parseRowList("4,5,6")).toEqual([3, 4, 5]);
    expect(parseRowList("3:5,9")).toEqual([2, 3, 4, 8]);
  });
});

/** Grid stub: sparse map keyed "sheet!row,col". */
function makeGrid(
  cells: Record<string, { value?: any; formula?: string; format?: CellFormat }>,
): CopyGridReader {
  return {
    getCell: (sheet, row, col) => {
      const c = cells[`${sheet}!${row},${col}`];
      if (!c || (c.value === undefined && !c.formula)) return null;
      return { value: c.value ?? null, formula: c.formula ?? null };
    },
    getCellFormat: (sheet, row, col) =>
      cells[`${sheet}!${row},${col}`]?.format ?? {},
  };
}

const setCells = (muts: any[]) => muts.filter((m) => m.type === "set_cell");
const formats = (muts: any[]) => muts.filter((m) => m.type === "set_format");

describe("copy_range interpretation", () => {
  it("parses valid args and defaults dest_sheet + mode", () => {
    const r = interpretToolCall("copy_range", {
      sheet: "Model",
      source: "B5:B7",
      dest: "C5",
    });
    expect(r).toEqual({
      kind: "copy",
      sheet: "Model",
      source: "B5:B7",
      dest_sheet: "Model",
      dest: "C5",
      mode: "all",
    });
  });

  it("rejects malformed ranges", () => {
    expect(
      interpretToolCall("copy_range", { sheet: "S", source: "nope", dest: "A1" }).kind,
    ).toBe("ignored");
    expect(
      interpretToolCall("copy_range", { sheet: "S", source: "A1:B2:C3", dest: "A1" }).kind,
    ).toBe("ignored");
    expect(
      interpretToolCall("copy_range", { sheet: "S", source: "A1", dest: "" }).kind,
    ).toBe("ignored");
  });
});

describe("expandCopy", () => {
  it("shifts relative refs by the paste offset, keeps $ anchors", () => {
    // Copy column B (rows 5-7, 0-indexed 4-6) one column right.
    const grid = makeGrid({
      "Model!4,1": { formula: "=A5*(1+$B$2)", value: 105 },
      "Model!5,1": { formula: "=B5*2", value: 210 },
      "Model!6,1": { value: 42 },
    });
    const r = expandCopy(grid, {
      sheet: "Model",
      source: "B5:B7",
      dest_sheet: "Model",
      dest: "C5",
      mode: "all",
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const sc = setCells(r.mutations);
    expect(sc).toHaveLength(3);
    expect(sc[0].address).toEqual({ sheet: "Model", row: 4, col: 2 });
    expect(sc[0].new_formula).toBe("=B5*(1+$B$2)");
    expect(sc[1].new_formula).toBe("=C5*2");
    expect(sc[2].new_formula).toBeNull();
    expect(sc[2].new_value).toBe(42);
    expect(r.dest_range).toBe("C5:C7");
  });

  it("mode=values pastes evaluated values without formulas or formats", () => {
    const grid = makeGrid({
      "S!0,0": { formula: "=X1+1", value: 7, format: { bold: true } },
    });
    const r = expandCopy(grid, {
      sheet: "S",
      source: "A1",
      dest_sheet: "S",
      dest: "B1",
      mode: "values",
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.mutations).toHaveLength(1);
    const m: any = r.mutations[0];
    expect(m.type).toBe("set_cell");
    expect(m.new_value).toBe(7);
    expect(m.new_formula).toBeNull();
  });

  it("mode=formats copies formatting only, grouped by identical style", () => {
    const grid = makeGrid({
      "S!0,0": { value: 1, format: { bold: true, number_format: "0.0%" } },
      "S!1,0": { value: 2, format: { bold: true, number_format: "0.0%" } },
      "S!2,0": { value: 3, format: { font_color: "#0000FF" } },
    });
    const r = expandCopy(grid, {
      sheet: "S",
      source: "A1:A3",
      dest_sheet: "S",
      dest: "D1",
      mode: "formats",
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(setCells(r.mutations)).toHaveLength(0);
    const fm = formats(r.mutations);
    expect(fm).toHaveLength(2);
    const grouped = fm.find((m: any) => m.cells.length === 2);
    expect(grouped.new_format).toEqual({ bold: true, number_format: "0.0%" });
    expect(grouped.cells).toEqual([
      { row: 0, col: 3 },
      { row: 1, col: 3 },
    ]);
  });

  it("copies across sheets, shifting refs but not sheet-qualified ones", () => {
    const grid = makeGrid({
      "Src!0,0": { formula: "=A2+Data!B1", value: 1 },
    });
    const r = expandCopy(grid, {
      sheet: "Src",
      source: "A1",
      dest_sheet: "Dst",
      dest: "C3",
      mode: "all",
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const m: any = r.mutations[0];
    expect(m.address.sheet).toBe("Dst");
    // Both refs shift by (+2,+2) — Excel shifts relative refs regardless of
    // sheet qualification; only the SHEET part stays pinned.
    expect(m.new_formula).toBe("=C4+Data!D3");
  });

  it("blank source cells preserve the destination (no clear mutations)", () => {
    const grid = makeGrid({
      "S!0,0": { value: 1 },
      // A2 empty, A3 has content
      "S!2,0": { value: 3 },
    });
    const r = expandCopy(grid, {
      sheet: "S",
      source: "A1:A3",
      dest_sheet: "S",
      dest: "B1",
      mode: "all",
    });
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(setCells(r.mutations)).toHaveLength(2);
  });

  it("fails loudly when a shifted ref would leave the sheet", () => {
    const grid = makeGrid({
      "S!4,1": { formula: "=A5", value: 1 }, // B5 = =A5; copying left → column -1
    });
    const r = expandCopy(grid, {
      sheet: "S",
      source: "B5",
      dest_sheet: "S",
      dest: "A5",
      mode: "all",
    });
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error).toBe("ref_out_of_bounds");
    expect(r.ref_errors?.length).toBe(1);
  });

  it("rejects oversized ranges, no-op copies, and all-empty sources", () => {
    const grid = makeGrid({ "S!0,0": { value: 1 } });
    expect(
      (expandCopy(grid, { sheet: "S", source: "A1:AZ200", dest_sheet: "S", dest: "BA1", mode: "all" }) as any).error,
    ).toBe("range_too_large");
    expect(
      (expandCopy(grid, { sheet: "S", source: "A1:A3", dest_sheet: "S", dest: "A1", mode: "all" }) as any).error,
    ).toBe("noop");
    expect(
      (expandCopy(grid, { sheet: "S", source: "F10:F12", dest_sheet: "S", dest: "G10", mode: "all" }) as any).error,
    ).toBe("empty_source");
  });
});
