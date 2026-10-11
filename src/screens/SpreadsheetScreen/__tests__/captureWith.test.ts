import { describe, expect, it } from "vitest";
import { captureWorkbookContextWith, type CellReader } from "../agent/captureContext";

/**
 * The viewport grid captures the agent's opening context from a snapshot it
 * assembles from engine windows plus a reader bound to that data. The
 * format must be the Univer-path format (same memo, same base+delta).
 */
function fixture(edit?: { r: number; c: number; v: unknown; f?: string }) {
  const cellData: Record<string, Record<string, any>> = {
    0: { 0: { v: "Item" }, 1: { v: "FY24" }, 2: { v: "FY25" } },
    1: { 0: { v: "Revenue" }, 1: { v: 100, s: 1 }, 2: { v: 110, f: "=B2*1.1", s: 1 } },
    2: { 0: { v: "Growth" }, 2: { v: 0.1, f: "=C2/B2-1", s: 2 } },
  };
  if (edit) (cellData[edit.r] ??= {})[edit.c] = { v: edit.v, ...(edit.f ? { f: edit.f } : {}) };
  const snapshot = {
    sheetOrder: ["s0", "s1"],
    sheets: {
      s0: { name: "Model", hidden: 0, rowCount: 100, columnCount: 26, cellData },
      s1: { name: "Scratch", hidden: 1, rowCount: 100, columnCount: 26, cellData: { 4: { 0: { v: "cache rows" } } } },
    },
  };
  const reader: CellReader = {
    getCell: (sheet, r, c) => {
      const cell = sheet === "Model" ? cellData[r]?.[c] : undefined;
      if (!cell) return null;
      const out: any = { value: cell.v ?? null, formula: cell.f ?? null };
      if (cell.s === 1) out.display = `${cell.v}.0`;
      return out;
    },
    getCellFormat: (sheet, r, c) => {
      const s = sheet === "Model" ? cellData[r]?.[c]?.s : undefined;
      return s === 1 ? { number_format: "#,##0.0", bold: true } : s === 2 ? { number_format: "0.0%" } : null;
    },
    getExternalPin: () => null,
  };
  return { snapshot, reader };
}

describe("captureWorkbookContextWith", () => {
  it("renders the Univer-format preview from a snapshot and reader", () => {
    const { snapshot, reader } = fixture();
    const ctx = captureWorkbookContextWith("/t1.xlsx", snapshot, reader, "v1");
    expect(ctx.sheets.map((s) => s.name)).toEqual(["Model", "Scratch"]);
    const m = ctx.sheets[0];
    expect(m.used_range).toBe("A1:C3");
    expect(m.cells_preview).toContain("Style conventions");
    expect(m.cells_preview).toContain("B2: bold, fmt #,##0.0");
    expect(m.cells_preview).toContain('A1 = "Item"');
    expect(m.cells_preview).toContain("B2 = 100 (100.0)  [bold, fmt #,##0.0]");
    expect(m.cells_preview).toContain("C2 = =B2*1.1  → 110 (110.0)  [bold, fmt #,##0.0]");
    expect(m.cells_preview).toContain("C3 = =C2/B2-1  → 0.1  [fmt 0.0%]");
    expect(ctx.sheets[1].cells_preview.startsWith("(This sheet is HIDDEN")).toBe(true);
    expect(ctx.sheets[1].cells_preview).toContain('row 5: "cache rows"');
  });

  it("memoizes on the content key and ships a delta on later changes", () => {
    const base = fixture();
    const ctx1 = captureWorkbookContextWith("/t2.xlsx", base.snapshot, base.reader, "v1");
    const again = captureWorkbookContextWith("/t2.xlsx", base.snapshot, base.reader, "v1");
    expect(again.sheets).toBe(ctx1.sheets);
    expect(again.delta).toBeUndefined();
    // An edit far outside the previewed cells rides in the delta; the base is kept.
    const edited = fixture({ r: 2000, c: 3, v: 2, f: "=1+1" });
    const ctx2 = captureWorkbookContextWith("/t2.xlsx", edited.snapshot, edited.reader, "v2");
    expect(ctx2.sheets).toBe(ctx1.sheets);
    expect(ctx2.delta).toContain("## Sheet: Model");
    expect(ctx2.delta).toContain("D2001 = =1+1  → 2");
    // A sheet-set change forces a fresh base.
    const renamed = fixture();
    renamed.snapshot.sheets.s1.name = "Cache";
    const ctx3 = captureWorkbookContextWith("/t2.xlsx", renamed.snapshot, renamed.reader, "v3");
    expect(ctx3.sheets).not.toBe(ctx1.sheets);
    expect(ctx3.delta).toBeUndefined();
    expect(ctx3.sheets[1].name).toBe("Cache");
  });
});
