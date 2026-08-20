import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { buildWorkbookIndex, findRowsInIndex } from "../agent/workbookIndex";

/**
 * Fixture test: run the index builder over a REAL vendor analyst model
 * (12 sheets, Income Statement 278×120) and assert the structure it infers
 * is sane. Assertions are deliberately shape-level, not value-exact — the
 * corpus file is an input we don't control, and the point is "does the
 * heuristic stack survive contact with a real model", not pixel parity.
 *
 * The adapter converts ExcelJS's parse into the snapshot subset the builder
 * reads (cellData v/f + inline styles). It is test-only: production builds
 * the index from Univer's own snapshot.
 */

const FIXTURE = fileURLToPath(
  new URL("../../../../eval/corpus/a0aac6d9-f059-40fc-8d1f-0aa50d848cda.xlsx", import.meta.url),
);

async function loadFixtureSnapshot(): Promise<any> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(FIXTURE);
  const sheets: Record<string, any> = {};
  const sheetOrder: string[] = [];
  for (const ws of wb.worksheets) {
    const id = `s${ws.id}`;
    sheetOrder.push(id);
    const cellData: Record<number, Record<number, any>> = {};
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const out: any = {};
        const raw: any = cell.value;
        if (raw !== null && raw !== undefined) {
          if (typeof raw === "object") {
            if (typeof raw.formula === "string") out.f = `=${raw.formula}`;
            if ("result" in raw && raw.result !== undefined && typeof raw.result !== "object") {
              out.v = raw.result;
            }
            if (Array.isArray(raw.richText)) {
              out.v = raw.richText.map((r: any) => r?.text ?? "").join("");
            }
            if (raw instanceof Date) out.v = raw.toISOString().slice(0, 10);
          } else {
            out.v = raw;
          }
        }
        if (out.v === undefined && out.f === undefined) return;
        const style: any = {};
        if (cell.font?.bold) style.bl = 1;
        if (typeof cell.numFmt === "string" && cell.numFmt) style.n = { pattern: cell.numFmt };
        if (Object.keys(style).length > 0) out.s = style;
        (cellData[rowNumber - 1] ??= {})[colNumber - 1] = out;
      });
    });
    sheets[id] = {
      name: ws.name,
      cellData,
      ...(ws.state === "hidden" || ws.state === "veryHidden" ? { hidden: 1 } : {}),
    };
  }
  return { sheetOrder, sheets, styles: {} };
}

describe("workbook index on a real analyst model", () => {
  it("produces a sane structural index", async () => {
    const snapshot = await loadFixtureSnapshot();
    const start = performance.now();
    const idx = buildWorkbookIndex(snapshot);
    const buildMs = performance.now() - start;

    // Budget check: the whole point is that this is cheap enough to
    // rebuild lazily per content change.
    expect(buildMs).toBeLessThan(2000);

    expect(idx.sheets.length).toBe(12);
    expect(idx.sheets.find((s) => s.name === "__FDSCACHE__")?.hidden).toBe(true);

    const income = idx.sheets.find((s) => s.name === "Income Statement")!;
    expect(income).toBeDefined();
    expect(income.usedRange).toMatch(/^A1:[A-Z]+\d+$/);
    expect(income.rows.length).toBeGreaterThan(50);

    // Sections: a real 278-row statement should decompose into a readable
    // table of contents — more than one lump, fewer than one per row.
    expect(income.sections.length).toBeGreaterThan(1);
    expect(income.sections.length).toBeLessThan(income.rows.length / 2);
    for (const sec of income.sections) {
      expect(sec.startRow).toBeLessThanOrEqual(sec.endRow);
    }
    expect(income.sections.some((s) => s.title !== null)).toBe(true);

    // Formula runs must compress: a 100+-column model stores far fewer run
    // entries than formula cells.
    const formulaCells = income.rows.reduce((n, r) => n + r.formulas, 0);
    expect(formulaCells).toBeGreaterThan(1000);
    expect(income.formulaRuns.length).toBeLessThan(formulaCells / 3);
    expect(income.formulaRunsTruncated).toBe(false);

    // find_rows: "revenue" must locate labeled rows on the statement sheets.
    const { matches } = findRowsInIndex(idx, "revenue", null, 50);
    expect(matches.length).toBeGreaterThan(0);
    expect(matches.some((m) => m.sheet === "Income Statement")).toBe(true);
    for (const m of matches) {
      expect(m.label.toLowerCase()).toContain("rev");
    }
  });
});
