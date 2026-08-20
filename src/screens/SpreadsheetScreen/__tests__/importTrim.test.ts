/**
 * Styled-blank tail trimming (importTrim.ts).
 *
 * The pathological input: a column formatted to the bottom of Excel's grid
 * writes ~1M rows of <row><c s="…"/></row> — style-only cells with custom
 * heights that must NOT reach the Univer snapshot. Content (values, formulas,
 * notes, merges) must always define the kept extent.
 *
 * Round-trip tests go through xlsx.writeBuffer → xlsx.load so the worksheet
 * under test is the same post-parse shape xlsxBytesToWorkbook sees.
 */
import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import {
  computeContentExtent,
  trimBoundsFor,
  STYLE_ONLY_COL_CAP,
  STYLE_ONLY_ROW_CAP,
  TRIM_COL_MARGIN,
  TRIM_ROW_MARGIN,
} from "../importTrim";

async function roundTrip(build: (ws: ExcelJS.Worksheet) => void): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook();
  build(wb.addWorksheet("Test"));
  const buf = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buf as ArrayBuffer);
  return wb2.getWorksheet("Test")!;
}

const GRAY_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFCCCCCC" },
};

describe("computeContentExtent", () => {
  it("tracks values and formulas, 0-indexed", async () => {
    const ws = await roundTrip((w) => {
      w.getCell("B2").value = 42;
      w.getCell("D5").value = { formula: "B2*2", result: 84 };
    });
    const e = computeContentExtent(ws);
    expect(e.contentMaxRow).toBe(4); // row 5
    expect(e.contentMaxCol).toBe(3); // col D
  });

  it("style-only cells extend anyMax but not contentMax", async () => {
    const ws = await roundTrip((w) => {
      w.getCell("A1").value = "x";
      for (let r = 2; r <= 3000; r++) {
        w.getCell(r, 1).fill = GRAY_FILL;
        w.getRow(r).height = 10.5;
      }
    });
    const e = computeContentExtent(ws);
    expect(e.contentMaxRow).toBe(0);
    expect(e.contentMaxCol).toBe(0);
    expect(e.anyMaxRow).toBe(2999);
  });

  it("a note on an otherwise empty cell counts as content", () => {
    // No round-trip here: ExcelJS's writeBuffer drops notes on valueless
    // cells, so the fixture can't survive a save. On LOAD, ExcelJS attaches
    // comments to cells present in sheetData (cell-xform reconcile), which
    // the includeEmpty walk visits — the in-memory shape matches that.
    const wb = new ExcelJS.Workbook();
    const w = wb.addWorksheet("Test");
    w.getCell("A1").value = "x";
    w.getCell("C100").note = "remember this";
    const e = computeContentExtent(w);
    expect(e.contentMaxRow).toBe(99);
    expect(e.contentMaxCol).toBe(2);
  });

  it("merged ranges count as content even when blank", async () => {
    const ws = await roundTrip((w) => {
      w.getCell("A1").value = "x";
      w.mergeCells("B200:D205");
    });
    // Mirrors xlsxBytesToWorkbook: merges are passed in pre-parsed.
    const e = computeContentExtent(ws, [{ endRow: 204, endColumn: 3 }]);
    expect(e.contentMaxRow).toBe(204);
    expect(e.contentMaxCol).toBe(3);
  });
});

describe("trimBoundsFor", () => {
  it("adds the margins to the content extent", () => {
    const b = trimBoundsFor({ contentMaxRow: 100, contentMaxCol: 20, anyMaxRow: 5000, anyMaxCol: 20 });
    expect(b.trimRow).toBe(100 + TRIM_ROW_MARGIN);
    expect(b.trimCol).toBe(20 + TRIM_COL_MARGIN);
  });

  it("caps style-only sheets instead of trimming to nothing", () => {
    const b = trimBoundsFor({ contentMaxRow: -1, contentMaxCol: -1, anyMaxRow: 1_048_406, anyMaxCol: 0 });
    expect(b.trimRow).toBe(STYLE_ONLY_ROW_CAP - 1);
    expect(b.trimCol).toBe(0);
  });

  it("keeps a small style-only sheet whole", () => {
    const b = trimBoundsFor({ contentMaxRow: -1, contentMaxCol: -1, anyMaxRow: 30, anyMaxCol: 4 });
    expect(b.trimRow).toBe(30);
    expect(b.trimCol).toBe(4);
    expect(STYLE_ONLY_COL_CAP - 1).toBeGreaterThan(4);
  });
});

describe("trim end-to-end shape (Share Tracker pathology, scaled down)", () => {
  it("drops the styled tail but keeps content and the margin band", async () => {
    const LAST_CONTENT_ROW = 50; // 0-indexed 49
    const TAIL_ROWS = 5000;
    const ws = await roundTrip((w) => {
      for (let r = 1; r <= LAST_CONTENT_ROW; r++) w.getCell(r, 2).value = r;
      for (let r = LAST_CONTENT_ROW + 1; r <= TAIL_ROWS; r++) {
        w.getCell(r, 1).fill = GRAY_FILL;
        w.getRow(r).height = 10.5;
      }
    });
    const e = computeContentExtent(ws);
    const { trimRow } = trimBoundsFor(e);
    expect(e.contentMaxRow).toBe(LAST_CONTENT_ROW - 1);
    expect(e.anyMaxRow).toBe(TAIL_ROWS - 1);
    expect(trimRow).toBe(LAST_CONTENT_ROW - 1 + TRIM_ROW_MARGIN);
    // Simulate the import walk's row filter: everything the converter would
    // keep sits at or below trimRow, so the kept row count collapses from
    // 5000 to content + margin.
    let kept = 0;
    ws.eachRow({ includeEmpty: true }, (_row, rowNumber) => {
      if (rowNumber - 1 <= trimRow) kept++;
    });
    expect(kept).toBe(trimRow + 1);
    expect(kept).toBeLessThan(TAIL_ROWS / 10);
  });
});
