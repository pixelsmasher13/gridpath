/**
 * End-to-end import conversion (xlsxImport.ts): real xlsx bytes through
 * ExcelJS into the Univer snapshot. Focuses on what the UniverGrid worker
 * split depends on:
 *   - styled-blank tails are trimmed out of cellData / rowData / rowCount
 *   - external-workbook formulas are pinned and surfaced as cloneable pairs
 *   - parseXlsxWorkbook's in-process fallback (no Worker in node) returns
 *     the full ParsedWorkbook shape including the retained ExcelJS object
 */
import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { parseXlsxWorkbook, xlsxBytesToWorkbook, externalPinsByWorkbook } from "../xlsxImport";
import { TRIM_ROW_MARGIN } from "../importTrim";

async function buildXlsx(build: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  build(wb);
  return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

const GRAY_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFCCCCCC" },
};

describe("xlsxBytesToWorkbook trim integration", () => {
  it("drops the styled tail from cellData, rowData and rowCount", async () => {
    const LAST_CONTENT_ROW = 40;
    const TAIL_ROWS = 8000;
    const bytes = await buildXlsx((wb) => {
      const ws = wb.addWorksheet("Tail");
      for (let r = 1; r <= LAST_CONTENT_ROW; r++) ws.getCell(r, 2).value = r;
      for (let r = LAST_CONTENT_ROW + 1; r <= TAIL_ROWS; r++) {
        ws.getCell(r, 1).fill = GRAY_FILL;
        ws.getRow(r).height = 10.5;
      }
    });
    const { univerData } = await xlsxBytesToWorkbook(bytes);
    const sheet = univerData.sheets["sheet_Tail"];
    expect(sheet).toBeDefined();
    const trimRow = LAST_CONTENT_ROW - 1 + TRIM_ROW_MARGIN;
    const maxCellRow = Math.max(...Object.keys(sheet.cellData).map(Number));
    expect(maxCellRow).toBeLessThanOrEqual(trimRow);
    const rowDataKeys = Object.keys(sheet.rowData).map(Number);
    expect(Math.max(...rowDataKeys)).toBeLessThanOrEqual(trimRow);
    // rowCount collapses to the default floor instead of tracking the tail.
    expect(sheet.rowCount).toBe(10000);
    // Content and the margin band of styled blanks survive.
    expect(sheet.cellData[LAST_CONTENT_ROW - 1][1].v).toBe(LAST_CONTENT_ROW);
    expect(sheet.cellData[LAST_CONTENT_ROW][0].s).toBeTruthy();
  });

  it("keeps styled blanks on small sheets untouched", async () => {
    const bytes = await buildXlsx((wb) => {
      const ws = wb.addWorksheet("Small");
      ws.getCell("A1").value = "header";
      ws.getCell("B3").fill = GRAY_FILL;
    });
    const { univerData } = await xlsxBytesToWorkbook(bytes);
    const sheet = univerData.sheets["sheet_Small"];
    expect(sheet.cellData[2][1].s).toBeTruthy();
  });
});

describe("external formula pinning", () => {
  it("strips external-workbook formulas and records cloneable pins", async () => {
    const bytes = await buildXlsx((wb) => {
      const ws = wb.addWorksheet("Pins");
      ws.getCell("A1").value = { formula: "'[2]Other'!A1*2", result: 84 } as any;
      ws.getCell("A2").value = { formula: "SUM(1,2)", result: 3 } as any;
    });
    const { excelJs, univerData } = await xlsxBytesToWorkbook(bytes);
    const pins = externalPinsByWorkbook.get(excelJs)!;
    expect(pins.size).toBe(1);
    const pinned = univerData.sheets["sheet_Pins"].cellData[0][0];
    expect(pinned.f).toBeUndefined(); // formula stripped, cached value kept
    expect(pinned.v).toBe(84);
    const normal = univerData.sheets["sheet_Pins"].cellData[1][0];
    expect(normal.f).toBe("=SUM(1,2)");
  });
});

describe("parseXlsxWorkbook in-process fallback", () => {
  it("returns the full ParsedWorkbook shape when Worker is unavailable", async () => {
    const bytes = await buildXlsx((wb) => {
      const ws = wb.addWorksheet("Data");
      ws.getCell("A1").value = 1;
      ws.getCell("B1").value = { formula: "'[2]Ext'!B9", result: 7 } as any;
    });
    expect(typeof Worker).toBe("undefined"); // node test env — fallback path
    const parsed = await parseXlsxWorkbook(bytes);
    expect(parsed.excelJs).not.toBeNull(); // in-process parse retains ExcelJS
    expect(parsed.univerData.sheets["sheet_Data"].cellData[0][0].v).toBe(1);
    expect(parsed.pins).toEqual([["Data 0,1", { f: "='[2]Ext'!B9", v: 7 }]]);
    expect(Array.isArray(parsed.images)).toBe(true);
  });
});

describe("comment parts at a non-standard path (openpyxl layout)", () => {
  it("opens instead of crashing in ExcelJS's worksheet reconcile", async () => {
    const JSZip = (await import("jszip")).default;
    const standard = await buildXlsx((wb) => {
      const ws = wb.addWorksheet("Notes");
      ws.getCell("A1").value = 42;
      ws.getCell("A1").note = "a note";
      ws.getCell("B2").value = "kept";
    });
    // Rewrite the package the way openpyxl lays it out: comments under
    // xl/comments/, the shapes under a different name, absolute targets.
    const zip = await JSZip.loadAsync(standard);
    const move = async (from: string, to: string) => {
      const data = await zip.file(from)!.async("uint8array");
      zip.remove(from);
      zip.file(to, data);
    };
    await move("xl/comments1.xml", "xl/comments/comment1.xml");
    await move("xl/drawings/vmlDrawing1.vml", "xl/drawings/commentsDrawing1.vml");
    const relsPath = "xl/worksheets/_rels/sheet1.xml.rels";
    const rels = (await zip.file(relsPath)!.async("string"))
      .replace("../comments1.xml", "/xl/comments/comment1.xml")
      .replace("../drawings/vmlDrawing1.vml", "/xl/drawings/commentsDrawing1.vml");
    expect(rels).toContain("/xl/comments/comment1.xml");
    zip.file(relsPath, rels);
    const ct = (await zip.file("[Content_Types].xml")!.async("string")).replace(
      "/xl/comments1.xml",
      "/xl/comments/comment1.xml",
    );
    zip.file("[Content_Types].xml", ct);
    const bytes = await zip.generateAsync({ type: "uint8array" });

    const { univerData } = await xlsxBytesToWorkbook(bytes);
    const sheet = univerData.sheets["sheet_Notes"];
    expect(sheet.cellData[0][0].v).toBe(42);
    expect(sheet.cellData[1][1].v).toBe("kept");
  });
});
