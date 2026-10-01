import { describe, expect, it } from "vitest";
import { reorderWorksheetChildren } from "./exportSanitize";

const wrap = (body: string) =>
  `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</worksheet>`;

describe("reorderWorksheetChildren", () => {
  it("moves legacyDrawing before tableParts (ExcelJS comment+table bug)", () => {
    const input = wrap(
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
        '<pageMargins left="0.7"/>' +
        '<tableParts count="1"><tablePart r:id="rId2"/></tableParts>' +
        '<legacyDrawing r:id="rId3"/>',
    );
    const out = reorderWorksheetChildren(input);
    expect(out.indexOf("<legacyDrawing")).toBeLessThan(out.indexOf("<tableParts"));
    expect(out.indexOf("<pageMargins")).toBeLessThan(out.indexOf("<legacyDrawing"));
    expect(out.replace(/<[^>]+>/g, "")).toBe(input.replace(/<[^>]+>/g, ""));
  });

  it("returns input unchanged when already in schema order", () => {
    const input = wrap('<sheetPr/><dimension ref="A1"/><sheetData/><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><tableParts/>');
    expect(reorderWorksheetChildren(input)).toBe(input);
  });

  it("keeps repeatable and unknown elements in relative order", () => {
    const input = wrap(
      '<sheetData/>' +
        '<conditionalFormatting sqref="A1"><cfRule type="expression"/></conditionalFormatting>' +
        '<conditionalFormatting sqref="B1"><cfRule type="expression"/></conditionalFormatting>' +
        '<x:unknownThing/>' +
        '<cols><col min="1" max="1"/></cols>',
    );
    const out = reorderWorksheetChildren(input);
    expect(out.indexOf("<cols>")).toBeLessThan(out.indexOf("<sheetData"));
    expect(out.indexOf('sqref="A1"')).toBeLessThan(out.indexOf('sqref="B1"'));
    expect(out.indexOf("<x:unknownThing")).toBeGreaterThan(out.indexOf('sqref="B1"'));
  });

  it("bails on unbalanced XML", () => {
    const input = wrap("<sheetData><row></sheetData><tableParts/>");
    expect(reorderWorksheetChildren(input)).toBe(input);
  });

  it("does not split on nested elements sharing a name with a top-level one", () => {
    const input = wrap('<tableParts><tablePart r:id="x"/></tableParts><drawing r:id="d"/><extLst><ext><drawing/></ext></extLst>');
    const out = reorderWorksheetChildren(input);
    expect(out.indexOf('<drawing r:id="d"/>')).toBeLessThan(out.indexOf("<tableParts"));
    expect(out).toContain("<extLst><ext><drawing/></ext></extLst>");
  });
});
