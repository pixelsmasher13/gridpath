import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { sanitizeExportedPackage } from "../exportSanitize";

async function zipOf(parts: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(parts)) zip.file(name, content);
  return zip.generateAsync({ type: "uint8array" });
}

async function readPart(bytes: Uint8Array, name: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file(name)!.async("string");
}

describe("sanitizeExportedPackage", () => {
  it("strips the invalid cfRule operators ExcelJS writes", async () => {
    const bytes = await zipOf({
      "xl/worksheets/sheet1.xml":
        `<worksheet><conditionalFormatting sqref="A1:B2">` +
        `<cfRule type="containsErrors" dxfId="1" priority="9" operator="containsErrors"><formula>ISERROR(A1)</formula></cfRule>` +
        `<cfRule type="notContainsBlanks" dxfId="2" priority="8" operator="notContainsBlanks"/>` +
        `</conditionalFormatting></worksheet>`,
      "xl/worksheets/sheet2.xml":
        `<worksheet><conditionalFormatting sqref="C1">` +
        `<cfRule type="cellIs" dxfId="3" priority="1" operator="equal"><formula>0</formula></cfRule>` +
        `</conditionalFormatting></worksheet>`,
    });
    const out = await sanitizeExportedPackage(bytes);
    const s1 = await readPart(out, "xl/worksheets/sheet1.xml");
    expect(s1).toContain('<cfRule type="containsErrors" dxfId="1" priority="9">');
    expect(s1).toContain('<cfRule type="notContainsBlanks" dxfId="2" priority="8"/>');
    expect(s1).not.toContain('operator="containsErrors"');
    // Legal operators untouched.
    const s2 = await readPart(out, "xl/worksheets/sheet2.xml");
    expect(s2).toContain('operator="equal"');
  });

  it("returns the input bytes unchanged when nothing needs fixing", async () => {
    const bytes = await zipOf({
      "xl/worksheets/sheet1.xml": `<worksheet><sheetData/></worksheet>`,
    });
    expect(await sanitizeExportedPackage(bytes)).toBe(bytes);
  });
});
