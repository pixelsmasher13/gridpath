import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import {
  auditExportLoss,
  describeFidelityRisks,
  describeMissingParts,
  gateFallbackSave,
  scanFidelityRisks,
  type FidelityRisks,
} from "../fidelityScan";

async function zipOf(parts: Record<string, string>): Promise<Uint8Array> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(parts)) zip.file(name, content);
  return zip.generateAsync({ type: "uint8array" });
}

const PLAIN = {
  "[Content_Types].xml": "<Types/>",
  "xl/workbook.xml": "<workbook/>",
  "xl/worksheets/sheet1.xml": "<worksheet/>",
  "xl/styles.xml": "<styleSheet/>",
};

const RICH = {
  ...PLAIN,
  "xl/charts/chart1.xml": "<c:chartSpace/>",
  "xl/charts/chart2.xml": "<c:chartSpace/>",
  "xl/comments1.xml":
    '<comments><commentList><comment ref="A1"/><comment ref="B2"/><comment ref="C3"/></commentList></comments>',
  "xl/externalLinks/externalLink1.xml": "<externalLink/>",
  "xl/printerSettings/printerSettings1.bin": "bin",
  "customXml/item1.xml": "<x/>",
};

describe("scanFidelityRisks", () => {
  it("finds nothing in a plain workbook", async () => {
    const risks = await scanFidelityRisks(await zipOf(PLAIN));
    expect(risks).not.toBeNull();
    expect(risks!.total).toBe(0);
  });

  it("counts charts, individual comments, links, and other at-risk parts", async () => {
    const risks = await scanFidelityRisks(await zipOf(RICH));
    expect(risks).toMatchObject({
      charts: 2,
      comments: 3,
      externalLinks: 1,
      printerSettings: 1,
      customXml: 1,
    });
    expect(risks!.total).toBe(8);
    expect(describeFidelityRisks(risks!)).toBe(
      "2 charts, 3 comments, 1 external workbook link, 1 custom XML part, 1 printer-settings part",
    );
  });
});

describe("gateFallbackSave", () => {
  const rich = { total: 3 } as FidelityRisks;
  const plain = { total: 0 } as FidelityRisks;

  it("exports silently for untitled and plain workbooks", () => {
    expect(gateFallbackSave({ isUntitled: true, isSaveAs: false, risks: null })).toBe("export");
    expect(gateFallbackSave({ isUntitled: false, isSaveAs: false, risks: plain })).toBe("export");
    // Scan failure on an existing file: treat as plain rather than lock the
    // user out of saving entirely.
    expect(gateFallbackSave({ isUntitled: false, isSaveAs: false, risks: null })).toBe("export");
  });

  it("never overwrites an at-risk file in place", () => {
    expect(gateFallbackSave({ isUntitled: false, isSaveAs: false, risks: rich })).toBe("block");
  });

  it("allows an at-risk copy with a warning", () => {
    expect(gateFallbackSave({ isUntitled: false, isSaveAs: true, risks: rich })).toBe(
      "export_with_warning",
    );
  });
});

describe("auditExportLoss", () => {
  it("names exactly the parts the export dropped", async () => {
    const source = await zipOf(RICH);
    const exported = await zipOf(PLAIN);
    const missing = await auditExportLoss(source, exported);
    expect(missing).toEqual(
      expect.arrayContaining(["xl/charts/chart1.xml", "xl/comments1.xml", "customXml/item1.xml"]),
    );
    expect(missing!.length).toBe(6);
    expect(describeMissingParts(missing!)).toBe(
      "2 chart parts, 1 comment part, 1 external-link part, 1 printer-settings part, 1 custom XML part",
    );
  });

  it("reports no loss for an identical package", async () => {
    const bytes = await zipOf(RICH);
    expect(await auditExportLoss(bytes, bytes)).toEqual([]);
  });
});
