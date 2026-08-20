import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { auditDefinedNameLoss, repairDefinedNames } from "../definedNamePreserve";

async function workbookZip(definedNamesXml: string, extra: Record<string, string> = {}) {
  const zip = new JSZip();
  zip.file(
    "xl/workbook.xml",
    `<?xml version="1.0"?><workbook><sheets><sheet name="Model" sheetId="1" r:id="rId1"/></sheets>${definedNamesXml}<calcPr/></workbook>`,
  );
  for (const [k, v] of Object.entries(extra)) zip.file(k, v);
  return zip.generateAsync({ type: "uint8array" });
}

async function namesOf(bytes: Uint8Array): Promise<Record<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("xl/workbook.xml")!.async("string");
  const out: Record<string, string> = {};
  for (const m of xml.matchAll(/<definedName\b[^>]*?name="([^"]*)"[^>]*?(?:\/>|>([\s\S]*?)<\/definedName>)/g)) {
    out[m[1]] = m[2] ?? "";
  }
  return out;
}

// The NFLX/Canalyst scenario: constant-valued vendor flags. ExcelJS's
// parser cannot decode "TRUE" as a range, so the round-trip either drops
// the name or emits an empty element (schema-invalid → Excel repair).
const SOURCE = `<definedNames>` +
  `<definedName name="AA.FreeCashFlow">"TRUE"</definedName>` +
  `<definedName name="WS.HasWorkingRange" localSheetId="0">"TRUE"</definedName>` +
  `<definedName name="Revenue">Model!$B$5:$B$12</definedName>` +
  `<definedName name="Multi">Model!$A$1,Model!$B$1,"flag"</definedName>` +
  `</definedNames>`;

describe("repairDefinedNames", () => {
  it("restores constants that the export gutted to empty elements", async () => {
    const source = await workbookZip(SOURCE);
    const exported = await workbookZip(
      `<definedNames>` +
        `<definedName name="AA.FreeCashFlow"></definedName>` +
        `<definedName name="WS.HasWorkingRange" localSheetId="0"/>` +
        `<definedName name="Revenue">Model!$B$5:$B$12</definedName>` +
        `</definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported);
    expect(repair).not.toBeNull();
    expect(repair!.restored.sort()).toEqual(["AA.FreeCashFlow", "WS.HasWorkingRange"]);
    const names = await namesOf(repair!.bytes);
    expect(names["AA.FreeCashFlow"]).toBe('"TRUE"');
    expect(names["WS.HasWorkingRange"]).toBe('"TRUE"');
    expect(names["Revenue"]).toBe("Model!$B$5:$B$12");
  });

  it("grafts back dropped constants but never sheet-referencing names", async () => {
    const source = await workbookZip(SOURCE);
    // Stock-ExcelJS behavior for <200 names: unparseable names vanish.
    const exported = await workbookZip(
      `<definedNames><definedName name="Revenue">Model!$B$5:$B$12</definedName></definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.grafted.sort()).toEqual(["AA.FreeCashFlow", "WS.HasWorkingRange"]);
    const names = await namesOf(repair!.bytes);
    expect(names["AA.FreeCashFlow"]).toBe('"TRUE"');
    // "Multi" references sheets — not provably independent, NOT grafted.
    expect(names["Multi"]).toBeUndefined();
  });

  it("restores partially stripped multi-segment names (subset check)", async () => {
    const source = await workbookZip(SOURCE);
    // The parser kept the decodable segments and dropped the "flag" constant.
    const exported = await workbookZip(
      `<definedNames>` +
        `<definedName name="AA.FreeCashFlow">"TRUE"</definedName>` +
        `<definedName name="WS.HasWorkingRange" localSheetId="0">"TRUE"</definedName>` +
        `<definedName name="Revenue">Model!$B$5:$B$12</definedName>` +
        `<definedName name="Multi">Model!$A$1,Model!$B$1</definedName>` +
        `</definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.restored).toEqual(["Multi"]);
    const names = await namesOf(repair!.bytes);
    expect(names["Multi"]).toBe('Model!$A$1,Model!$B$1,"flag"');
  });

  it("drops empty entries that never existed in the source (schema validity)", async () => {
    const source = await workbookZip(
      `<definedNames><definedName name="Revenue">Model!$B$5:$B$12</definedName></definedNames>`,
    );
    const exported = await workbookZip(
      `<definedNames>` +
        `<definedName name="Ghost"></definedName>` +
        `<definedName name="Revenue">Model!$B$5:$B$12</definedName>` +
        `</definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.dropped).toEqual(["Ghost"]);
    const names = await namesOf(repair!.bytes);
    expect("Ghost" in names).toBe(false);
  });

  it("leaves names the user redefined this session alone", async () => {
    const source = await workbookZip(SOURCE);
    const exported = await workbookZip(
      `<definedNames><definedName name="Revenue">Model!$C$5:$C$12</definedName></definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported, new Set(["Revenue"]));
    const names = await namesOf(repair!.bytes);
    expect(names["Revenue"]).toBe("Model!$C$5:$C$12");
  });

  it("does not rewrite intentional changes (not a subset of the original)", async () => {
    const source = await workbookZip(SOURCE);
    const exported = await workbookZip(
      `<definedNames>` +
        `<definedName name="AA.FreeCashFlow">"TRUE"</definedName>` +
        `<definedName name="WS.HasWorkingRange" localSheetId="0">"TRUE"</definedName>` +
        `<definedName name="Revenue">Model!$B$5:$B$12</definedName>` +
        `<definedName name="Multi">Model!$Z$9</definedName>` +
        `</definedNames>`,
    );
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.restored).toEqual([]);
    const names = await namesOf(repair!.bytes);
    expect(names["Multi"]).toBe("Model!$Z$9");
  });

  it("inserts a definedNames block when the export has none", async () => {
    const source = await workbookZip(
      `<definedNames><definedName name="AA.Flag">"TRUE"</definedName></definedNames>`,
    );
    const exported = await workbookZip("");
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.grafted).toEqual(["AA.Flag"]);
    const names = await namesOf(repair!.bytes);
    expect(names["AA.Flag"]).toBe('"TRUE"');
  });

  it("returns identical bytes when nothing needs repair", async () => {
    const source = await workbookZip(SOURCE);
    const exported = await workbookZip(SOURCE);
    const repair = await repairDefinedNames(source, exported);
    expect(repair!.restored).toEqual([]);
    expect(repair!.grafted).toEqual([]);
    expect(repair!.dropped).toEqual([]);
    expect(repair!.bytes).toBe(exported);
  });
});

describe("auditDefinedNameLoss", () => {
  it("names what is still missing after export", async () => {
    const source = await workbookZip(SOURCE);
    const exported = await workbookZip(
      `<definedNames><definedName name="Revenue">Model!$B$5:$B$12</definedName></definedNames>`,
    );
    const missing = await auditDefinedNameLoss(source, exported);
    expect(missing!.sort()).toEqual(["AA.FreeCashFlow", "Multi", "WS.HasWorkingRange"]);
  });

  it("reports nothing for a faithful export", async () => {
    const bytes = await workbookZip(SOURCE);
    expect(await auditDefinedNameLoss(bytes, bytes)).toEqual([]);
  });
});
