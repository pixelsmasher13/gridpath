/**
 * Corpus validation for the full-export fallback (plan: fidelity-tests).
 *
 * Runs against real workbooks in eval/corpus (gitignored, proprietary) or
 * $GRIDPATH_XLSX_CORPUS; skips silently when neither exists. For each file:
 * ExcelJS load → export → sanitize → defined-name repair (the actual
 * fallback save pipeline), then assert
 *   1. the output has NO empty <definedName/> elements and NO invalid
 *      cfRule operators (both trip Excel's repair dialog), and
 *   2. every constant-valued defined name survives with its original text.
 */
import { beforeAll, describe, expect, it } from "vitest";
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { applyExcelJsPatches } from "../exceljsPatches";
import { repairDefinedNames } from "../definedNamePreserve";
import { sanitizeExportedPackage } from "../exportSanitize";

function corpusFiles(): string[] {
  const roots = [
    process.env.GRIDPATH_XLSX_CORPUS,
    join(__dirname, "../../../../eval/corpus"),
  ].filter((p): p is string => !!p);
  const out: string[] = [];
  for (const root of roots) {
    try {
      for (const f of readdirSync(root)) {
        if (f.toLowerCase().endsWith(".xlsx")) out.push(join(root, f));
      }
    } catch {
      // root missing — fine
    }
  }
  return out;
}

/** Decode XML entities so `&quot;TRUE&quot;` and `"TRUE"` compare equal —
 *  writers legitimately differ in escaping; Excel reads both identically. */
function xmlDecode(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

async function definedNameMap(bytes: Uint8Array): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("xl/workbook.xml")!.async("string");
  const out = new Map<string, string>();
  for (const m of xml.matchAll(
    /<definedName\b[^>]*?name="([^"]*)"[^>]*?(?:\/>|>([\s\S]*?)<\/definedName>)/g,
  )) {
    if (!out.has(m[1])) out.set(m[1], xmlDecode(m[2] ?? ""));
  }
  return out;
}

const files = corpusFiles();

describe.skipIf(files.length === 0)("corpus round-trip (full-export fallback)", () => {
  beforeAll(() => applyExcelJsPatches());

  for (const file of files) {
    it(`${file.split("/").pop()}: export is schema-valid and keeps constant names`, async () => {
      const source = new Uint8Array(readFileSync(file));
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength) as ArrayBuffer);
      let exported = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
      exported = await sanitizeExportedPackage(exported);

      const repair = await repairDefinedNames(source, exported);
      expect(repair).not.toBeNull();
      const bytes = repair!.bytes;

      const sourceNames = await definedNameMap(source);
      const outNames = await definedNameMap(bytes);

      // 1. No empty elements — Excel repair strips exactly those.
      for (const [name, text] of outNames) {
        expect(text.trim().length, `"${name}" serialized empty`).toBeGreaterThan(0);
      }

      // 1b. No schema-invalid cfRule operators anywhere in the package.
      const outZip = await JSZip.loadAsync(bytes);
      for (const part of Object.keys(outZip.files)) {
        if (!/^xl\/worksheets\/[^/]+\.xml$/.test(part)) continue;
        const xml = await outZip.file(part)!.async("string");
        expect(
          / operator="(containsBlanks|notContainsBlanks|containsErrors|notContainsErrors)"/.test(xml),
          `${part} carries an invalid cfRule operator`,
        ).toBe(false);
      }

      // 2. Constants (no sheet reference) survive verbatim.
      for (const [name, text] of sourceNames) {
        if (!text.trim() || text.includes("!")) continue;
        expect(outNames.get(name), `constant name "${name}" lost`).toBe(text);
      }
    }, 120_000);
  }
});
