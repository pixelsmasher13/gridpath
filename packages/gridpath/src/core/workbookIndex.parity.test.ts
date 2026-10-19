import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { WorkbookSession, loadEngine } from "../workbook/session";
import { buildWorkbookIndex, type SheetIndex } from "./workbookIndex";

/**
 * Parity between the two index builders: the TypeScript one over the
 * engine's Univer-shaped snapshot (what the desktop grid path and older
 * engines use) and the Rust one inside the engine
 * (`gridpath_engine_core::index`, served to describe_workbook / find_rows
 * / query_rows). Both read the same model, so every structural fact —
 * used range, header row and headers, row labels and counts, bold flags,
 * sections, formula-run spans — must agree. Formula TEXT of a run may
 * differ only by the `@` implicit-intersection marker the snapshot mapper
 * strips.
 */

const FIXTURES = [
  path.resolve(import.meta.dirname, "../../../../eval/fixtures/rich-model.xlsx"),
  path.resolve(import.meta.dirname, "../../../../eval/corpus/a0aac6d9-f059-40fc-8d1f-0aa50d848cda.xlsx"),
].filter((f) => fs.existsSync(f));

let built = true;
try {
  loadEngine();
} catch {
  built = false;
}

const stripAt = (f: string) => f.replace(/@/g, "");

function shape(s: SheetIndex) {
  return {
    name: s.name,
    hidden: s.hidden,
    usedRange: s.usedRange,
    headerRow: s.headerRow,
    headers: s.headers,
    rows: s.rows,
    sections: s.sections,
    runSpans: s.formulaRuns.map((r) => [r.row, r.startCol, r.endCol]),
    runFormulas: s.formulaRuns.map((r) => stripAt(r.formula)),
    formulaRunsTruncated: s.formulaRunsTruncated,
  };
}

describe.skipIf(!built || FIXTURES.length === 0)("workbook index: Rust and TypeScript builders agree", () => {
  for (const fixture of FIXTURES) {
    it(path.basename(fixture), () => {
      const s = WorkbookSession.open(fixture, new Uint8Array(fs.readFileSync(fixture)));
      const ts = buildWorkbookIndex(s.snapshot());
      const rs = s.workbookIndex();
      expect(ts.sheets.length).toBeGreaterThan(0);
      expect(ts.sheets.reduce((n, x) => n + x.rows.length, 0)).toBeGreaterThan(10);
      expect(ts.sheets.some((x) => x.headerRow !== null)).toBe(true);
      expect(ts.sheets.some((x) => x.formulaRuns.length > 0)).toBe(true);
      expect(rs.sheets.map((x) => x.name)).toEqual(ts.sheets.map((x) => x.name));
      for (let i = 0; i < ts.sheets.length; i++) {
        expect(shape(rs.sheets[i])).toEqual(shape(ts.sheets[i]));
      }
      expect(rs.definedNames.length).toBeGreaterThanOrEqual(0);
    });
  }
});
