/**
 * Engine smoke tests: open → snapshot → edit → evaluate → patch, and the
 * fidelity property the whole product rests on — every package part the
 * edit didn't touch comes back byte-identical.
 *
 * Requires `npm run build:engine` first (engine/ is git-ignored).
 */
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { describe, it, expect } from "vitest";
import JSZip from "jszip";

const require = createRequire(import.meta.url);
const ENGINE = path.resolve(import.meta.dirname, "../engine/gridpath_engine.js");
const built = fs.existsSync(ENGINE);
const engine = built ? require(ENGINE) : null;
const FIXTURES = path.resolve(import.meta.dirname, "../../../eval/fixtures");

async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const [name, f] of Object.entries(zip.files)) {
    if (!f.dir) out.set(name, await f.async("uint8array"));
  }
  return out;
}

function same(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

describe.skipIf(!built)("engine", () => {
  const bytes = fs.readFileSync(path.join(FIXTURES, "rich-model.xlsx"));

  it("opens, snapshots and reads a cell", () => {
    const wb = engine.Workbook.open(bytes);
    const names: string[] = wb.sheetNames();
    expect(names.length).toBeGreaterThan(0);
    const snap = JSON.parse(wb.snapshot());
    expect(snap.sheets[0].name).toBe(names[0]);
    expect(snap.sheets[0].cells.length).toBeGreaterThan(0);
    const first = snap.sheets[0].cells[0];
    const cell = JSON.parse(wb.cell(0, first.r, first.c));
    expect(cell.value).toEqual(first.v);
  });

  it("writes a formula, evaluates, and reads the result back", () => {
    const wb = engine.Workbook.open(bytes);
    wb.setInput(0, 500, 1, "=1+2");
    wb.evaluate();
    expect(JSON.parse(wb.cell(0, 500, 1)).value).toBe(3);
  });

  it("shifts references when rows are inserted", () => {
    const wb = engine.Workbook.open(bytes);
    wb.setInput(0, 500, 1, "7");
    wb.setInput(0, 501, 1, "=A500*2");
    wb.insertRows(0, 1, 2);
    wb.evaluate();
    expect(JSON.parse(wb.cell(0, 503, 1)).content).toBe("=A502*2");
    expect(JSON.parse(wb.cell(0, 503, 1)).value).toBe(14);
  });

  it("patch leaves every untouched part byte-identical", async () => {
    const wb = engine.Workbook.open(bytes);
    const sheet = wb.sheetNames()[0];
    const patched: Uint8Array = engine.patch(
      bytes,
      JSON.stringify({ version: 1, sheets: [{ name: sheet, cells: [{ r: 0, c: 30, v: { t: "s", s: "probe" } }] }] }),
    );
    const before = await parts(bytes);
    const after = await parts(patched);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    const changed = [...before.keys()].filter((k) => !same(before.get(k)!, after.get(k)!));
    // Only the edited sheet, plus workbook.xml when the patcher sets fullCalcOnLoad.
    expect(changed.every((k) => k === "xl/workbook.xml" || /^xl\/worksheets\/sheet\d+\.xml$/.test(k))).toBe(true);
    expect(changed.some((k) => k.startsWith("xl/worksheets/"))).toBe(true);
  });
});
