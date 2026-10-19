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

  it("snapshot formulas carry exactly one leading '='", () => {
    const wb = engine.Workbook.open(bytes);
    const snap = JSON.parse(wb.snapshot());
    const formulas = snap.sheets.flatMap((s: any) => s.cells.filter((c: any) => c.f).map((c: any) => c.f as string));
    expect(formulas.length).toBeGreaterThan(0);
    expect(formulas.every((f: string) => f.startsWith("=") && !f.startsWith("=="))).toBe(true);
  });

  it("writes a formula, evaluates, and reads the result back", () => {
    const wb = engine.Workbook.open(bytes);
    wb.setInput(0, 500, 1, "=1+2");
    wb.evaluate();
    expect(JSON.parse(wb.cell(0, 500, 1)).value).toBe(3);
  });

  it("treats date serial 0 and the phantom 29 Feb 1900 the way Excel does", () => {
    // Blank or zero date cells feed YEAR/MONTH/EOMONTH in real models; Excel
    // reads serial 0 as "0 January 1900" rather than erroring.
    const wb = engine.Workbook.open(bytes);
    const cases: Array<[string, number]> = [
      ["=YEAR(0)", 1900],
      ["=MONTH(0)", 1],
      ["=DAY(0)", 0],
      ["=EOMONTH(0,0)", 31],
      ["=EOMONTH(DATE(YEAR(0),3,15),0)", 91],
      ["=DAY(60)", 29],
      ["=DATE(1900,1,1)", 1],
      ["=DATE(1900,3,1)", 61],
      ["=WEEKDAY(1)", 1],
    ];
    cases.forEach(([f], i) => wb.setInput(0, 600 + i, 1, f));
    wb.setInput(0, 650, 1, "=YEAR(-1)");
    wb.evaluate();
    cases.forEach(([f, want], i) => expect([f, JSON.parse(wb.cell(0, 600 + i, 1)).value]).toEqual([f, want]));
    expect(JSON.parse(wb.cell(0, 650, 1)).value).toBe("#NUM!");
  });

  it("evaluateEdited recalculates only the dirty closure and reports it", () => {
    const wb = engine.Workbook.open(bytes);
    wb.setInput(0, 700, 1, "5");
    wb.setInput(0, 701, 1, "=A700*2");
    wb.setInput(0, 702, 1, "=A701+1");
    let d = JSON.parse(wb.evaluateEdited());
    const at = (r: number) => d.cells.find((c: any) => c.sheet === 0 && c.r === r && c.c === 1);
    expect(d.full).toBe(false);
    expect(at(701).v).toBe(10);
    expect(at(702).v).toBe(11);

    wb.setInput(0, 700, 1, "6");
    d = JSON.parse(wb.evaluateEdited());
    expect(d.full).toBe(false);
    expect(d.cells.map((c: any) => c.r).sort()).toEqual([700, 701, 702]);
    expect(at(702).v).toBe(13);
    expect(JSON.parse(wb.cell(0, 702, 1)).value).toBe(13);

    // Clearing reports the cleared cell (empty) and its dependents.
    wb.clearContents(0, 700, 1, 1, 1);
    d = JSON.parse(wb.evaluateEdited());
    expect(at(700).v ?? null).toBeNull();
    expect(at(702).v).toBe(1);

    // Structural edits invalidate the dependency map: full recalc, full snapshot.
    wb.setInput(0, 700, 1, "6");
    wb.insertRows(0, 1, 1);
    d = JSON.parse(wb.evaluateEdited());
    expect(d.full).toBe(true);
    expect(JSON.parse(wb.cell(0, 703, 1)).value).toBe(13);
    // And the map is live again afterwards.
    wb.setInput(0, 701, 1, "7");
    d = JSON.parse(wb.evaluateEdited());
    expect(d.full).toBe(false);
    expect(JSON.parse(wb.cell(0, 703, 1)).value).toBe(15);
  });

  it("incremental recalculation matches a full evaluate on the same edits", () => {
    const a = engine.Workbook.open(bytes);
    const b = engine.Workbook.open(bytes);
    const snap = JSON.parse(a.snapshot());
    // Overwrite a handful of numeric inputs across sheets.
    const inputs: Array<[number, number, number, number]> = [];
    snap.sheets.forEach((sh: any, si: number) => {
      for (const c of sh.cells) {
        if (!c.f && typeof c.v === "number" && c.v !== 0 && inputs.length < 12 && inputs.filter((x) => x[0] === si).length < 3) inputs.push([si, c.r, c.c, c.v]);
      }
    });
    expect(inputs.length).toBeGreaterThan(3);
    for (const [si, r, c, v] of inputs) {
      a.setInput(si, r, c, String(v * 1.37));
      b.setInput(si, r, c, String(v * 1.37));
    }
    const d = JSON.parse(a.evaluateEdited());
    expect(d.full).toBe(false);
    b.evaluate();
    const va = JSON.parse(a.snapshot()), vb = JSON.parse(b.snapshot());
    expect(va).toEqual(vb);
    // The delta reported every cell whose value moved.
    const before = new Map<string, unknown>();
    for (const sh of snap.sheets) for (const c of sh.cells) before.set(`${sh.name}!${c.r},${c.c}`, c.v);
    const moved: string[] = [];
    for (const sh of vb.sheets) for (const c of sh.cells) if (before.get(`${sh.name}!${c.r},${c.c}`) !== c.v) moved.push(`${sh.name}!${c.r},${c.c}`);
    const reported = new Set(d.cells.map((c: any) => `${vb.sheets[c.sheet].name}!${c.r},${c.c}`));
    expect(moved.filter((k) => !reported.has(k))).toEqual([]);
    expect(moved.length).toBeGreaterThan(inputs.length);
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
