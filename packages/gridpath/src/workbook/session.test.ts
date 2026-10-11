/**
 * Headless session: open → read → mutate → readback → save, and the
 * fidelity property on the saved bytes.
 */
import fs from "node:fs";
import path from "node:path";
import { describe, it, expect } from "vitest";
import JSZip from "jszip";
import { WorkbookSession, a1Of, loadEngine } from "./session";
import { buildWorkbookIndex, describeWorkbookPayload, findRowsInIndex } from "../core/workbookIndex";

const FIXTURES = path.resolve(import.meta.dirname, "../../../../eval/fixtures");
let built = true;
try {
  loadEngine();
} catch {
  built = false;
}

async function parts(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = await JSZip.loadAsync(bytes);
  const out = new Map<string, Uint8Array>();
  for (const [name, f] of Object.entries(zip.files)) if (!f.dir) out.set(name, await f.async("uint8array"));
  return out;
}
const same = (a: Uint8Array, b: Uint8Array) => a.length === b.length && a.every((x, i) => x === b[i]);

describe.skipIf(!built)("WorkbookSession", () => {
  const file = path.join(FIXTURES, "rich-model.xlsx");
  const open = () => WorkbookSession.open(file, fs.readFileSync(file));

  it("patches the cached snapshot in place to exactly what a full rebuild gives", () => {
    const s = open();
    const snap = s.snapshot();
    const sheet = s.sheetNames()[0];
    const sid = snap.sheetOrder[0];
    // Pick a numeric input and a formula that will move with it.
    let input: [number, number, number] | null = null;
    for (const [r, row] of Object.entries(snap.sheets[sid].cellData)) {
      for (const [c, cell] of Object.entries(row)) if (!cell.f && typeof cell.v === "number" && cell.v !== 0) { input = [Number(r), Number(c), cell.v]; break; }
      if (input) break;
    }
    expect(input).not.toBeNull();
    const [r, c, v] = input!;
    s.apply({
      id: "b1", prompt: "t", justification: "t",
      mutations: [
        { type: "set_cell", address: { sheet, row: r, col: c }, old_value: v, new_value: v * 3 },
        { type: "set_cell", address: { sheet, row: 2000, col: 0 }, old_value: null, new_value: null, new_formula: `=${a1Of(r, c)}*2` },
        { type: "set_format", sheet, cells: [{ row: 2000, col: 0 }], new_format: { bold: true, number_format: "0.00" } } as any,
      ],
    });
    const patched = structuredClone(s.snapshot());
    const patchedFormat = s.reader().getCellFormat(sheet, 2000, 0);
    (s as any).mapped = null; // force the full re-serialise path
    const rebuilt = s.snapshot();
    expect(patched).toEqual(rebuilt);
    expect(patchedFormat).toEqual(s.reader().getCellFormat(sheet, 2000, 0));
    expect(s.cell(sheet, 2000, 0).value).toBe(v * 6);
  });

  it("snapshot feeds the workbook index", () => {
    const s = open();
    const snap = s.snapshot();
    expect(snap.sheetOrder.length).toBe(s.sheetNames().length);
    const index = buildWorkbookIndex(snap);
    const described = describeWorkbookPayload(index, null) as any;
    expect(JSON.stringify(described)).toContain(s.sheetNames()[0]);
    const hits = findRowsInIndex(index, "revenue", null, 10);
    expect(hits).toBeTruthy();
  });

  it("readRange returns evaluated values with display text", () => {
    const s = open();
    const sheet = s.sheetNames()[0];
    const { cells } = s.readRange(sheet, "A1:F12");
    expect(cells.length).toBeGreaterThan(0);
    expect(cells.every((c) => /^[A-Z]+\d+$/.test(c.cell))).toBe(true);
  });

  it("set_cell + insert_rows land in the model with shifted references", () => {
    const s = open();
    const sheet = s.sheetNames()[0];
    s.apply({
      id: "b1",
      prompt: "test",
      justification: "test",
      mutations: [
        { type: "set_cell", address: { sheet, row: 499, col: 0 }, old_value: null, new_value: 21 },
        { type: "set_cell", address: { sheet, row: 500, col: 0 }, old_value: null, new_value: null, new_formula: "=A500*2" },
      ],
    });
    expect(s.cell(sheet, 500, 0).value).toBe(42);
    s.apply({ id: "b2", prompt: "t", justification: "t", mutations: [{ type: "insert_rows", sheet, before: 0, count: 3 }] });
    expect(s.cell(sheet, 503, 0)).toMatchObject({ value: 42, formula: "=A503*2" });
  });

  it("set_format reaches the model and the reader", () => {
    const s = open();
    const sheet = s.sheetNames()[0];
    s.apply({
      id: "b1",
      prompt: "t",
      justification: "t",
      mutations: [
        {
          type: "set_format",
          sheet,
          range: "A1",
          cells: [{ row: 0, col: 0 }],
          old_format: [],
          new_format: { bold: true, number_format: "0.0%", background_color: "#1F4E79" },
        },
      ],
    });
    const fmt = s.reader().getCellFormat(sheet, 0, 0);
    expect(fmt).toMatchObject({ bold: true, number_format: "0.0%", background_color: "#1F4E79" });
  });

  it("buildSave patches only the edited sheet; everything else byte-identical", async () => {
    const s = open();
    const sheet = s.sheetNames()[0];
    s.apply({
      id: "b1",
      prompt: "t",
      justification: "t",
      mutations: [{ type: "set_cell", address: { sheet, row: 0, col: 40 }, old_value: null, new_value: "probe" }],
    });
    const saved = s.buildSave();
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    const before = await parts(s.originalBytes);
    const after = await parts(saved.bytes);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    const changed = [...before.keys()].filter((k) => !same(before.get(k)!, after.get(k)!));
    expect(changed.every((k) => k === "xl/workbook.xml" || /^xl\/worksheets\/sheet\d+\.xml$/.test(k))).toBe(true);
    // And the edit is really in the file: reopen the saved bytes.
    const again = WorkbookSession.open(file, saved.bytes);
    expect(again.cell(sheet, 0, 40).value).toBe("probe");
  });
});
