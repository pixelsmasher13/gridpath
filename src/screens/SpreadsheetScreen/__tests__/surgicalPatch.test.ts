import { describe, expect, it } from "vitest";
import {
  buildCellBaseline,
  buildWorkbookPatch,
  describePatchFallback,
  shiftFormulaA1,
  type PatchBuildResult,
} from "../surgicalPatch";

/** Minimal Univer-shaped snapshot. cells = { "r,c": {v?, f?, si?, t?} } */
function snapshot(sheets: Record<string, Record<string, any>>) {
  const out: any = { sheetOrder: [], sheets: {} };
  let i = 0;
  for (const [name, cells] of Object.entries(sheets)) {
    const id = `sheet-${i++}`;
    out.sheetOrder.push(id);
    const cellData: any = {};
    for (const [key, cell] of Object.entries(cells)) {
      const [r, c] = key.split(",");
      (cellData[r] ??= {})[c] = cell;
    }
    out.sheets[id] = { name, cellData };
  }
  return out;
}

function baselineOf(sheets: Record<string, Record<string, any>>) {
  return buildCellBaseline(snapshot(sheets));
}

function expectOk(res: PatchBuildResult): any {
  expect(res.ok, res.ok ? "" : `${(res as any).reason}: ${(res as any).detail}`).toBe(true);
  return (res as any).patch;
}

describe("buildWorkbookPatch", () => {
  it("returns an empty patch when nothing changed", () => {
    const cells = { "0,0": { v: 1 }, "1,2": { f: "A1*2", v: 2 } };
    const res = buildWorkbookPatch(snapshot({ S: cells }), baselineOf({ S: cells }));
    const patch = expectOk(res);
    expect(patch).toEqual({ version: 1 });
  });

  it("diffs changed values and cleared cells", () => {
    const base = baselineOf({ S: { "0,0": { v: 1 }, "0,1": { v: "x" } } });
    const res = buildWorkbookPatch(snapshot({ S: { "0,0": { v: 5 } } }), base);
    const patch = expectOk(res);
    const cells = patch.sheets[0].cells;
    expect(cells).toContainEqual({ r: 0, c: 0, v: { t: "n", n: 5 } });
    expect(cells).toContainEqual({ r: 0, c: 1, clear: true });
  });

  it("writes formula cells with their evaluated value cached", () => {
    const base = baselineOf({ S: { "0,0": { v: 1 } } });
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { v: 1 }, "0,1": { f: "=A1*2", v: 2 } } }),
      base,
    );
    const patch = expectOk(res);
    expect(patch.sheets[0].cells).toEqual([
      { r: 0, c: 1, f: "A1*2", v: { t: "n", n: 2 } },
    ]);
  });

  it("treats evaluated-value drift on an unchanged formula as a change", () => {
    // Same formula text, new computed result (an upstream input changed, or
    // the baseline came from a file saved without cached values): the fresh
    // value must reach the file or reopening would show a stale cache.
    const res = buildWorkbookPatch(
      snapshot({ S: { "1,2": { f: "A1*2", v: 10 } } }),
      baselineOf({ S: { "1,2": { f: "A1*2", v: 2 } } }),
    );
    const patch = expectOk(res);
    expect(patch.sheets[0].cells).toEqual([
      { r: 1, c: 2, f: "A1*2", v: { t: "n", n: 10 } },
    ]);
  });

  it("writes the formula bare when its value can't be encoded", () => {
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { f: "COMPLEX(1,2)", v: { rich: "obj" } } } }),
      baselineOf({ S: {} }),
    );
    const patch = expectOk(res);
    expect(patch.sheets[0].cells).toEqual([{ r: 0, c: 0, f: "COMPLEX(1,2)" }]);
  });

  it("reports unsupported_value for value shapes the patch can't carry", () => {
    const base = baselineOf({ S: {} });
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { v: { rich: "text" } } } }),
      base,
    );
    expect(res).toMatchObject({ ok: false, reason: "unsupported_value" });
  });

  it("reports sheet_structure for a live sheet with no baseline and no create op", () => {
    const res = buildWorkbookPatch(snapshot({ Ghost: { "0,0": { v: 1 } } }), baselineOf({ S: {} }));
    expect(res).toMatchObject({ ok: false, reason: "sheet_structure" });
  });

  it("reports sheet_structure for a baseline sheet missing without a delete op", () => {
    const res = buildWorkbookPatch(snapshot({ S: {} }), baselineOf({ S: {}, Gone: {} }));
    expect(res).toMatchObject({ ok: false, reason: "sheet_structure" });
  });

  it("emits sheet create ops and diffs the new sheet against an empty baseline", () => {
    const res = buildWorkbookPatch(
      snapshot({ S: {}, Notes: { "0,0": { v: "hi" } } }),
      baselineOf({ S: {} }),
      { sheetOps: [{ kind: "create", name: "Notes", tabColor: "#112233" }] },
    );
    const patch = expectOk(res);
    expect(patch.sheetOps).toEqual([{ op: "create", name: "Notes", tabColor: "#112233" }]);
    expect(patch.sheets).toEqual([
      { name: "Notes", cells: [{ r: 0, c: 0, v: { t: "s", s: "hi" } }] },
    ]);
  });

  it("follows renames without spurious diffs", () => {
    const res = buildWorkbookPatch(
      snapshot({ NewName: { "0,0": { v: 1 } } }),
      baselineOf({ Old: { "0,0": { v: 1 } } }),
      { sheetOps: [{ kind: "rename", oldName: "Old", newName: "NewName" }] },
    );
    const patch = expectOk(res);
    expect(patch.sheetOps).toEqual([{ op: "rename", oldName: "Old", newName: "NewName" }]);
    expect(patch.sheets).toBeUndefined();
  });

  it("tolerates deleted sheets when a delete op explains them", () => {
    const res = buildWorkbookPatch(
      snapshot({ S: {} }),
      baselineOf({ S: {}, Gone: { "0,0": { v: 9 } } }),
      { sheetOps: [{ kind: "delete", name: "Gone" }] },
    );
    const patch = expectOk(res);
    expect(patch.sheetOps).toEqual([{ op: "delete", name: "Gone" }]);
  });

  it("shifts the baseline for row inserts so unmoved content doesn't diff", () => {
    // Baseline rows 0 and 1; two rows inserted before row 1 → live rows 0 and 3.
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { v: "top" }, "3,0": { v: "bottom" }, "1,0": { v: "new" } } }),
      baselineOf({ S: { "0,0": { v: "top" }, "1,0": { v: "bottom" } } }),
      { rowColOps: [{ kind: "insertRows", sheet: "S", before: 1, count: 2 }] },
    );
    const patch = expectOk(res);
    expect(patch.rowColOps).toEqual([{ op: "insertRows", sheet: "S", before: 1, count: 2 }]);
    // Only the genuinely new cell is in the diff.
    expect(patch.sheets).toEqual([
      { name: "S", cells: [{ r: 1, c: 0, v: { t: "s", s: "new" } }] },
    ]);
  });

  it("drops deleted rows from the baseline instead of clearing shifted cells", () => {
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { v: "top" }, "1,0": { v: "last" } } }),
      baselineOf({ S: { "0,0": { v: "top" }, "1,0": { v: "mid" }, "2,0": { v: "last" } } }),
      { rowColOps: [{ kind: "deleteRows", sheet: "S", start: 1, count: 1 }] },
    );
    const patch = expectOk(res);
    expect(patch.rowColOps).toEqual([{ op: "deleteRows", sheet: "S", start: 1, count: 1 }]);
    expect(patch.sheets).toBeUndefined();
  });

  it("shifts baseline columns for column ops", () => {
    const res = buildWorkbookPatch(
      snapshot({ S: { "0,0": { v: 1 }, "0,2": { v: 2 } } }),
      baselineOf({ S: { "0,0": { v: 1 }, "0,1": { v: 2 } } }),
      { rowColOps: [{ kind: "insertColumns", sheet: "S", before: 1, count: 1 }] },
    );
    const patch = expectOk(res);
    expect(patch.sheets).toBeUndefined();
  });

  it("resolves Univer shared formulas from the group master", () => {
    // Master B1 (=A1*2) with si=7; member B2 carries only si.
    const live = snapshot({
      S: {
        "0,1": { f: "A1*2", si: 7, v: 2 },
        "1,1": { si: 7, v: 4 },
      },
    });
    // Baseline had the same formulas materialized — no diff expected.
    const base = baselineOf({ S: { "0,1": { f: "A1*2", v: 2 }, "1,1": { f: "A2*2", v: 4 } } });
    const patch = expectOk(buildWorkbookPatch(live, base));
    expect(patch.sheets).toBeUndefined();
  });

  it("reports shared_formula when a group has no master", () => {
    const live = snapshot({ S: { "1,1": { si: 7, v: 4 } } });
    const res = buildWorkbookPatch(live, baselineOf({ S: {} }));
    expect(res).toMatchObject({ ok: false, reason: "shared_formula" });
  });

  it("has copy for every fallback reason", () => {
    for (const reason of [
      "sheet_structure",
      "row_col_structure",
      "shared_formula",
      "unsupported_value",
      "missing_baseline",
    ] as const) {
      expect(describePatchFallback({ reason }).length).toBeGreaterThan(10);
    }
  });
});

describe("shiftFormulaA1", () => {
  it("shifts relative refs and keeps absolutes", () => {
    expect(shiftFormulaA1("A1+B2", 1, 0)).toBe("A2+B3");
    expect(shiftFormulaA1("$A$1+B2", 5, 3)).toBe("$A$1+E7");
    expect(shiftFormulaA1("SUM(A1:A10)", 0, 2)).toBe("SUM(C1:C10)");
    expect(shiftFormulaA1("$A1+A$1", 1, 1)).toBe("$A2+B$1");
  });

  it("skips strings, identifiers and function names", () => {
    expect(shiftFormulaA1('IF(A1="B2",LOG10(A1),MY_A1)', 1, 0)).toBe(
      'IF(A2="B2",LOG10(A2),MY_A1)',
    );
    expect(shiftFormulaA1("'Bal A1 Sheet'!B2+Plain!C3", 1, 1)).toBe(
      "'Bal A1 Sheet'!C3+Plain!D4",
    );
  });

  it("returns null when a ref would leave the sheet", () => {
    expect(shiftFormulaA1("A1", -1, 0)).toBeNull();
    expect(shiftFormulaA1("XFD1", 0, 1)).toBeNull();
  });
});
