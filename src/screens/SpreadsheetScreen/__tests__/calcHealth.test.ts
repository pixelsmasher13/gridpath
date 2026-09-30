import { describe, expect, it } from "vitest";
import {
  computeCalcHealth,
  fileSavedIfDivergent,
  formatCalcHealthLine,
  valuesDiverge,
} from "../agent/calcHealth";
import type { BaselineCell } from "../surgicalPatch";

const saved = (f: string, v: any): BaselineCell => ({ f, v });

describe("fileSavedIfDivergent", () => {
  it("flags an engine zero against the file's cached value", () => {
    expect(fileSavedIfDivergent(0, "=SUM(B5:B8)", saved("SUM(B5:B8)", 596.6))).toBe(596.6);
  });

  it("matches formulas up to '=' prefix, spacing and case", () => {
    expect(fileSavedIfDivergent(0, "= sum(b5:b8) ", saved("SUM(B5:B8)", 42))).toBe(42);
  });

  it("never flags a cell whose formula changed this session", () => {
    expect(fileSavedIfDivergent(0, "=SUM(B5:B9)", saved("SUM(B5:B8)", 596.6))).toBeUndefined();
  });

  it("never flags non-formula cells or cells the file cached no value for", () => {
    expect(fileSavedIfDivergent(5, null, { v: 7 })).toBeUndefined();
    expect(fileSavedIfDivergent(5, "=A1", { f: "A1" })).toBeUndefined();
    expect(fileSavedIfDivergent(5, "=A1", null)).toBeUndefined();
  });

  it("tolerates engine rounding noise", () => {
    expect(fileSavedIfDivergent(24367.000001, "=B71/B67*1000", saved("B71/B67*1000", 24367))).toBeUndefined();
  });

  it("flags a live error string against a saved number", () => {
    expect(fileSavedIfDivergent("#VALUE!", "=A1*B1", saved("A1*B1", 3356))).toBe(3356);
  });

  it("flags a live null (engine produced nothing) against a saved value", () => {
    expect(fileSavedIfDivergent(null, "=INDIRECT(A1)", saved("INDIRECT(A1)", 1042))).toBe(1042);
  });

  it("flags a live empty string the same as null", () => {
    expect(fileSavedIfDivergent("", "=INDIRECT(A1)", saved("INDIRECT(A1)", 1042))).toBe(1042);
  });

  it("does not flag a computed boolean that disagrees with Excel", () => {
    expect(fileSavedIfDivergent(true, "=A1>B1", saved("A1>B1", true))).toBeUndefined();
    expect(fileSavedIfDivergent(1, "=A1>B1", saved("A1>B1", true))).toBeUndefined();
    expect(fileSavedIfDivergent(false, "=A1>B1", saved("A1>B1", true))).toBeUndefined();
  });

  it("does not flag two plausible numbers that merely disagree", () => {
    expect(fileSavedIfDivergent(0.15, "=EBIT/Rev", saved("EBIT/Rev", 0.17))).toBeUndefined();
    expect(fileSavedIfDivergent(0.17, "=EBIT/Rev", saved("EBIT/Rev", 0.191))).toBeUndefined();
  });

  it("does not flag two plausible strings that merely disagree", () => {
    expect(fileSavedIfDivergent("FY2026E ", '=T(A1)', saved("T(A1)", "FY2026E"))).toBeUndefined();
    expect(fileSavedIfDivergent("FY2027E", '=T(A1)', saved("T(A1)", "FY2026E"))).toBeUndefined();
  });

  it("does not attach file_saved when Excel's cache is also unusable", () => {
    expect(fileSavedIfDivergent("#VALUE!", "=A1*B1", saved("A1*B1", "#REF!"))).toBeUndefined();
    expect(fileSavedIfDivergent(null, "=A1", saved("A1", null as unknown as number))).toBeUndefined();
  });
});

describe("valuesDiverge", () => {
  it("uses relative tolerance on large magnitudes", () => {
    expect(valuesDiverge(1_304_194.0001, 1_304_194)).toBe(false);
    expect(valuesDiverge(1_304_194, 1_312_800)).toBe(true);
  });
});

describe("computeCalcHealth", () => {
  const baseline = new Map<string, Map<string, BaselineCell>>([
    [
      "Model",
      new Map<string, BaselineCell>([
        ["0,0", saved("SUM(B1:B2)", 100)], // diverges (live 0)
        ["1,0", saved("A1*2", 200)], // agrees
        ["2,0", saved("A1*3", 300)], // formula edited in-session → skipped
        ["3,0", { v: 42 }], // literal → never checked
        ["4,0", saved("A1*5", 500)], // cleared in-session → skipped
        ["5,0", saved("EBIT/Rev", 0.17)], // live 0.15 — plausible vs plausible, not broken
      ]),
    ],
    ["Gone", new Map([["0,0", saved("A1", 1)]])], // sheet renamed → skipped
  ]);
  const snapshot = {
    sheets: {
      s1: {
        name: "Model",
        cellData: {
          0: { 0: { f: "=SUM(B1:B2)", v: 0 } },
          1: { 0: { f: "=A1*2", v: 200 } },
          2: { 0: { f: "=A1*30", v: 9000 } },
          3: { 0: { v: 42 } },
          5: { 0: { f: "=EBIT/Rev", v: 0.15 } },
        },
      },
    },
  };

  it("counts only comparable cells and attributes divergence per sheet", () => {
    const h = computeCalcHealth(snapshot, baseline);
    expect(h.checked).toBe(3); // rows 0, 1, and 5
    expect(h.divergent).toBe(1);
    expect(h.bySheet).toEqual([{ sheet: "Model", divergent: 1 }]);
    expect(h.truncated).toBe(false);
  });

  it("stops at the cap and reports truncation", () => {
    const h = computeCalcHealth(snapshot, baseline, 1);
    expect(h.checked).toBe(1);
    expect(h.truncated).toBe(true);
  });
});

describe("formatCalcHealthLine", () => {
  it("is empty when nothing diverges", () => {
    expect(formatCalcHealthLine(null)).toBe("");
    expect(
      formatCalcHealthLine({ checked: 10, divergent: 0, bySheet: [], truncated: false }),
    ).toBe("");
  });

  it("names counts and worst sheets", () => {
    const line = formatCalcHealthLine({
      checked: 48210,
      divergent: 312,
      bySheet: [
        { sheet: "Model", divergent: 290 },
        { sheet: "ModelSummary", divergent: 22 },
      ],
      truncated: false,
    });
    expect(line).toContain("312 of 48210");
    expect(line).toContain("Model: 290");
    expect(line).toContain("file_saved");
    expect(line).toContain("not a defect");
  });
});
