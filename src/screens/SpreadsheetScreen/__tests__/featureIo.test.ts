import { describe, expect, it } from "vitest";
import {
  CF_RESOURCE_NAME,
  DV_RESOURCE_NAME,
  NOTE_RESOURCE_NAME,
  applyFeatureMirror,
  coalesceAddressesToRanges,
  excelCfToUniverRules,
  excelDvToUniverRules,
  featureDriftDetail,
  notesToUniverResource,
  noteTextFromExcelJs,
  parseA1Range,
  parseIterativeCalc,
  perSheetFeatureState,
  rangeToA1,
  univerCfToExcelCf,
  univerDvToExcelDvModel,
  univerNotesToList,
} from "../featureIo";

describe("A1 helpers", () => {
  it("parses plain, anchored and multi-cell refs", () => {
    expect(parseA1Range("B2")).toEqual({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 });
    expect(parseA1Range("$A$1:$C$10")).toEqual({ startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 });
    expect(parseA1Range("nope!")).toBeNull();
  });

  it("handles full-column refs", () => {
    const r = parseA1Range("A:B");
    expect(r?.startColumn).toBe(0);
    expect(r?.endColumn).toBe(1);
    expect(r?.startRow).toBe(0);
    expect(r?.endRow).toBe(1_048_575);
  });

  it("round-trips through rangeToA1", () => {
    expect(rangeToA1(parseA1Range("C3:D5")!)).toBe("C3:D5");
    expect(rangeToA1(parseA1Range("B2")!)).toBe("B2");
    expect(rangeToA1(parseA1Range("AA10:AB12")!)).toBe("AA10:AB12");
  });
});

describe("full-column range clamping", () => {
  const bounds = { maxRow: 800, maxCol: 20 };

  it("clamps whole-column CF refs to the used area plus margin", () => {
    const { rules } = excelCfToUniverRules(
      [{ ref: "A:A", rules: [{ type: "duplicateValues", priority: 1, style: {} }] }],
      null,
      bounds,
    );
    expect(rules[0].ranges[0]).toEqual({ startRow: 0, endRow: 1300, startColumn: 0, endColumn: 0 });
  });

  it("leaves hand-authored ranges below the span threshold untouched", () => {
    const { rules } = excelCfToUniverRules(
      [{ ref: "A1:A2000", rules: [{ type: "duplicateValues", priority: 1, style: {} }] }],
      null,
      bounds,
    );
    expect(rules[0].ranges[0].endRow).toBe(1999);
  });
});

describe("parseIterativeCalc", () => {
  it("reads the opt-in and count from calcPr", () => {
    const xml = '<workbook><calcPr calcId="171027" iterate="1" iterateCount="50" iterateDelta="0.001"/></workbook>';
    expect(parseIterativeCalc(xml)).toEqual({ iterateCount: 50 });
  });

  it("defaults to 100 rounds when iterateCount is absent", () => {
    expect(parseIterativeCalc('<calcPr iterate="1"/>')).toEqual({ iterateCount: 100 });
  });

  it("returns null when iterative calc is off or calcPr is missing", () => {
    expect(parseIterativeCalc('<calcPr calcId="171027" fullCalcOnLoad="1"/>')).toBeNull();
    expect(parseIterativeCalc("<workbook/>")).toBeNull();
  });

  it("caps iteration counts at Excel's default of 100", () => {
    expect(parseIterativeCalc('<calcPr iterate="1" iterateCount="150"/>')).toEqual({ iterateCount: 100 });
    expect(parseIterativeCalc('<calcPr iterate="1" iterateCount="32767"/>')).toEqual({ iterateCount: 100 });
  });
});

describe("conditional formatting import", () => {
  it("maps cellIs with numeric formulae to a number highlight, ordered by priority", () => {
    const { rules } = excelCfToUniverRules(
      [
        {
          ref: "A1:A10",
          rules: [
            { type: "cellIs", operator: "lessThan", formulae: ["0"], priority: 2, style: {} },
            {
              type: "cellIs",
              operator: "between",
              formulae: [1, 5],
              priority: 1,
              style: { font: { bold: true }, fill: { fgColor: { argb: "FFFFC7CE" } } },
            },
          ],
        },
      ],
      null,
    );
    expect(rules).toHaveLength(2);
    // priority 1 first
    expect(rules[0].rule).toMatchObject({
      type: "highlightCell",
      subType: "number",
      operator: "between",
      value: [1, 5],
    });
    expect(rules[0].rule.style).toMatchObject({ bl: 1, bg: { rgb: "#FFC7CE" } });
    expect(rules[0].ranges).toEqual([{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 0 }]);
    expect(rules[1].rule).toMatchObject({ subType: "number", operator: "lessThan", value: 0 });
    expect(rules.map((r: any) => r.cfId)).toEqual(["cf-1", "cf-2"]);
  });

  it("maps expression rules to formula highlights", () => {
    const { rules } = excelCfToUniverRules(
      [{ ref: "B2:B9", rules: [{ type: "expression", formulae: ["$B2>$C2"], priority: 1, style: {} }] }],
      null,
    );
    expect(rules[0].rule).toMatchObject({ subType: "formula", value: "=$B2>$C2" });
  });

  it("recovers the literal text from Excel's generated SEARCH formula", () => {
    const { rules } = excelCfToUniverRules(
      [
        {
          ref: "A1:A5",
          rules: [
            {
              type: "containsText",
              operator: "containsText",
              formulae: ['NOT(ISERROR(SEARCH("risk",A1)))'],
              priority: 1,
              style: {},
            },
          ],
        },
      ],
      null,
    );
    expect(rules[0].rule).toMatchObject({ subType: "text", operator: "containsText", value: "risk" });
  });

  it("maps blanks/errors operators without needing a text value", () => {
    const { rules } = excelCfToUniverRules(
      [
        {
          ref: "A1",
          rules: [
            { type: "containsText", operator: "containsBlanks", formulae: [], priority: 1, style: {} },
          ],
        },
      ],
      null,
    );
    expect(rules[0].rule).toMatchObject({ subType: "text", operator: "containsBlanks" });
  });

  it("maps colorScale, dataBar, top10, aboveAverage and duplicates", () => {
    const { rules, dropped } = excelCfToUniverRules(
      [
        {
          ref: "C1:C20",
          rules: [
            {
              type: "colorScale",
              priority: 1,
              cfvo: [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }],
              color: [{ argb: "FFF8696B" }, { argb: "FFFFEB84" }, { argb: "FF63BE7B" }],
            },
            { type: "dataBar", priority: 2, cfvo: [{ type: "min" }, { type: "max" }], color: { argb: "FF638EC6" } },
            { type: "top10", priority: 3, rank: 5, percent: true, bottom: false, style: {} },
            { type: "aboveAverage", priority: 4, aboveAverage: false, style: {} },
            { type: "duplicateValues", priority: 5, style: {} },
            { type: "iconSet", priority: 6, iconSet: "3TrafficLights1", cfvo: [] },
          ],
        },
      ],
      null,
    );
    expect(rules).toHaveLength(5);
    expect(rules[0].rule.type).toBe("colorScale");
    expect(rules[0].rule.config).toHaveLength(3);
    expect(rules[0].rule.config[1]).toMatchObject({ index: 1, color: "#FFEB84", value: { type: "percentile", value: 50 } });
    expect(rules[1].rule).toMatchObject({ type: "dataBar", config: { positiveColor: "#638EC6", isGradient: true } });
    expect(rules[2].rule).toMatchObject({ subType: "rank", value: 5, isPercent: true, isBottom: false });
    expect(rules[3].rule).toMatchObject({ subType: "average", operator: "lessThan" });
    expect(rules[4].rule).toMatchObject({ subType: "duplicateValues" });
    expect(dropped).toEqual(["iconSet"]);
  });
});

describe("conditional formatting export (Univer → ExcelJS)", () => {
  it("round-trips number, formula and containsText families", () => {
    const { rules } = excelCfToUniverRules(
      [
        {
          ref: "A1:A10",
          rules: [
            { type: "cellIs", operator: "greaterThan", formulae: [10], priority: 1, style: { font: { bold: true } } },
            { type: "expression", formulae: ["$A1<0"], priority: 2, style: {} },
          ],
        },
      ],
      null,
    );
    const { entries, dropped } = univerCfToExcelCf(rules);
    expect(dropped).toEqual([]);
    expect(entries).toHaveLength(2);
    expect(entries[0].ref).toBe("A1:A10");
    expect(entries[0].rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan", formulae: [10] });
    expect(entries[0].rules[0].style.font.bold).toBe(true);
    expect(entries[1].rules[0]).toMatchObject({ type: "expression", formulae: ["$A1<0"] });
  });

  it("emits expression equivalents for text operators ExcelJS cannot write", () => {
    const { entries } = univerCfToExcelCf([
      {
        cfId: "cf-1",
        ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: { type: "highlightCell", subType: "text", operator: "beginsWith", value: "ok", style: {} },
      },
    ]);
    expect(entries[0].rules[0].type).toBe("expression");
    expect(entries[0].rules[0].formulae[0]).toBe('LEFT(A1,LEN("ok"))="ok"');
  });

  it("drops duplicate/unique-value rules (unwritable by ExcelJS) and reports them", () => {
    const { entries, dropped } = univerCfToExcelCf([
      {
        cfId: "cf-1",
        ranges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: { type: "highlightCell", subType: "duplicateValues", style: {} },
      },
    ]);
    expect(entries).toHaveLength(0);
    expect(dropped).toEqual(["duplicateValues"]);
  });
});

describe("data validation import", () => {
  it("regroups per-cell shared rules into rectangles", () => {
    const rule = { type: "list", formulae: ['"Yes,No"'], allowBlank: true };
    const model: Record<string, any> = {};
    for (let r = 1; r <= 3; r++) for (const c of ["A", "B"]) model[`${c}${r}`] = rule;
    const out = excelDvToUniverRules(model);
    expect(out).toHaveLength(1);
    expect(out[0].ranges).toEqual([{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }]);
    expect(out[0]).toMatchObject({ type: "list", formula1: "Yes,No", showDropDown: true });
  });

  it("prefixes '=' on list range refs and keeps numeric bounds plain", () => {
    const listRule = { type: "list", formulae: ["Sheet2!$A$1:$A$5"] };
    const wholeRule = { type: "whole", operator: "between", formulae: [1, 10], showErrorMessage: true, error: "1-10 only", errorStyle: "warning" };
    const out = excelDvToUniverRules({ A1: listRule, B1: wholeRule });
    const list = out.find((r) => r.type === "list")!;
    const whole = out.find((r) => r.type === "whole")!;
    expect(list.formula1).toBe("=Sheet2!$A$1:$A$5");
    expect(whole).toMatchObject({
      formula1: "1",
      formula2: "10",
      operator: "between",
      showErrorMessage: true,
      error: "1-10 only",
      errorStyle: 2,
    });
  });

  it("inverts OOXML's suppress-dropdown flag", () => {
    const out = excelDvToUniverRules({ A1: { type: "list", formulae: ['"a,b"'], showDropDown: true } });
    expect(out[0].showDropDown).toBe(false);
  });

  it("drops rules that fragment into pathologically many ranges", () => {
    // Vendor artifact shape: one shared rule on cells scattered every 256
    // columns — coalescing can't merge them, so each cell stays its own
    // range. Hydrating thousands of these hangs Univer's RuleMatrix.
    const scattered = { type: "list", formulae: ['"a,b"'] };
    const model: Record<string, any> = { A1: { type: "list", formulae: ['"Yes,No"'] } };
    for (let i = 0; i < 300; i++) model[`${colLetters(i * 2)}5`] = scattered;
    const out = excelDvToUniverRules(model);
    expect(out).toHaveLength(1);
    expect(out[0].ranges).toEqual([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }]);
  });
});

/** 0 → A, 25 → Z, 26 → AA … (test helper for scattered-cell fixtures). */
function colLetters(idx: number): string {
  let s = "";
  for (let n = idx; n >= 0; n = Math.floor(n / 26) - 1) {
    s = String.fromCharCode(65 + (n % 26)) + s;
  }
  return s;
}

describe("data validation export (Univer → ExcelJS)", () => {
  it("expands ranges to per-cell entries sharing one rule object", () => {
    const { model } = univerDvToExcelDvModel([
      {
        uid: "dv-1",
        type: "list",
        ranges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }],
        formula1: "Yes,No",
        showDropDown: true,
      },
    ]);
    expect(Object.keys(model).sort()).toEqual(["A1", "A2"]);
    expect(model.A1).toBe(model.A2); // shared object → ExcelJS re-squeezes sqref
    expect(model.A1).toMatchObject({ type: "list", formulae: ['"Yes,No"'] });
    expect(model.A1.showDropDown).toBeUndefined();
  });

  it("drops oversized and inexpressible rules with a report", () => {
    const { model, dropped } = univerDvToExcelDvModel([
      { uid: "dv-1", type: "checkbox", ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }] },
      { uid: "dv-2", type: "decimal", ranges: [{ startRow: 0, endRow: 999_999, startColumn: 0, endColumn: 0 }], formula1: "1" },
    ]);
    expect(Object.keys(model)).toHaveLength(0);
    expect(dropped).toHaveLength(2);
  });
});

describe("coalesceAddressesToRanges", () => {
  it("merges an expanded rectangle back into one range", () => {
    const addrs: string[] = [];
    for (let r = 2; r <= 5; r++) for (const c of ["B", "C", "D"]) addrs.push(`${c}${r}`);
    expect(coalesceAddressesToRanges(addrs)).toEqual([
      { startRow: 1, endRow: 4, startColumn: 1, endColumn: 3 },
    ]);
  });

  it("keeps disjoint blocks apart", () => {
    const out = coalesceAddressesToRanges(["A1", "A2", "C1", "C2"]);
    expect(out).toHaveLength(2);
  });
});

describe("notes", () => {
  it("flattens rich-text notes and keeps plain strings", () => {
    expect(noteTextFromExcelJs("hello")).toBe("hello");
    expect(noteTextFromExcelJs({ texts: [{ text: "a " }, { text: "b" }] })).toBe("a b");
    expect(noteTextFromExcelJs({})).toBeNull();
    expect(noteTextFromExcelJs("")).toBeNull();
  });

  it("round-trips through the resource shape", () => {
    const res = notesToUniverResource([{ sheetId: "s1", row: 2, col: 3, text: "check this" }]);
    expect(res.s1[2][3]).toMatchObject({ note: "check this", row: 2, col: 3 });
    expect(univerNotesToList(res.s1)).toEqual([{ row: 2, col: 3, text: "check this" }]);
  });
});

describe("feature drift detection", () => {
  const names = new Map([["s1", "Sheet1"]]);
  const resOf = (cfRules: any[]) => [{ name: CF_RESOURCE_NAME, data: JSON.stringify({ s1: cfRules }) }];

  it("sees no drift for identical state regardless of key order", () => {
    const a = perSheetFeatureState([
      { name: DV_RESOURCE_NAME, data: JSON.stringify({ s1: [{ uid: "dv-1", type: "list" }] }) },
    ]);
    const b = perSheetFeatureState([
      { name: DV_RESOURCE_NAME, data: JSON.stringify({ s1: [{ type: "list", uid: "dv-1" }] }) },
    ]);
    expect(featureDriftDetail(a, b, names)).toBeNull();
  });

  it("treats empty per-sheet entries as absent", () => {
    const a = perSheetFeatureState([]);
    const b = perSheetFeatureState([{ name: CF_RESOURCE_NAME, data: JSON.stringify({ s1: [] }) }]);
    expect(featureDriftDetail(a, b, names)).toBeNull();
  });

  it("reports the changed family and sheet name", () => {
    const a = perSheetFeatureState(resOf([]));
    const b = perSheetFeatureState(resOf([{ cfId: "cf-1", ranges: [], stopIfTrue: false, rule: {} }]));
    expect(featureDriftDetail(a, b, names)).toBe("conditional formatting (Sheet1)");
  });

  it("ignores sheets that no longer exist", () => {
    const a = perSheetFeatureState([
      { name: NOTE_RESOURCE_NAME, data: JSON.stringify({ gone: { 0: { 0: { note: "x" } } } }) },
    ]);
    const b = perSheetFeatureState([]);
    expect(featureDriftDetail(a, b, names)).toBeNull();
  });
});

describe("applyFeatureMirror", () => {
  /** Minimal duck-typed ExcelJS worksheet good enough for the mirror. */
  function mockWorksheet() {
    const cells = new Map<string, any>();
    const ws: any = {
      conditionalFormattings: [] as any[],
      dataValidations: { model: {} },
      addConditionalFormatting: (cf: any) => ws.conditionalFormattings.push(cf),
      removeConditionalFormatting: (filter: any) => {
        ws.conditionalFormattings = ws.conditionalFormattings.filter(filter);
      },
      getCell: (r: number, c: number) => {
        const key = `${r},${c}`;
        if (!cells.has(key)) {
          const cell: any = { _comment: undefined };
          Object.defineProperty(cell, "note", {
            get: () => cell._comment?.note,
            set: (v) => { cell._comment = { note: v }; },
          });
          cells.set(key, cell);
        }
        return cells.get(key);
      },
      eachRow: (_opts: any, fn: (row: any) => void) => {
        const rows = new Map<number, any[]>();
        for (const [key, cell] of cells) {
          const r = Number(key.split(",")[0]);
          (rows.get(r) ?? rows.set(r, []).get(r)!).push(cell);
        }
        for (const rowCells of rows.values()) {
          fn({ eachCell: (_o: any, cb: (cell: any) => void) => rowCells.forEach(cb) });
        }
      },
      _cells: cells,
    };
    return ws;
  }

  const snapshotWith = (resources: any[]) => ({
    sheets: { s1: { name: "Sheet1" } },
    resources,
  });

  it("writes drifted CF and DV state and leaves clean sheets alone", () => {
    const ws = mockWorksheet();
    ws.conditionalFormattings.push({ ref: "Z1", rules: [{ type: "iconSet" }] }); // pre-existing parsed state
    const wb = { getWorksheet: (name: string) => (name === "Sheet1" ? ws : undefined) };

    const cfRules = [
      {
        cfId: "cf-1",
        ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
        stopIfTrue: false,
        rule: { type: "highlightCell", subType: "number", operator: "greaterThan", value: 1, style: {} },
      },
    ];
    const resources = [{ name: CF_RESOURCE_NAME, data: JSON.stringify({ s1: cfRules }) }];

    // Baseline equals current → untouched (parsed iconSet survives).
    const baseline = perSheetFeatureState(resources);
    applyFeatureMirror(wb, snapshotWith(resources), baseline);
    expect(ws.conditionalFormattings[0].rules[0].type).toBe("iconSet");

    // Baseline empty → drift → replaced with the live rules.
    applyFeatureMirror(wb, snapshotWith(resources), new Map());
    expect(ws.conditionalFormattings).toHaveLength(1);
    expect(ws.conditionalFormattings[0].rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan" });
  });

  it("clears stale notes and writes live ones on drift", () => {
    const ws = mockWorksheet();
    ws.getCell(1, 1).note = "stale";
    const wb = { getWorksheet: () => ws };
    const resources = [
      {
        name: NOTE_RESOURCE_NAME,
        data: JSON.stringify({ s1: { 4: { 2: { id: "n", row: 4, col: 2, width: 1, height: 1, note: "fresh" } } } }),
      },
    ];
    applyFeatureMirror(wb, snapshotWith(resources), new Map());
    expect(ws.getCell(1, 1).note).toBeUndefined();
    expect(ws.getCell(5, 3).note).toBe("fresh");
  });
});
