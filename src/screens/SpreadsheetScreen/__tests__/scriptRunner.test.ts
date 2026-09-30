import { describe, expect, it } from "vitest";
import {
  buildScriptModelFromSnapshot,
  buildWorkerSource,
  executeSheetScript,
  scriptCore,
  scriptOpsToMutations,
  type ScriptModel,
} from "../agent/scriptRunner";

const LIMITS = { maxTouched: 20_000, maxLogs: 100 };

/** Tiny model builder: cells as { "Sheet1": { A1: 5, B2: "=A1*2" } } with
 * A1-keyed values; strings starting with "=" become formulas whose cached
 * value is null (mirrors an unevaluated import, good enough for reads). */
function model(spec: Record<string, Record<string, string | number | boolean | null>>): ScriptModel {
  const sheets: ScriptModel["sheets"] = {};
  for (const [name, cells] of Object.entries(spec)) {
    const data: ScriptModel["sheets"][string] = {};
    for (const [a1, raw] of Object.entries(cells)) {
      const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
      let col = 0;
      for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
      const key = `${parseInt(m[2], 10) - 1},${col - 1}`;
      const isFormula = typeof raw === "string" && raw.startsWith("=");
      data[key] = { v: isFormula ? null : raw, f: isFormula ? raw : null };
    }
    sheets[name] = data;
  }
  return { sheetNames: Object.keys(spec), sheets };
}

function run(code: string, m: ScriptModel, limits = LIMITS) {
  return scriptCore({ code, model: m, limits });
}

describe("scriptCore — writes", () => {
  it("records loop writes with computed addresses", () => {
    const r = run(
      `const s = sheet("Model");
       for (let i = 0; i < 3; i++) s.set("B" + (i + 1), i * 10);`,
      model({ Model: {} }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ops).toEqual([
      { kind: "set", sheet: "Model", row: 0, col: 1, value: 0 },
      { kind: "set", sheet: "Model", row: 1, col: 1, value: 10 },
      { kind: "set", sheet: "Model", row: 2, col: 1, value: 20 },
    ]);
  });

  it("treats '=' strings as formulas and reads them back unevaluated", () => {
    const r = run(
      `const s = sheet("Model");
       s.set("C1", "=A1*2");
       const back = s.get("C1");
       log(JSON.stringify(back));`,
      model({ Model: { A1: 21 } }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ops[0]).toEqual({ kind: "set", sheet: "Model", row: 0, col: 2, value: "=A1*2" });
    expect(r.logs[0]).toBe('{"value":null,"formula":"=A1*2"}');
  });

  it("overlays own literal writes on reads; snapshot supplies the rest", () => {
    const r = run(
      `const s = sheet("Model");
       s.set("A2", 99);
       log(s.get("A1").value, s.get("A2").value, s.get("Z9").value);`,
      model({ Model: { A1: 5 } }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.logs[0]).toBe("5 99 null");
  });

  it("values() returns a rows×cols 2D array", () => {
    const r = run(
      `log(JSON.stringify(sheet("Model").values("A1:B2")));`,
      model({ Model: { A1: 1, B1: 2, A2: 3 } }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(JSON.parse(r.logs[0])).toEqual([
      [1, 2],
      [3, null],
    ]);
  });

  it("setValues preserves null/'' cells like set_range", () => {
    const r = run(
      `sheet("Model").setValues("B2", [["x", null], ["", 7]]);`,
      model({ Model: {} }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ops).toEqual([
      { kind: "set", sheet: "Model", row: 1, col: 1, value: "x" },
      { kind: "set", sheet: "Model", row: 2, col: 2, value: 7 },
    ]);
  });

  it("rejects NaN writes with a pointed error", () => {
    const r = run(`sheet("Model").set("A1", 1 / 0 - 1 / 0);`, model({ Model: {} }));
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error).toContain("NaN");
  });

  it("enforces the touched-cell cap", () => {
    const r = run(
      `const s = sheet("Model");
       for (let i = 1; i <= 10; i++) s.set("A" + i, i);`,
      model({ Model: {} }),
      { maxTouched: 5, maxLogs: 100 },
    );
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error).toContain("write cap exceeded");
  });
});

describe("scriptCore — sheets, format, clear", () => {
  it("a new sheet name auto-registers and its writes land against it", () => {
    const r = run(
      `sheet("Notes").set("A1", "hello"); sheet("Model").set("B2", 5);`,
      model({ Model: {}, Assumptions: {} }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    const noteOp = r.ops.find((o: any) => o.sheet === "Notes");
    expect(noteOp).toBeTruthy();
  });

  it("auto-registered sheets read back as empty, and appear in sheets()", () => {
    const r = run(
      `const s = sheet("Fresh"); log(s.usedRange()); log(sheets().join(","));`,
      model({ Model: {} }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.logs[0]).toBe("null");
    expect(r.logs[1]).toContain("Fresh");
  });

  it("sheet() without a name works only for single-sheet workbooks", () => {
    const single = run(`sheet().set("A1", 1);`, model({ Only: {} }));
    expect(single.ok).toBe(true);
    const multi = run(`sheet().set("A1", 1);`, model({ A: {}, B: {} }));
    expect(multi.ok).toBe(false);
  });

  it("format records the op and rejects unknown keys", () => {
    const good = run(`sheet("M").format("A1:B1", { bold: true, number_format: "0.0%" });`, model({ M: {} }));
    expect(good.ok).toBe(true);
    if (good.ok) {
      expect(good.ops[0]).toEqual({
        kind: "format",
        sheet: "M",
        range: "A1:B1",
        format: { bold: true, number_format: "0.0%" },
      });
    }
    const bad = run(`sheet("M").format("A1", { blod: true });`, model({ M: {} }));
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.error).toContain('"blod"');
  });

  it("clear records the op and blanks subsequent reads", () => {
    const r = run(
      `const s = sheet("M");
       s.clear("A1:A2");
       log(s.get("A1").value);`,
      model({ M: { A1: 5, A2: 6 } }),
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.ops[0]).toEqual({ kind: "clear", sheet: "M", range: "A1:A2" });
    expect(r.logs[0]).toBe("null");
  });

  it("usedRange reports the pre-script extent", () => {
    const r = run(`log(sheet("M").usedRange());`, model({ M: { C7: 1, A1: 2 } }));
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.logs[0]).toBe("A1:C7");
  });

  it("caps log entries", () => {
    const r = run(
      `for (let i = 0; i < 500; i++) log(i);`,
      model({ M: {} }),
      { maxTouched: 100, maxLogs: 3 },
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.logs).toEqual(["0", "1", "2"]);
  });

  it("surfaces script runtime errors without throwing", () => {
    const r = run(`nonexistentFn();`, model({ M: {} }));
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error).toContain("nonexistentFn");
  });
});

describe("scriptOpsToMutations", () => {
  it("maps set ops to set_cell mutations, formulas flagged", () => {
    const muts = scriptOpsToMutations([
      { kind: "set", sheet: "M", row: 0, col: 0, value: 7 },
      { kind: "set", sheet: "M", row: 1, col: 0, value: "=A1*2" },
    ]);
    expect(muts).toEqual([
      {
        type: "set_cell",
        address: { sheet: "M", row: 0, col: 0 },
        old_value: null,
        new_value: 7,
        new_formula: null,
      },
      {
        type: "set_cell",
        address: { sheet: "M", row: 1, col: 0 },
        old_value: null,
        new_value: "=A1*2",
        new_formula: "=A1*2",
      },
    ]);
  });

  it("coalesces repeated writes to one cell — last write wins", () => {
    const muts = scriptOpsToMutations([
      { kind: "set", sheet: "M", row: 0, col: 0, value: 1 },
      { kind: "set", sheet: "M", row: 0, col: 0, value: 2 },
      { kind: "set", sheet: "M", row: 0, col: 0, value: 3 },
    ]);
    expect(muts).toHaveLength(1);
    expect((muts[0] as any).new_value).toBe(3);
  });

  it("does NOT coalesce across an intervening clear (order must hold)", () => {
    const muts = scriptOpsToMutations([
      { kind: "set", sheet: "M", row: 0, col: 0, value: 1 },
      { kind: "clear", sheet: "M", range: "A1" },
      { kind: "set", sheet: "M", row: 0, col: 0, value: 2 },
    ]);
    expect(muts.map((m) => m.type)).toEqual(["set_cell", "clear_range", "set_cell"]);
    expect((muts[2] as any).new_value).toBe(2);
  });

  it("expands format and clear ranges into cell lists", () => {
    const muts = scriptOpsToMutations([
      { kind: "format", sheet: "M", range: "A1:B2", format: { bold: true } },
      { kind: "clear", sheet: "M", range: "C1:C2" },
    ]);
    expect(muts[0].type).toBe("set_format");
    expect((muts[0] as any).cells).toHaveLength(4);
    expect(muts[1].type).toBe("clear_range");
    expect((muts[1] as any).cells).toEqual([
      { row: 0, col: 2, old_value: null, old_formula: null },
      { row: 1, col: 2, old_value: null, old_formula: null },
    ]);
  });
});

describe("buildScriptModelFromSnapshot", () => {
  it("walks sheetOrder and extracts v/f, skipping empty cells", () => {
    const m = buildScriptModelFromSnapshot({
      sheetOrder: ["id2", "id1"],
      sheets: {
        id1: { name: "B", cellData: {} },
        id2: {
          name: "A",
          cellData: {
            "0": { "0": { v: 5 }, "1": { v: null, f: "=A1*2" }, "2": {} },
          },
        },
      },
    });
    expect(m.sheetNames).toEqual(["A", "B"]);
    expect(m.sheets["A"]).toEqual({
      "0,0": { v: 5, f: null },
      "0,1": { v: null, f: "=A1*2" },
    });
    expect(m.sheets["B"]).toEqual({});
  });

  it("throws on a missing snapshot", () => {
    expect(() => buildScriptModelFromSnapshot(null)).toThrow(/no workbook/);
  });
});

describe("executeSheetScript (direct path, no Worker in node)", () => {
  it("runs end-to-end from a grid handle to mutations", async () => {
    const snapshot = {
      sheetOrder: ["s1"],
      sheets: {
        s1: {
          name: "Model",
          cellData: { "0": { "0": { v: 10 }, "1": { v: 20 } } },
        },
      },
    };
    const r = await executeSheetScript(
      { getWorkbookSnapshot: () => snapshot },
      `const s = sheet("Model");
       const row = s.values("A1:B1")[0];
       s.set("C1", "=A1+B1");
       log("sum inputs:", row[0] + row[1]);`,
    );
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.logs).toEqual(["sum inputs: 30"]);
    expect(r.mutations).toEqual([
      {
        type: "set_cell",
        address: { sheet: "Model", row: 0, col: 2 },
        old_value: null,
        new_value: "=A1+B1",
        new_formula: "=A1+B1",
      },
    ]);
  });

  it("worker source is self-contained — the serialized core runs standalone", () => {
    // Execute the exact string the Blob worker would run, with a stub
    // `self`. If a build transform ever makes scriptCore lean on
    // module-scope helpers, toString() serialization breaks and ONLY
    // running the string catches it.
    const messages: any[] = [];
    const fakeSelf: any = { postMessage: (m: any) => messages.push(m) };
    new Function("self", buildWorkerSource())(fakeSelf);
    expect(typeof fakeSelf.onmessage).toBe("function");
    fakeSelf.onmessage({
      data: {
        code: `sheet("M").set("A1", "=B1*2"); log("done");`,
        model: model({ M: { B1: 4 } }),
        limits: LIMITS,
      },
    });
    expect(messages).toHaveLength(1);
    expect(messages[0].ok).toBe(true);
    expect(messages[0].ops).toEqual([{ kind: "set", sheet: "M", row: 0, col: 0, value: "=B1*2" }]);
    expect(messages[0].logs).toEqual(["done"]);
  });

  it("returns ok:false (never throws) on a broken script", async () => {
    const r = await executeSheetScript(
      { getWorkbookSnapshot: () => ({ sheetOrder: ["s1"], sheets: { s1: { name: "M", cellData: {} } } }) },
      `throw new Error("boom");`,
    );
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.error).toContain("boom");
  });
});
