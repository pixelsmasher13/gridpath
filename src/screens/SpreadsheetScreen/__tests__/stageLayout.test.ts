import { describe, it, expect } from "vitest";
import { layoutStageBlock, colLetter } from "../agent/stageLayout";
import { interpretToolCall } from "../agent/toolToMutation";

const payload = {
  sheet: "Data",
  title: "AAPL 10-Q Q3 FY2026 — income statement (reported)",
  source: "https://www.sec.gov/...aapl-20260627.htm",
  units: "$M",
  columns: ["FY2024A", "FY2025A"],
  rows: [
    { label: "Revenue — Products", values: [294866, 307001] },
    { label: "Revenue — Services", values: [96169, 109160] },
    { label: "Sparse row", values: [null, 42] },
  ],
};

describe("colLetter", () => {
  it("maps 0-indexed columns to A1 letters", () => {
    expect(colLetter(0)).toBe("A");
    expect(colLetter(1)).toBe("B");
    expect(colLetter(25)).toBe("Z");
    expect(colLetter(26)).toBe("AA");
  });
});

describe("layoutStageBlock", () => {
  it("lays out title / source / header / data from row 0 on an empty sheet", () => {
    const { mutations, report } = layoutStageBlock(payload, 0);
    const cells = mutations.filter((m: any) => m.type === "set_cell") as any[];
    const at = (row: number, col: number) => cells.find((c) => c.address.row === row && c.address.col === col);
    expect(at(0, 0)?.new_value).toContain("AAPL 10-Q");
    expect(at(1, 0)?.new_value).toContain("Source:");
    expect(at(1, 0)?.new_value).toContain("$M");
    expect(at(2, 1)?.new_value).toBe("FY2024A");
    expect(at(2, 2)?.new_value).toBe("FY2025A");
    expect(at(3, 0)?.new_value).toBe("Revenue — Products");
    expect(at(3, 1)?.new_value).toBe(294866);
    expect(at(4, 2)?.new_value).toBe(109160);
    // Report is 1-indexed A1.
    expect(report.sheet).toBe("Data");
    expect(report.header_row).toBe(3);
    expect(report.cols).toEqual({ FY2024A: "B", FY2025A: "C" });
    expect(report.rows["Revenue — Products"]).toBe(4);
    expect(report.rows["Sparse row"]).toBe(6);
    expect(report.block).toBe("A1:C6");
  });

  it("stacks below existing content when startRow is offset", () => {
    const { mutations, report } = layoutStageBlock(payload, 10);
    const cells = mutations.filter((m: any) => m.type === "set_cell") as any[];
    expect(Math.min(...cells.map((c) => c.address.row))).toBe(10);
    expect(report.rows["Revenue — Products"]).toBe(14);
    expect(report.block).toBe("A11:C16");
  });

  it("blank values are gaps, not writes", () => {
    const { mutations } = layoutStageBlock(payload, 0);
    const cells = mutations.filter((m: any) => m.type === "set_cell") as any[];
    // Sparse row (row 5, 0-indexed): only the label and the second value land.
    const sparse = cells.filter((c) => c.address.row === 5);
    expect(sparse.map((c) => c.address.col).sort()).toEqual([0, 2]);
  });

  it("emits title/source/header formats", () => {
    const { mutations } = layoutStageBlock(payload, 0);
    const fmts = mutations.filter((m: any) => m.type === "set_format") as any[];
    expect(fmts).toHaveLength(3);
    expect(fmts[0].range).toBe("A1");
    expect(fmts[0].new_format.bold).toBe(true);
    expect(fmts[2].range).toBe("A3:C3");
  });
});

describe("interpretToolCall stage_data", () => {
  it("valid input yields a stage result with defaulted sheet", () => {
    const r = interpretToolCall("stage_data", {
      title: "t",
      source: "s",
      columns: ["FY24"],
      rows: [{ label: "Revenue", values: [1] }],
    });
    expect(r.kind).toBe("stage");
    if (r.kind !== "stage") return;
    expect(r.sheet).toBe("Data");
    expect(r.rows[0]).toEqual({ label: "Revenue", values: [1] });
  });

  it("accepts shorthand row arrays", () => {
    const r = interpretToolCall("stage_data", {
      title: "t",
      source: "s",
      columns: ["A", "B"],
      rows: [["Revenue", 1, 2]],
    });
    expect(r.kind).toBe("stage");
    if (r.kind !== "stage") return;
    expect(r.rows[0]).toEqual({ label: "Revenue", values: [1, 2] });
  });

  it("missing source is rejected with a provenance message", () => {
    const r = interpretToolCall("stage_data", { title: "t", columns: ["A"], rows: [["x", 1]] });
    expect(r.kind).toBe("ignored");
    if (r.kind !== "ignored") return;
    expect(r.reason).toContain("source");
  });

  it("row wider than columns is rejected", () => {
    const r = interpretToolCall("stage_data", {
      title: "t",
      source: "s",
      columns: ["A"],
      rows: [["x", 1, 2]],
    });
    expect(r.kind).toBe("ignored");
  });
});
