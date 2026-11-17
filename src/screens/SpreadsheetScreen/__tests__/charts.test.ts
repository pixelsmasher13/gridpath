import { describe, it, expect, vi, beforeEach } from "vitest";
import type { NativeChart } from "../../../../packages/grid/src/charts/types";

// Charts on the desktop: what the renderer is told to draw for each kind
// of chart the reader understands, and the bridge that keeps a tab's
// charts current. The Tauri side is a fake.
const calls: Array<{ cmd: string; args: any }> = [];
let answers: Record<string, (args: any) => any> = {};
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    calls.push({ cmd, args });
    const answer = answers[cmd];
    return answer ? answer(args) : null;
  },
}));

const { options } = await import("../../../../packages/grid/src/charts/chartRenderer");
const engine = await import("../ironcalcEngine");

const series = (name: string, values: Array<number | null>, over: Partial<NativeChart["series"][number]> = {}): NativeChart["series"][number] => ({
  name, nameRef: null, categoriesRef: null, valuesRef: null, categories: ["FY24", "FY25", "FY26"], values, color: "#4472C4", width: 2, markers: false, ...over,
});
const chart = (over: Partial<NativeChart>): NativeChart => ({
  id: "c1", part: "xl/charts/chart1.xml", sheet: "Model", title: "Revenue", kind: "barChart", unsupported: null,
  anchor: { from: { row: 1, col: 1, dx: 0, dy: 0 }, to: null, width: 480, height: 300, absolute: false },
  series: [], legend: true, legendPosition: "b", yMin: null, yMax: null, numberFormat: "", blanks: "gap", ...over,
});

describe("what the renderer draws", () => {
  it("columns side by side", () => {
    const o: any = options(chart({ barDir: "col", barGrouping: "clustered", series: [series("SNPS", [5, 6, 7], { kind: "bar" }), series("SNOW", [2, 3, 4], { kind: "bar", color: "#ED7D31" })] }));
    expect(o.series.map((s: any) => [s.type, s.stack, s.data])).toEqual([["bar", undefined, [5, 6, 7]], ["bar", undefined, [2, 3, 4]]]);
    expect(o.series[1].itemStyle.color).toBe("#ED7D31");
    // Excel's default gap: a bar group is 40% of its category's slot.
    expect(o.series[0]).toMatchObject({ barCategoryGap: "60%", barGap: "0%" });
    const tight: any = options(chart({ barDir: "col", barGrouping: "clustered", barGapWidth: 50, barOverlap: -20, series: [series("A", [1], { kind: "bar" })] }));
    expect(tight.series[0]).toMatchObject({ barCategoryGap: "33%", barGap: "20%" });
    expect(o.xAxis).toMatchObject({ type: "category", boundaryGap: true, data: ["FY24", "FY25", "FY26"] });
    expect(o.yAxis.type).toBe("value");
  });

  it("stacked columns, and stacked to 100%", () => {
    const two = [series("A", [1, 3, 0], { kind: "bar" }), series("B", [3, 1, 0], { kind: "bar" })];
    const stacked: any = options(chart({ barDir: "col", barGrouping: "stacked", series: two }));
    expect(stacked.series.map((s: any) => s.stack)).toEqual(["total", "total"]);
    expect(stacked.series[0].data).toEqual([1, 3, 0]);
    const share: any = options(chart({ barDir: "col", barGrouping: "percentStacked", series: two }));
    expect(share.series.map((s: any) => s.data)).toEqual([[0.25, 0.75, 0], [0.75, 0.25, 0]]);
    expect(share.yAxis.max).toBe(1);
    expect(share.yAxis.axisLabel.formatter(0.5)).toBe("50%");
  });

  it("lying bars put the categories on the upright axis", () => {
    const o: any = options(chart({ barDir: "bar", barGrouping: "clustered", series: [series("A", [1, 2, 3], { kind: "bar" })] }));
    expect(o.yAxis.type).toBe("category");
    expect(o.xAxis.type).toBe("value");
  });

  it("columns with a line on the second axis", () => {
    const o: any = options(chart({
      barDir: "col", barGrouping: "clustered", numberFormat: "#,##0", numberFormat2: "0%", y2Max: 1,
      series: [series("Revenue", [5, 6, 7], { kind: "bar" }), series("Margin", [0.6, 0.65, 0.7], { kind: "line", secondary: true, smooth: true, markers: true, color: "#C00000" })],
    }));
    expect(o.yAxis).toHaveLength(2);
    expect(o.yAxis[1]).toMatchObject({ position: "right", max: 1 });
    expect(o.yAxis[1].axisLabel.formatter(0.65)).toBe("65%");
    expect(o.series[0]).toMatchObject({ type: "bar" });
    expect(o.series[0].yAxisIndex).toBeUndefined();
    expect(o.series[1]).toMatchObject({ type: "line", yAxisIndex: 1, smooth: true, showSymbol: true });
    // Bars sit between the ticks even with a line over them.
    expect(o.xAxis.boundaryGap).toBe(true);
  });

  it("a line chart from an engine that says nothing about kinds is still a line chart", () => {
    const o: any = options(chart({ kind: "lineChart", series: [series("A", [1, null, 3])] }));
    expect(o.series[0]).toMatchObject({ type: "line", data: [1, null, 3], smooth: false });
    // No number format: the axis shows the numbers as they are.
    expect(o.yAxis.axisLabel.formatter).toBeUndefined();
    expect(o.xAxis.boundaryGap).toBe(false);
    const zero: any = options(chart({ kind: "lineChart", blanks: "zero", series: [series("A", [1, null, 3])] }));
    expect(zero.series[0].data).toEqual([1, 0, 3]);
  });
});

describe("the bridge keeps a tab's charts current", () => {
  const layout = () => new TextEncoder().encode(JSON.stringify({ sheets: [] })).buffer;
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  beforeEach(() => {
    calls.length = 0;
    answers = { calc_engine_workbook_layout: () => layout() };
    (globalThis as any).localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
  });

  it("reads them on open and again when values change, and tells the grid", async () => {
    let revenue = 5;
    answers.calc_engine_load = () => ({ sheets: ["Model"], load_ms: 0, eval_ms: 0, others: [], needs_saving: false });
    answers.calc_engine_charts = () => [chart({ series: [series("Revenue", [revenue], { kind: "bar" })] })];
    answers.calc_engine_edit_batch = () => {
      revenue = 9;
      return { recalc_ms: 1, changed: [{ sheet: 0, row: 2, column: 2, value: 9 }], others: [], sheets: null, links: null };
    };
    let told = 0;
    const off = engine.subscribeIroncalcCharts("t-charts", () => told++);
    expect(engine.ironcalcEngineCharts("t-charts")).toEqual([]);
    await engine.ironcalcEngineOpenLayout("t-charts", "/m/Model.xlsx");
    engine.registerIroncalcApplier("t-charts", () => true);
    await wait(120);
    const first = engine.ironcalcEngineCharts("t-charts");
    expect(first[0].series[0].values).toEqual([5]);
    expect(told).toBe(1);
    // The same array until something changes (the grid compares by identity).
    expect(engine.ironcalcEngineCharts("t-charts")).toBe(first);

    engine.ironcalcEngineWriteCells("t-charts", "Model", "typed", [{ row: 2, column: 2, value: "9" }]);
    await engine.ironcalcEngineSettled("t-charts");
    await wait(120);
    expect(engine.ironcalcEngineCharts("t-charts")[0].series[0].values).toEqual([9]);
    expect(told).toBe(2);
    off();
    engine.ironcalcEngineOnClose("t-charts");
    expect(engine.ironcalcEngineCharts("t-charts")).toEqual([]);
  });

  it("does not keep asking a workbook that has no charts", async () => {
    answers.calc_engine_load = () => ({ sheets: ["Model"], load_ms: 0, eval_ms: 0, others: [], needs_saving: false });
    answers.calc_engine_charts = () => [];
    answers.calc_engine_edit_batch = () => ({ recalc_ms: 1, changed: [{ sheet: 0, row: 1, column: 1, value: 1 }], others: [], sheets: null, links: null });
    await engine.ironcalcEngineOpenLayout("t-none", "/m/Plain.xlsx");
    engine.registerIroncalcApplier("t-none", () => true);
    await wait(120);
    engine.ironcalcEngineWriteCells("t-none", "Model", "typed", [{ row: 1, column: 1, value: "1" }]);
    await engine.ironcalcEngineSettled("t-none");
    await wait(120);
    expect(calls.filter((c) => c.cmd === "calc_engine_charts")).toHaveLength(1);
    engine.ironcalcEngineOnClose("t-none");
  });
});

// The agent's create_chart: from the tool call to the chart on the sheet,
// what the agent is told about it, and taking it away again.
describe("a chart the agent makes", () => {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  beforeEach(() => {
    calls.length = 0;
    answers = { calc_engine_workbook_layout: () => new TextEncoder().encode(JSON.stringify({ sheets: [] })).buffer };
    (globalThis as any).localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
  });

  it("reads the tool call into a chart to make", async () => {
    const { interpretToolCall } = await import("../agent/toolToMutation");
    const call = {
      sheet: "Model", anchor: "$H$2", type: "column", title: " Revenue and margin ", categories: "B1:G1",
      series: [
        { name_cell: "A2", values: "B2:G2" },
        { name: "Margin", values: "'P&L'!B9:G9", type: "line", secondary_axis: true, color: "#C00000" },
      ],
    };
    expect(interpretToolCall("create_chart", call)).toEqual({
      kind: "chart",
      spec: {
        sheet: "Model", row: 1, col: 7, width: null, height: null, title: "Revenue and margin", kind: "column", grouping: "clustered", categories: "B1:G1",
        series: [
          { name: null, nameRef: "A2", values: "B2:G2", kind: null, secondary: false, color: null },
          { name: "Margin", nameRef: null, values: "'P&L'!B9:G9", kind: "line", secondary: true, color: "#C00000" },
        ],
      },
    });
    const kinds = (type: string) => {
      const r = interpretToolCall("create_chart", { ...call, type });
      return r.kind === "chart" ? [r.spec.kind, r.spec.grouping] : r;
    };
    expect(kinds("stacked_column")).toEqual(["column", "stacked"]);
    expect(kinds("percent_stacked_column")).toEqual(["column", "percentStacked"]);
    expect(kinds("stacked_bar")).toEqual(["bar", "stacked"]);
    expect(kinds("line")).toEqual(["line", "clustered"]);
    // The series list as a JSON string, as models sometimes send it.
    expect(interpretToolCall("create_chart", { ...call, series: JSON.stringify(call.series) }).kind).toBe("chart");
    // What cannot be read says what is wrong.
    const why = (over: object) => {
      const r = interpretToolCall("create_chart", { ...call, ...over });
      return r.kind === "ignored" ? r.reason : "";
    };
    expect(why({ type: "pie" })).toContain("must be one of column, stacked_column");
    expect(why({ anchor: "H" })).toContain("top-left cell");
    expect(why({ sheet: " " })).toContain("missing sheet");
    expect(why({ series: [] })).toContain("1 to 24 series");
    expect(why({ series: [{ name: "Revenue" }] })).toContain("series 1 has no 'values' range");
    expect(why({ series: [{ values: "B2:G2", type: "area" }] })).toContain('"line" or "column"');
  });

  it("tells the agent what is plotted, and what looks wrong", async () => {
    const { describeNewChart } = await import("../agent/chartReport");
    const made = chart({
      id: "new:1", part: "", title: "Revenue and margin", barDir: "col", barGrouping: "clustered",
      anchor: { from: { row: 1, col: 7, dx: 0, dy: 0 }, to: null, width: 480, height: 288, absolute: false },
      series: [
        series("Revenue", [100, 200, 300], { kind: "bar", valuesRef: "Model!$B$2:$D$2", categoriesRef: "Model!$B$1:$D$1" }),
        series("Margin", [0.25, null, 0.3], { kind: "line", secondary: true, valuesRef: "Model!$B$9:$D$9", categoriesRef: "Model!$B$1:$D$1" }),
        series("Typo", [null, null, null], { kind: "line", valuesRef: "Model!$B$40:$D$40", categoriesRef: "Model!$B$1:$D$1" }),
      ],
    });
    const { summary, result } = describeNewChart(made) as any;
    expect(summary).toBe("columns + line, 3 series, 3 points");
    expect(result.chart).toMatchObject({ sheet: "Model", top_left: "H2", size_px: [480, 288], type: "columns + line", categories: { range: "Model!$B$1:$D$1", count: 3, first: "FY24", last: "FY26" } });
    expect(result.chart.series[0]).toEqual({ name: "Revenue", values: "Model!$B$2:$D$2", drawn_as: "column", axis: "main", points: 3, first: 100, last: 300 });
    expect(result.chart.series[1]).toMatchObject({ drawn_as: "line", axis: "second (right)", first: 0.25, last: 0.3 });
    expect(result.warnings).toEqual([
      '"Margin": 1 of 3 cells are empty or not numbers and leave gaps.',
      '"Typo" (Model!$B$40:$D$40) has no numbers, so nothing is drawn for it: check the range.',
    ]);
    // Nothing to warn about: no warnings field at all.
    expect(describeNewChart(chart({ barDir: "bar", barGrouping: "stacked", series: [series("A", [1, 2, 3], { kind: "bar" })] }))).toMatchObject({ summary: "stacked bars, 1 series, 3 points", result: { chart: { categories: null } } });
    expect((describeNewChart(chart({ series: [series("A", [1], { kind: "line" })] })).result as any).warnings).toBeUndefined();
  });

  it("is added through the engine, shown, read back for the save, and taken away again", async () => {
    const spec = { id: "new:7", sheet: "Model", row: 1, col: 7, kind: "column" as const, series: [{ values: "B2:D2" }] };
    let held: NativeChart[] = [chart({ id: "file", series: [series("Old", [1], { kind: "bar" })] })];
    answers.calc_engine_load = () => ({ sheets: ["Model"], load_ms: 0, eval_ms: 0, others: [], needs_saving: false });
    answers.calc_engine_charts = () => held;
    answers.calc_engine_add_chart = (args) => {
      if (args.spec.sheet !== "Model") throw `there is no sheet "${args.spec.sheet}" to put the chart on. Sheets: Model`;
      const made = chart({ id: args.spec.id, part: "", series: [series("Revenue", [100, 200, 300], { kind: "bar" })] });
      held = [...held.filter((c) => c.id !== made.id), made];
      return made;
    };
    answers.calc_engine_remove_chart = (args) => {
      const before = held.length;
      held = held.filter((c) => c.id !== args.id);
      return held.length < before;
    };
    await engine.ironcalcEngineOpenLayout("t-make", "/m/Model.xlsx");
    await wait(120);
    const made = await engine.ironcalcEngineAddChart("t-make", spec);
    expect(made).toMatchObject({ id: "new:7", part: "" });
    expect(calls.find((c) => c.cmd === "calc_engine_add_chart")?.args).toEqual({ tab: "t-make", spec });
    await wait(120);
    expect(engine.ironcalcEngineCharts("t-make").map((c) => c.id)).toEqual(["file", "new:7"]);
    // A save writes the charts of this session, not the ones the file has.
    expect((await engine.ironcalcEngineSessionCharts("t-make")).map((c) => c.id)).toEqual(["new:7"]);
    // What the engine refuses comes back as the reason.
    await expect(engine.ironcalcEngineAddChart("t-make", { ...spec, sheet: "Charts" })).rejects.toThrow('there is no sheet "Charts"');

    // Reject / undo of the batch: the mutation takes the chart away; redo adds it again.
    const { applyMutationForward, reverseMutation } = await import("../agent/batchMutations");
    const grid: any = {
      addChart: (s: any) => engine.ironcalcEngineAddChart("t-make", s).then((c) => ({ ok: true, chart: c })),
      removeChart: (id: string) => engine.ironcalcEngineRemoveChart("t-make", id).then(() => {}),
    };
    reverseMutation(grid, { type: "create_chart", chart: spec });
    await engine.ironcalcEngineSettled("t-make");
    await wait(120);
    expect(engine.ironcalcEngineCharts("t-make").map((c) => c.id)).toEqual(["file"]);
    expect(await engine.ironcalcEngineSessionCharts("t-make")).toEqual([]);
    applyMutationForward(grid, { type: "create_chart", chart: spec });
    await engine.ironcalcEngineSettled("t-make");
    await wait(120);
    expect(engine.ironcalcEngineCharts("t-make").map((c) => c.id)).toEqual(["file", "new:7"]);
    // A grid that has no charts (an untitled workbook) is left alone.
    expect(() => applyMutationForward({} as any, { type: "create_chart", chart: spec })).not.toThrow();
    engine.ironcalcEngineOnClose("t-make");
  });
});
