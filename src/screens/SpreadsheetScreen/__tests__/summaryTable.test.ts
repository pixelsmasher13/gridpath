import { describe, it, expect } from "vitest";
import { layoutSummary, sheetQualifier, summaryResult, type SummaryFacts, type SummaryRequest } from "../agent/summaryLayout";
import { gatherSummaryFacts } from "../agent/summaryFacts";
import { interpretToolCall } from "../agent/toolToMutation";

// create_summary_table: a pivot-style summary laid out as formulas. The
// layout is pure; the facts come from the engine's row query (a fake here).

const request = (over: Partial<SummaryRequest> = {}): SummaryRequest => ({
  sourceSheet: "Data", headerRow: null, rows: "Region", columns: null, values: [{ field: "Revenue", fn: "sum" }],
  destSheet: "Summary", dest: { row: 0, col: 0 }, title: null, sort: "label", ...over,
});
const facts = (over: Partial<SummaryFacts> = {}): SummaryFacts => ({
  sheet: "Data", firstRow: 2, lastRow: 500,
  rowField: { col: 2, header: "Region", keys: ["East", "West"], format: null },
  colField: null,
  measures: [{ col: 7, header: "Revenue", fn: "sum", format: "#,##0" }],
  ...over,
});
/** The laid-out cells as { A1: what is typed }. */
const cellsOf = (mutations: any[]) => {
  const out: Record<string, unknown> = {};
  for (const m of mutations) {
    if (m.type !== "set_cell") continue;
    const col = String.fromCharCode(65 + m.address.col);
    out[`${col}${m.address.row + 1}`] = m.new_formula ?? m.new_value;
  }
  return out;
};
const formatsOf = (mutations: any[]) => mutations.filter((m) => m.type === "set_format").map((m) => [m.range, m.new_format]);

describe("the table as formulas", () => {
  it("one row per group, a total from the data, and a check", () => {
    const { mutations, report } = layoutSummary(request({ title: "Revenue by region" }), facts());
    expect(cellsOf(mutations)).toEqual({
      A1: "Revenue by region",
      A2: "Region", B2: "Sum of Revenue",
      A3: "East", B3: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A3)",
      A4: "West", B4: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A4)",
      A5: "Total", B5: "=SUM(Data!$H$2:$H$500)",
      A6: "Check (0 = the table holds every row of the data)", B6: "=SUM(B3:B4)-B5",
    });
    expect(report).toMatchObject({ sheet: "Summary", block: "A1:B6", header_row: 2, first_data_row: 3, last_data_row: 4, total_row: 5, check_cell: "B6", rows: { East: 3, West: 4 }, cols: { "Sum of Revenue": "B" } });
    expect(report.note).toContain("Rows added below row 500");
    // Labels and totals stand out; the amounts read as the data's do.
    expect(formatsOf(mutations)).toEqual([
      ["A1", { bold: true }],
      ["B3:B5", { number_format: "#,##0" }],
      ["A2:B2", { bold: true }],
      ["A5:B5", { bold: true }],
      ["A6:B6", { italic: true, font_color: "#808080" }],
    ]);
    // Every write is a cell the review can show and undo.
    expect(mutations.every((m: any) => m.type === "set_cell" || m.type === "set_format")).toBe(true);
    expect(mutations.find((m: any) => m.address?.row === 2 && m.address?.col === 1)).toMatchObject({ new_value: null, new_formula: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A3)" });
  });

  it("several values side by side, each its own kind of total", () => {
    const { mutations, report } = layoutSummary(
      request({ dest: { row: 3, col: 2 } }),
      facts({
        sheet: "P&L data",
        measures: [
          { col: 7, header: "Revenue", fn: "sum", format: null },
          { col: 7, header: "Revenue", fn: "average", format: "0.0" },
          { col: 0, header: "Deal", fn: "count", format: "0.00" },
          { col: 8, header: "Margin", fn: "min", format: "0%" },
          { col: 8, header: "Margin", fn: "max", format: "0%" },
        ],
      }),
    );
    const c = cellsOf(mutations);
    expect(c.C4).toBe("Region");
    expect([c.D4, c.E4, c.F4, c.G4, c.H4]).toEqual(["Sum of Revenue", "Average of Revenue", "Count of Deal", "Min of Margin", "Max of Margin"]);
    const k = "'P&L data'!$C$2:$C$500,$C5";
    expect(c.D5).toBe(`=SUMIFS('P&L data'!$H$2:$H$500,${k})`);
    expect(c.E5).toBe(`=IFERROR(AVERAGEIFS('P&L data'!$H$2:$H$500,${k}),"")`);
    expect(c.F5).toBe(`=COUNTIFS(${k},'P&L data'!$A$2:$A$500,"<>")`);
    expect(c.G5).toBe(`=IF(COUNTIFS(${k},'P&L data'!$I$2:$I$500,"<>")=0,"",MINIFS('P&L data'!$I$2:$I$500,${k}))`);
    expect(c.H5).toBe(`=IF(COUNTIFS(${k},'P&L data'!$I$2:$I$500,"<>")=0,"",MAXIFS('P&L data'!$I$2:$I$500,${k}))`);
    // Totals come from the data, not from the rows above.
    expect([c.D7, c.E7, c.F7, c.G7, c.H7]).toEqual([
      "=SUM('P&L data'!$H$2:$H$500)", `=IFERROR(AVERAGE('P&L data'!$H$2:$H$500),"")`, "=COUNTA('P&L data'!$A$2:$A$500)", "=MIN('P&L data'!$I$2:$I$500)", "=MAX('P&L data'!$I$2:$I$500)",
    ]);
    // Only what adds up is checked: the sum and the count.
    expect([c.D8, c.E8, c.F8, c.G8]).toEqual(["=SUM(D5:D6)-D7", undefined, "=SUM(F5:F6)-F7", undefined]);
    expect(report).toMatchObject({ block: "C4:H8", check_cell: "D8" });
    // A count is a whole number whatever the field's format.
    expect(formatsOf(mutations).map(([range]) => range)).not.toContain("F5:F7");
  });

  it("a second field across the top", () => {
    const { mutations, report } = layoutSummary(
      request({ columns: "Year" }),
      facts({ colField: { col: 4, header: "Year", keys: [2024, 2025], format: null } }),
    );
    expect(cellsOf(mutations)).toEqual({
      A1: "Region", B1: 2024, C1: 2025, D1: "Total",
      A2: "East",
      B2: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A2,Data!$E$2:$E$500,B$1)",
      C2: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A2,Data!$E$2:$E$500,C$1)",
      D2: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A2)",
      A3: "West",
      B3: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A3,Data!$E$2:$E$500,B$1)",
      C3: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A3,Data!$E$2:$E$500,C$1)",
      D3: "=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A3)",
      A4: "Total",
      B4: "=SUMIFS(Data!$H$2:$H$500,Data!$E$2:$E$500,B$1)",
      C4: "=SUMIFS(Data!$H$2:$H$500,Data!$E$2:$E$500,C$1)",
      D4: "=SUM(Data!$H$2:$H$500)",
      A5: "Check (0 = the table holds every row of the data)", D5: "=SUM(B2:C3)-D4",
    });
    expect(report).toMatchObject({ block: "A1:D5", check_cell: "D5", cols: { "2024": "B", "2025": "C", Total: "D" } });
    expect(report.note).toContain("a new Region or Year that is not listed");
    // An average has no check: averages do not add up.
    const avg = layoutSummary(request({ columns: "Year" }), facts({ colField: { col: 4, header: "Year", keys: [2024], format: null }, measures: [{ col: 7, header: "Revenue", fn: "average", format: null }] }));
    expect(avg.report).toMatchObject({ block: "A1:C4", check_cell: null });
  });

  it("keys that SUMIFS would read as a pattern are matched exactly; dates keep their format", () => {
    const { mutations } = layoutSummary(
      request(),
      facts({ rowField: { col: 2, header: "Code", keys: ["A*", ">10", 'say "hi"?', "plain"], format: null } }),
    );
    const c = cellsOf(mutations);
    expect(c.B2).toBe('=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,"=A~*")');
    expect(c.B3).toBe('=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,"=>10")');
    expect(c.B4).toBe('=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,"=say ""hi""~?")');
    expect(c.B5).toBe("=SUMIFS(Data!$H$2:$H$500,Data!$C$2:$C$500,$A5)");
    const dated = layoutSummary(request(), facts({ rowField: { col: 1, header: "Month", keys: [45658, 45689], format: "mmm yy" } }));
    expect(cellsOf(dated.mutations).A2).toBe(45658);
    expect(formatsOf(dated.mutations)).toContainEqual(["A2:A3", { number_format: "mmm yy" }]);
    expect(sheetQualifier("Data")).toBe("Data!");
    expect(sheetQualifier("Q1")).toBe("'Q1'!");
    expect(sheetQualifier("Bob's data")).toBe("'Bob''s data'!");
  });
});

describe("what the agent is told", () => {
  it("says when the table does not account for every row", () => {
    const { report } = layoutSummary(request(), facts());
    const fine: any = summaryResult(report, () => 0);
    expect(fine).toMatchObject({ block: "A1:B5", check_cell: "B5", check_value: 0, row_count: 2 });
    expect(fine.check_warning).toBeUndefined();
    const off: any = summaryResult({ ...report, rows_without_a_key: 11 }, (sheet, row, col) => (sheet === "Summary" && row === 4 && col === 1 ? -51780 : null));
    expect(off.check_value).toBe(-51780);
    expect(off.check_warning).toContain("B5 reads -51780, not 0");
    expect(off.check_warning).toContain("11 data rows have an empty key");
    expect(off.check_warning).toContain("last_row");
    // A table of averages has nothing to check.
    const avg = layoutSummary(request(), facts({ measures: [{ col: 7, header: "Revenue", fn: "average", format: null }] }));
    expect(summaryResult(avg.report, () => 1)).not.toHaveProperty("check_value");
    // Many rows: the map of labels stops at 60, the count does not.
    const many = layoutSummary(request(), facts({ rowField: { col: 2, header: "Region", keys: Array.from({ length: 200 }, (_, i) => `r${i}`), format: null } }));
    expect(Object.keys(many.report.rows)).toHaveLength(60);
    expect(many.report.row_count).toBe(200);
  });
});

describe("what the engine says about the data", () => {
  // A fake row query over a "Data" sheet: Region in C, Year in E, Revenue in H.
  const asked: Array<Record<string, unknown>> = [];
  const io = (groups: Record<string, any[]>, over: Record<string, any> = {}) => ({
    query: async (sheet: string, query: Record<string, unknown>) => {
      asked.push({ sheet, ...query });
      if (sheet !== "Data") return { ok: false as const, error: "unknown_sheet", message: `No sheet named "${sheet}". Sheets: Data` };
      if (typeof query.group_by === "string") {
        const field = query.group_by.toLowerCase();
        if (!(field in groups)) return { ok: false as const, error: "bad_query", message: `unknown column "${query.group_by}"; use a column letter or one of the header texts: C="Region", E="Year", H="Revenue"` };
        const all = groups[field];
        return { ok: true as const, result: { sheet: "Data", header_row: 1, row_range: over.row_range ?? `${query.start_row ?? 2}-${query.end_row ?? 500}`, group_by: { column: field === "region" ? "C" : "E", header: field === "region" ? "Region" : "Year" }, group_count: over.group_count ?? all.length, groups: all.slice(0, Number(query.limit)) } };
      }
      return { ok: true as const, result: { columns: [{ column: "H", header: "Revenue" }], rows: [] } };
    },
    numberFormat: (_sheet: string, _row: number, col: number) => (col === 7 ? "#,##0" : col === 4 ? "General" : null),
    prefetch: async () => {},
  });
  const regions = [
    { key: "West", value: "West", count: 3, sum_Revenue: 900 },
    { key: "east", value: "east", count: 2, sum_Revenue: 500 },
    { key: "(blank)", count: 4, sum_Revenue: 40 },
    { key: "North", value: "North", count: 1, sum_Revenue: 100 },
  ];
  const years = [{ key: "2025", value: 2025, count: 5 }, { key: "2024", value: 2024, count: 5 }];

  it("finds the fields, the rows and the keys, in label order", async () => {
    asked.length = 0;
    const got = await gatherSummaryFacts({ ...request({ columns: "Year" }), firstRow: null, lastRow: 480 }, io({ region: regions, year: years }));
    expect(got).toEqual({
      ok: true,
      skippedBlank: 4,
      facts: {
        sheet: "Data", firstRow: 2, lastRow: 480,
        rowField: { col: 2, header: "Region", keys: ["east", "North", "West"], format: null },
        colField: { col: 4, header: "Year", keys: [2024, 2025], format: null },
        measures: [{ col: 7, header: "Revenue", fn: "sum", format: "#,##0" }],
      },
    });
    // The agent's bound on the data rows reaches every query.
    expect(asked.every((q) => q.end_row === 480)).toBe(true);
    expect(asked[0]).toMatchObject({ group_by: "Region", aggregate: [{ column: "Revenue", fn: "sum" }] });
  });

  it("orders the rows largest first when asked", async () => {
    const got = await gatherSummaryFacts({ ...request({ sort: "value", values: [{ field: "Revenue", fn: "average" }] }) }, io({ region: regions }));
    expect(got.ok && got.facts.rowField.keys).toEqual(["West", "east", "North"]);
    expect(got.ok && got.facts.measures[0].fn).toBe("average");
  });

  it("says why a table cannot be made", async () => {
    const why = async (req: SummaryRequest, groups: Record<string, any[]>, over = {}) => {
      const got = await gatherSummaryFacts(req, io(groups, over));
      return got.ok ? "" : `${got.error}: ${got.message}`;
    };
    expect(await why(request({ sourceSheet: "Dta" }), { region: regions })).toContain('unknown_sheet: No sheet named "Dta"');
    expect(await why(request({ rows: "Regoin" }), { region: regions })).toContain('unknown column "Regoin"');
    const many = Array.from({ length: 501 }, (_, i) => ({ key: `c${i}`, value: `c${i}`, count: 1 }));
    expect(await why(request(), { region: many }, { group_count: 9000 })).toContain('too_many_groups: "Region" has 9000 distinct values');
    expect(await why(request(), { region: [{ key: "(blank)", count: 9 }] })).toContain("no_groups");
    expect(await why(request(), { region: regions }, { row_range: "2-1" })).toContain("no_data");
    expect(await why(request({ columns: "Year" }), { region: regions, year: Array.from({ length: 61 }, (_, i) => ({ key: `${i}`, value: i, count: 1 })) })).toContain("a field across the top takes up to 60");
  });
});

describe("the tool call", () => {
  it("is read into a request", () => {
    const call = { source_sheet: "Data", rows: "Region", columns: " Year ", values: [{ field: "Revenue", fn: "AVG" }], dest_sheet: "Summary", dest: "$B$3", title: "By region", sort: "value", last_row: 480 };
    expect(interpretToolCall("create_summary_table", call)).toEqual({
      kind: "summary", sourceSheet: "Data", headerRow: null, firstRow: null, lastRow: 480, rows: "Region", columns: "Year",
      values: [{ field: "Revenue", fn: "average" }], destSheet: "Summary", dest: { row: 2, col: 1 }, title: "By region", sort: "value",
    });
    // The default is a sum, by label.
    expect(interpretToolCall("create_summary_table", { ...call, columns: undefined, sort: undefined, values: [{ field: "Revenue" }, { field: "Deal", fn: "count" }] })).toMatchObject({ columns: null, sort: "label", values: [{ field: "Revenue", fn: "sum" }, { field: "Deal", fn: "count" }] });
    const why = (over: object) => {
      const r = interpretToolCall("create_summary_table", { ...call, ...over });
      return r.kind === "ignored" ? r.reason : "";
    };
    expect(why({ source_sheet: "" })).toContain("missing source_sheet");
    expect(why({ rows: null })).toContain("missing 'rows'");
    expect(why({ dest: "top" })).toContain("'dest_sheet' and 'dest'");
    expect(why({ values: [] })).toContain("1 to 6 entries");
    expect(why({ values: [{ field: "Revenue", fn: "median" }] })).toContain("sum, count, average, min, max");
    expect(why({ values: [{ field: "Revenue" }, { field: "Cost" }] })).toContain("give ONE value");
  });
});
