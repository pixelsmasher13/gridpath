/**
 * The MCP server: tool schemas served verbatim from tools.json (plus the
 * batch/save tools), calls dispatched to ToolHandlers. Transport-agnostic;
 * the bin wires stdio, tests wire an in-memory pair.
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { ToolHandlers, toolSchemas, type HandlerOptions } from "../tools/handlers";

export const INSTRUCTIONS = `GridPath edits real .xlsx files. Every tool takes \`path\`.

Start with describe_workbook (structure, headers, sections) and find_rows before reading or writing; query_rows to filter, project or group the rows of a data sheet by column values (engine-served, any size); read_range for exact cells you have located. Write formulas, not computed numbers, and raw numbers, never formatted strings (formatting is set_format's job). Use run_script for anything loop-shaped; reads inside scripts are free.

The engine evaluates like Excel 365: dynamic arrays and spills (FILTER, UNIQUE, SORT, SORTBY, SEQUENCE, VSTACK/HSTACK, TAKE/DROP, CHOOSECOLS/CHOOSEROWS, TOCOL, WRAPROWS), LET/LAMBDA with MAP/BYROW/SCAN, XLOOKUP, the *IFS family, SUMPRODUCT with --(cond), ROW(range) arrays, TEXTJOIN over arrays, OFFSET/INDIRECT, date and finance functions. RANK, PERCENTILE, QUARTILE, MEDIAN, STDEV, SUBTOTAL and AGGREGATE work (all 19 AGGREGATE functions and 8 options, so AGGREGATE(15,6,ROW(rng)/(cond),k) finds the k-th matching row). Not implemented (#NAME?): GETPIVOTDATA, HYPERLINK, GROUPBY, PIVOTBY — UNIQUE + SUMIFS instead of GROUPBY. A derived table (ranked / filtered / combined view of data sheets) is ONE spilled formula — LET + VSTACK + FILTER + SORTBY + TAKE — in its top-left cell, not rows of INDEX links.

Several workbooks can be worked on together, and a formula in one can read another. Write the other file's name in brackets before the sheet: =[NOW.xlsx]Model!B23, or with quotes when the file or sheet name has a space or punctuation: ='[NOW 2026.xlsx]Income Statement'!B23. Only the file name is needed when the file is in the same folder or already open (otherwise put its folder before the bracket, inside the quotes). The link is really evaluated: the other workbook is opened, and an edit there, including rows or columns inserted above the cell, recalculates every workbook that reads it. Find the cell first (describe_workbook, find_rows and read_range on the other file's own path), and prefer a link that finds its row and column by label, =INDEX('[NOW.xlsx]Model'!$A$1:$CZ$2000, MATCH("Total revenue", '[NOW.xlsx]Model'!$A$1:$A$2000, 0), MATCH(C$1, '[NOW.xlsx]Model'!$A$5:$CZ$5, 0)), over one fixed cell per period when the two workbooks lay their periods out differently. Size those ranges to cover the other sheet's used_range, and check with find_rows that the label is the first of its kind in that column, because MATCH takes the first. A write whose link leads nowhere (no such file or sheet) is refused with the reason and the sheet names that do exist. list_workbooks shows the workbooks in a folder; create_workbook makes a new empty one to pull figures into. save_workbook writes the links into the file the way Excel does (the other file's path and the last values read), so the workbook opens in Excel with its figures in place, and the links are live again when it is opened here. When the workbook that is read has unsaved row, column or sheet changes, save that one first (the refusal lists it under save_first); and after saving a workbook whose rows moved, save the workbooks the result lists under also_save.

Writes do NOT touch the file: they accumulate as pending batches and every write returns a readback of what landed plus a review_url. When you finish a change, ALWAYS show the user the review_url — it opens a local page where they see every changed cell (before/after), can reject a batch, and save. Call save_workbook when the edit is complete (it patches only the parts you changed; everything else in the file stays byte-identical). If the server was started with --review-required, save_workbook does not write: the user saves from the review page. reject_batch drops a pending batch.`;

export function createServer(opts: HandlerOptions & { version?: string } = {}): { server: Server; handlers: ToolHandlers } {
  const handlers = new ToolHandlers(opts);
  const server = new Server({ name: "gridpath", version: opts.version ?? "0.0.1" }, { capabilities: { tools: {} }, instructions: INSTRUCTIONS });

  // Annotations let hosts treat reads gently (fewer approval prompts) and know
  // that nothing but save_workbook changes a file (create_workbook makes a
  // new, empty one and never replaces anything).
  const READ_ONLY = new Set(["describe_workbook", "find_rows", "query_rows", "read_range", "list_batches", "list_workbooks"]);
  const tools = toolSchemas().map((t) => ({
    name: t.name,
    description: t.description,
    inputSchema: t.input_schema,
    annotations: READ_ONLY.has(t.name)
      ? { title: t.name.replace(/_/g, " "), readOnlyHint: true, destructiveHint: false, openWorldHint: false }
      : t.name === "save_workbook"
        ? { title: "save workbook", readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false }
        : { title: t.name.replace(/_/g, " "), readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
  }));
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));
  server.setRequestHandler(CallToolRequestSchema, async (req) => {
    const result = await handlers.call(req.params.name, (req.params.arguments ?? {}) as Record<string, unknown>);
    const isError = typeof result === "object" && result !== null && (result as any).ok === false;
    return { content: [{ type: "text", text: JSON.stringify(result) }], isError };
  });

  return { server, handlers };
}
