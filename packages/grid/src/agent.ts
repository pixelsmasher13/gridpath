/**
 * gridpath/agent: the agent tool definitions and instructions, with no
 * browser or React code, for the server side of an agent loop. The server
 * should send these itself (never take a system prompt or tool list from
 * the browser), and the browser runs the model's tool calls with
 * `sheet.ai.run()`.
 */
import toolsJson from "../../gridpath/src/tools/tools.json";

export type ToolDefinition = { name: string; description: string; input_schema: Record<string, unknown> };
export type OpenAIToolDefinition = { type: "function"; function: { name: string; description: string; parameters: Record<string, unknown> } };

// Tools that need the desktop app or the network.
const EXCLUDED = new Set(["done", "read_reference", "fetch_web", "edgar_lookup", "keep_pages", "read_source", "stage_data"]);

/** GridPath's spreadsheet tools, in Anthropic Messages format (default) or OpenAI's. */
export function agentTools(format?: "anthropic"): ToolDefinition[];
export function agentTools(format: "openai"): OpenAIToolDefinition[];
export function agentTools(format: "anthropic" | "openai" = "anthropic"): ToolDefinition[] | OpenAIToolDefinition[] {
  const defs = (toolsJson as ToolDefinition[]).filter((t) => !EXCLUDED.has(t.name));
  if (format === "openai") return defs.map((t) => ({ type: "function" as const, function: { name: t.name, description: t.description, parameters: t.input_schema } }));
  return defs.map((t) => ({ ...t }));
}

export const AGENT_INSTRUCTIONS = `You are editing the spreadsheet open in this app through GridPath's tools.

Start with describe_workbook (structure, headers, sections) and find_rows before reading or writing; query_rows to filter, project or group the rows of a data sheet by column values (engine-served, any size); read_range for exact cells you have located. Write formulas, not computed numbers, and raw numbers, never formatted strings (formatting is set_format's job). Use run_script for anything loop-shaped; reads inside scripts are free.

The engine evaluates like Excel 365: dynamic arrays and spills (FILTER, UNIQUE, SORT, SORTBY, SEQUENCE, VSTACK/HSTACK, TAKE/DROP, CHOOSECOLS/CHOOSEROWS, TOCOL, WRAPROWS), LET/LAMBDA with MAP/BYROW/SCAN, XLOOKUP, the *IFS family, SUMPRODUCT with --(cond), ROW(range) arrays, TEXTJOIN over arrays, OFFSET/INDIRECT, date and finance functions. RANK, PERCENTILE, QUARTILE, MEDIAN and STDEV work. Not implemented (#NAME?): AGGREGATE, GETPIVOTDATA, HYPERLINK, GROUPBY, PIVOTBY — use SUBTOTAL or LARGE/SMALL over IFERROR(rng, "") instead of AGGREGATE, UNIQUE + SUMIFS instead of GROUPBY. A derived table (ranked / filtered / combined view of data sheets) is ONE spilled formula — LET + VSTACK + FILTER + SORTBY + TAKE — in its top-left cell, not rows of INDEX links.

Each write applies to the open workbook at once and returns a readback of what landed. The user sees your changes highlighted and accepts or rejects them; you do not save. Check the readback (write_check, values) before moving on, and say what you changed when you are done.`;
