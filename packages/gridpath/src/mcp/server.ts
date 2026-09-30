/**
 * The MCP server: tool schemas served verbatim from tools.json (plus the
 * batch/save tools), calls dispatched to ToolHandlers. Transport-agnostic;
 * the bin wires stdio, tests wire an in-memory pair.
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { ToolHandlers, toolSchemas, type HandlerOptions } from "../tools/handlers";

export const INSTRUCTIONS = `GridPath edits real .xlsx files. Every tool takes \`path\`.

Start with describe_workbook (structure, headers, sections) and find_rows before reading or writing; read_range for exact values. Write formulas, not computed numbers, and raw numbers, never formatted strings (formatting is set_format's job). Use run_script for anything loop-shaped.

Writes do NOT touch the file: they accumulate as pending batches and every write returns a readback of what landed plus a review_url. When you finish a change, ALWAYS show the user the review_url — it opens a local page where they see every changed cell (before/after), can reject a batch, and save. Call save_workbook when the edit is complete (it patches only the parts you changed; everything else in the file stays byte-identical). If the server was started with --review-required, save_workbook does not write: the user saves from the review page. reject_batch drops a pending batch.`;

export function createServer(opts: HandlerOptions & { version?: string } = {}): { server: Server; handlers: ToolHandlers } {
  const handlers = new ToolHandlers(opts);
  const server = new Server({ name: "gridpath", version: opts.version ?? "0.0.1" }, { capabilities: { tools: {} }, instructions: INSTRUCTIONS });

  const tools = toolSchemas().map((t) => ({ name: t.name, description: t.description, inputSchema: t.input_schema }));
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));
  server.setRequestHandler(CallToolRequestSchema, async (req) => {
    const result = await handlers.call(req.params.name, (req.params.arguments ?? {}) as Record<string, unknown>);
    const isError = typeof result === "object" && result !== null && (result as any).ok === false;
    return { content: [{ type: "text", text: JSON.stringify(result) }], isError };
  });

  return { server, handlers };
}
