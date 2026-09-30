/**
 * gridpath — let your agent edit Excel files without butchering them.
 *
 * Public surface (v0): the headless workbook session used by the MCP server.
 * The engine (IronCalc evaluation + surgical xlsx patching) is WASM under
 * ./engine and is loaded lazily on first use.
 */
export * from "./core/types";
export { interpretToolCall } from "./core/toolToMutation";
export { buildWorkbookIndex, describeWorkbookPayload, findRowsInIndex } from "./core/workbookIndex";
export { buildWorkbookPatch } from "./core/surgicalPatch";
export { ToolHandlers, toolSchemas } from "./tools/handlers";
export { createServer } from "./mcp/server";
export { startReviewServer } from "./review/api";
export { WorkbookSession } from "./workbook/session";
