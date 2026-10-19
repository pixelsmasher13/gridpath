/**
 * GridPath's agent tools over a live spreadsheet: the same tool set and
 * result shapes as the `gridpath` MCP server, but acting on the grid the
 * user is looking at. Bring your own model and loop: hand `tools()` to it,
 * run each tool call through `run()`, and give the result back.
 *
 * Every write is applied at once and kept as a pending batch, highlighted
 * in the grid, until `accept` or `reject`. `save()` writes what is applied.
 */
import { interpretToolCall } from "../../../gridpath/src/core/toolToMutation";
import { expandCopy } from "../../../gridpath/src/core/copyRange";
import { describeWorkbookPayload, findRowsInIndex, firstLabelIn } from "../../../gridpath/src/core/workbookIndex";
import { executeSheetScript } from "../../../gridpath/src/core/scriptRunner";
import { stripImplicitIntersection, parseA1Range } from "../../../gridpath/src/workbook/snapshot";
import type { ChangeBatch, UniverMutation } from "../../../gridpath/src/core/types";
import type { UniverGridHandle } from "../../../../src/screens/SpreadsheetScreen/components/UniverGrid";
import { reverseMutation } from "../../../../src/screens/SpreadsheetScreen/agent/batchMutations";
import type { BrowserEngine } from "../browserEngine";
import { applyMutations } from "./apply";

import { AGENT_INSTRUCTIONS, agentTools, type OpenAIToolDefinition, type ToolDefinition } from "../agent";

export type { ToolDefinition, OpenAIToolDefinition } from "../agent";

/** One cell as tools report it (0-based internally, A1 in `cell`). */
export type ReadCell = { cell: string; value: string | number | boolean | null; formula: string | null; display?: string };

/** A pending or accepted change from one tool call. */
export type AgentBatch = {
  id: string;
  tool: string;
  status: "pending" | "accepted";
  created_at: string;
  /** Cells it wrote, cleared or reformatted. */
  cells: Array<{ sheet: string; row: number; col: number }>;
  mutations: UniverMutation[];
};

export const INSTRUCTIONS = AGENT_INSTRUCTIONS;

type CellRef = { sheet: string; row: number; col: number };

/** Where an earlier batch's cell is after a later structural mutation
 * (null when the row, column or sheet is gone). */
function moveCell(c: CellRef, m: UniverMutation): CellRef | null {
  switch (m.type) {
    case "insert_rows":
      return m.sheet === c.sheet && c.row >= m.before ? { ...c, row: c.row + m.count } : c;
    case "delete_rows":
      if (m.sheet !== c.sheet || c.row < m.start) return c;
      return c.row < m.start + m.count ? null : { ...c, row: c.row - m.count };
    case "insert_columns":
      return m.sheet === c.sheet && c.col >= m.before ? { ...c, col: c.col + m.count } : c;
    case "delete_columns":
      if (m.sheet !== c.sheet || c.col < m.start) return c;
      return c.col < m.start + m.count ? null : { ...c, col: c.col - m.count };
    case "rename_sheet":
      return c.sheet === m.old_name ? { ...c, sheet: m.new_name } : c;
    case "delete_sheet":
      return c.sheet === m.name ? null : c;
    default:
      return c;
  }
}

const a1 = (row: number, col: number) => {
  let s = "";
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return `${s}${row + 1}`;
};

/** The agent surface of a spreadsheet (`sheet.ai`). */
export interface SheetAgent {
  /** Tool definitions, in Anthropic Messages format (default) or OpenAI's. */
  tools(format?: "anthropic"): ToolDefinition[];
  tools(format: "openai"): OpenAIToolDefinition[];
  /** How to use the tools: put it in your system prompt. */
  instructions(): string;
  /** The workbook overview (describe_workbook for every sheet), for the first turn. */
  context(): Promise<string>;
  /** Run one tool call; the result is JSON for the model. Never throws. */
  run(name: string, input?: Record<string, unknown>): Promise<unknown>;
  /** Pending and accepted batches, oldest first. */
  list(): AgentBatch[];
  /** Called whenever batches are added, accepted or rejected. */
  onChange(fn: (batches: AgentBatch[]) => void): () => void;
  /** Keep a batch (or every pending one). */
  accept(id?: string): void;
  /**
   * Undo a batch and every batch applied after it, newest first (with no
   * id: every pending batch and those after it). Undoing in reverse order
   * is what keeps rows and columns that later batches moved correct.
   * Resolves with the ids undone.
   */
  reject(id?: string): Promise<string[]>;
}

export class AgentToolkit implements SheetAgent {
  private batches: AgentBatch[] = [];
  private seq = 0;
  private listeners = new Set<(batches: AgentBatch[]) => void>();
  private indexCache: { version: string; index: any } | null = null;

  constructor(
    private readonly grid: UniverGridHandle,
    private readonly engine: BrowserEngine,
  ) {}

  /** Tool definitions, in Anthropic Messages format (default) or OpenAI's. */
  tools(format?: "anthropic"): ToolDefinition[];
  tools(format: "openai"): OpenAIToolDefinition[];
  tools(format: "anthropic" | "openai" = "anthropic"): ToolDefinition[] | OpenAIToolDefinition[] {
    return format === "openai" ? agentTools("openai") : agentTools();
  }

  /** How to use the tools: put it in your system prompt. */
  instructions(): string {
    return INSTRUCTIONS;
  }

  /** The workbook overview (describe_workbook for every sheet), for the first turn. */
  async context(): Promise<string> {
    return JSON.stringify(await this.run("describe_workbook", {}));
  }

  /** Run one tool call; the result is JSON for the model. Never throws. */
  async run(name: string, input: Record<string, unknown> = {}): Promise<unknown> {
    try {
      return await this.dispatch(name, input);
    } catch (e) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
  }

  /** Pending and accepted batches, oldest first. */
  list(): AgentBatch[] {
    return this.batches.map((b) => ({ ...b }));
  }

  /** Called whenever batches are added, accepted or rejected. */
  onChange(fn: (batches: AgentBatch[]) => void): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  /** Keep a batch (or every pending one): its highlight goes. */
  accept(id?: string): void {
    for (const b of this.batches) if (b.status === "pending" && (id === undefined || b.id === id)) b.status = "accepted";
    this.changed();
  }

  async reject(id?: string): Promise<string[]> {
    const from = id === undefined ? this.batches.findIndex((b) => b.status === "pending") : this.batches.findIndex((b) => b.id === id);
    if (from < 0) return [];
    const undone = this.batches.slice(from).reverse();
    for (const b of undone) for (const m of [...b.mutations].reverse()) reverseMutation(this.grid, m);
    this.batches = this.batches.slice(0, from);
    await this.grid.whenCalculated(30_000);
    this.changed();
    return undone.map((b) => b.id);
  }

  /** Every applied batch, for the save mirror. */
  appliedBatches(): ChangeBatch[] {
    return this.batches.map((b) => ({ id: b.id, prompt: b.tool, justification: b.tool, mutations: b.mutations, status: b.status, created_at: b.created_at }));
  }

  // ---- internals ----

  private changed(): void {
    const pending = this.batches.filter((b) => b.status === "pending").flatMap((b) => b.cells);
    (this.grid as any).setHighlights?.(pending);
    const snapshot = this.list();
    for (const fn of this.listeners) fn(snapshot);
  }

  private sheetIndex(name: string): number {
    const i = this.grid.getSheetNames().indexOf(name);
    if (i < 0) throw new Error(`Sheet not found: ${name}`);
    return i;
  }

  private async index(): Promise<any> {
    const version = this.grid.getContentVersion();
    if (this.indexCache?.version !== version) this.indexCache = { version, index: await this.engine.workbookIndex() };
    return this.indexCache.index;
  }

  /** One cell (0-based) from the engine, as the MCP session reports it. */
  private async cell(sheet: string, row: number, col: number): Promise<ReadCell> {
    const c = await this.engine.cellAt(this.sheetIndex(sheet), row + 1, col + 1);
    const formula = typeof c.content === "string" && c.content.startsWith("=") ? stripImplicitIntersection(c.content) : null;
    const out: ReadCell = { cell: a1(row, col), value: (c.value as ReadCell["value"]) ?? null, formula };
    if (c.display && String(c.value) !== c.display) out.display = c.display;
    return out;
  }

  private async readRange(sheet: string, range: string, limit = 500): Promise<{ cells: ReadCell[]; truncated: boolean }> {
    const r = parseA1Range(range);
    if (!r) throw new Error(`Bad range: ${range}`);
    const win = await this.engine.readRect(this.sheetIndex(sheet), r.startRow + 1, r.endRow + 1, r.startColumn + 1, r.endColumn + 1);
    const cells: ReadCell[] = [];
    const sorted = [...(win?.cells ?? [])].sort((x: any, y: any) => x.r - y.r || x.c - y.c);
    for (const w of sorted) {
      const value = (w.live !== undefined ? w.live : w.v) ?? null;
      const formula = typeof w.f === "string" && w.f ? stripImplicitIntersection(w.f) : null;
      if (value === null && !formula) continue;
      if (cells.length >= limit) return { cells, truncated: true };
      const out: ReadCell = { cell: a1(w.r - 1, w.c - 1), value, formula };
      if (w.d && String(value) !== w.d) out.display = w.d;
      cells.push(out);
    }
    return { cells, truncated: false };
  }

  private async dispatch(name: string, input: Record<string, unknown>): Promise<unknown> {
    const result = interpretToolCall(name, input);
    await this.grid.whenCalculated(60_000);
    switch (result.kind) {
      case "describe_workbook":
        return describeWorkbookPayload(await this.index(), result.sheet);
      case "query_rows":
        return this.engine.queryRows({ ...result.query, sheet: this.sheetIndex(result.sheet) });
      case "find_rows": {
        const { matches, total } = findRowsInIndex(await this.index(), result.query, result.sheet, result.max_results);
        const out = [];
        for (const m of matches as any[]) {
          const o: any = { sheet: m.sheet, row: m.row, label: m.label };
          if (m.section !== null) o.section = m.section;
          if (m.column !== undefined) o.column = m.column;
          if (m.sampleCol !== null) {
            const c = await this.cell(m.sheet, m.row - 1, m.sampleCol);
            if (c.value !== null || c.formula) o.sample = c.formula ? `${c.cell} = ${c.formula} → ${c.value ?? ""}` : `${c.cell} = ${c.value}`;
          }
          out.push(o);
        }
        return { matches: out, total, truncated: total > matches.length };
      }
      case "read": {
        const { cells, truncated } = await this.readRange(result.sheet, result.range);
        return { sheet: result.sheet, range: result.range, cells, truncated };
      }
      case "mutations":
        return this.applyAndReadBack(name, result.mutations);
      case "copy": {
        // The copy reads its source from the grid: fetch those tiles first.
        const src = parseA1Range(result.source);
        if (src) await this.grid.prefetchRange?.(result.sheet, src.startRow, src.startColumn, src.endRow, src.endColumn);
        const expansion = expandCopy(
          { getCell: (sh, row, col) => this.grid.getCell(sh, row, col), getCellFormat: (sh, row, col) => this.grid.getCellFormat(sh, row, col) ?? {} },
          result,
        );
        if (!expansion.ok) return expansion;
        const back = (await this.applyAndReadBack(name, expansion.mutations)) as Record<string, unknown>;
        return { ...back, dest_range: expansion.dest_range, copied_cells: expansion.copied_cells };
      }
      case "script": {
        const names = this.grid.getSheetNames();
        const run = await executeSheetScript({ getWorkbookSnapshot: () => null }, result.code, {
          source: { sheetNames: names, fetchSheet: (n) => this.engine.sheetCellsBuffer(names.indexOf(n)) },
        });
        if (!run.ok) return { ok: false, error: run.error, script_logs: run.logs };
        const back = (await this.applyAndReadBack(name, run.mutations)) as Record<string, unknown>;
        return { ...back, writes: run.writes, script_logs: run.logs };
      }
      case "ignored":
        return { ok: false, error: result.reason };
      default:
        return { ok: false, error: `${name} is not available here` };
    }
  }

  private async applyAndReadBack(tool: string, mutations: UniverMutation[]): Promise<unknown> {
    if (mutations.length === 0) return { ok: false, error: "no cells to write — check the arguments" };
    const applied = await applyMutations(this.grid, mutations);
    // Earlier batches' cells follow rows, columns and sheets this one moved.
    for (const m of applied.mutations) {
      if (!["insert_rows", "delete_rows", "insert_columns", "delete_columns", "rename_sheet", "delete_sheet"].includes(m.type)) continue;
      for (const b of this.batches) b.cells = b.cells.map((c) => moveCell(c, m)).filter((c): c is CellRef => c !== null);
    }
    const batch: AgentBatch = { id: `b${++this.seq}`, tool, status: "pending", created_at: new Date().toISOString(), cells: applied.touched, mutations: applied.mutations };
    this.batches.push(batch);
    await this.grid.whenCalculated(60_000);
    this.changed();

    // Read back what landed (the MCP server's shape).
    const touched = new Map<string, { sheet: string; row: number; col: number }>();
    const expected = new Set<string>();
    for (const m of applied.mutations) {
      if (m.type !== "set_cell") continue;
      const k = `${m.address.sheet}!${m.address.row},${m.address.col}`;
      touched.set(k, m.address);
      if (m.new_value != null || m.new_formula != null) expected.add(k);
    }
    const cells: Array<ReadCell & { sheet: string }> = [];
    let landed = 0;
    for (const [k, a] of [...touched].slice(0, 200)) {
      const c = await this.cell(a.sheet, a.row, a.col);
      cells.push({ sheet: a.sheet, ...c });
      if (expected.has(k) && (c.value !== null || c.formula)) landed++;
    }
    const rows = new Map<string, { sheet: string; row: number }>();
    for (const a of touched.values()) rows.set(`${a.sheet}!${a.row}`, a);
    const row_map: Array<{ sheet: string; row: number; label: string }> = [];
    for (const { sheet, row } of [...rows.values()].slice(0, 100)) {
      const vals = [];
      for (let col = 0; col < 4; col++) vals.push((await this.cell(sheet, row, col)).value);
      const label = firstLabelIn(vals);
      if (label) row_map.push({ sheet, row: row + 1, label });
    }
    row_map.sort((a, b) => a.row - b.row);
    const checked = Math.min(expected.size, 200);
    const write_check = checked === 0 ? "ok" : landed === 0 ? "FAILED" : landed < checked ? "partial" : "ok";
    return {
      ok: write_check !== "FAILED",
      batch_id: batch.id,
      write_check,
      cells,
      row_map,
      pending_batches: this.batches.filter((b) => b.status === "pending").length,
      note: "Applied to the open workbook and highlighted for the user, who accepts or rejects it.",
    };
  }
}
