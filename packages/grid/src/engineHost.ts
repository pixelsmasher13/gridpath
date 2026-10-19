/**
 * The WASM engine behind a small command protocol. Runs inside the engine
 * worker (engine.worker.ts) or, with `worker: false`, on the page itself.
 * One host holds one workbook. Commands mirror the desktop's
 * calc_engine_* Tauri commands; rows and columns are 1-based.
 */
import init, { Workbook, blankXlsx, patch } from "../engine/gridpath_engine.js";

export type CellDelta = { sheet: number; row: number; column: number; value: unknown };

export type HostRequest = { id: number; cmd: string; args?: any };
export type HostResponse = { id: number; ok: true; result: unknown } | { id: number; ok: false; error: string };

/** `evaluateEdited` output as value deltas. A `full` reply (after a
 * structural or sheet-level change) carries none: the grid reloads its
 * layout and refetches the cells it draws. */
function deltas(json: string): CellDelta[] {
  const res = JSON.parse(json);
  if (res.full) return [];
  return (res.cells ?? []).map((c: any) => ({ sheet: c.sheet, row: c.r, column: c.c, value: c.v ?? null }));
}

export class EngineHost {
  private wb: Workbook | null = null;
  private initPromise: Promise<unknown> | null = null;

  private book(): Workbook {
    if (!this.wb) throw new Error("no workbook loaded");
    return this.wb;
  }

  async handle(cmd: string, a: any = {}): Promise<unknown> {
    if (cmd === "init") {
      this.initPromise ??= init(a.wasmUrl ? { module_or_path: a.wasmUrl } : undefined);
      await this.initPromise;
      return true;
    }
    if (!this.initPromise) throw new Error("engine not initialised");
    await this.initPromise;
    switch (cmd) {
      case "blank":
        return blankXlsx(a.sheets?.length ? a.sheets : ["Sheet1"]);
      case "load": {
        this.wb?.free();
        this.wb = Workbook.openDeferred(a.bytes);
        return { sheets: this.wb.sheetNames() };
      }
      case "layout":
        return this.book().workbookLayout();
      case "window":
        return this.book().sheetWindow(a.sheet, a.r0, a.r1, a.c0, a.c1);
      case "evaluatePending":
        return JSON.parse(this.book().evaluatePending()) as { ms: number; cells: CellDelta[] };
      case "editBatch": {
        const wb = this.book();
        for (const e of a.edits as Array<{ sheet: number; row: number; column: number; value: string }>) wb.setInput(e.sheet, e.row, e.column, e.value);
        return deltas(wb.evaluateEdited());
      }
      case "sheetOp": {
        const wb = this.book();
        const names = wb.sheetNames();
        if (a.op === "add") {
          const i = wb.newSheet();
          wb.renameSheet(i, a.name);
        } else if (a.op === "rename") wb.renameSheet(names.indexOf(a.name), a.newName);
        else if (a.op === "delete") wb.deleteSheet(names.indexOf(a.name));
        const cells = deltas(wb.evaluateEdited());
        return { sheets: wb.sheetNames(), cells };
      }
      case "structural": {
        const wb = this.book();
        if (a.op === "insert_rows") wb.insertRows(a.sheet, a.index, a.count);
        else if (a.op === "delete_rows") wb.deleteRows(a.sheet, a.index, a.count);
        else if (a.op === "insert_columns") wb.insertColumns(a.sheet, a.index, a.count);
        else if (a.op === "delete_columns") wb.deleteColumns(a.sheet, a.index, a.count);
        else throw new Error(`unknown structural op ${a.op}`);
        return deltas(wb.evaluateEdited());
      }
      case "setStyle": {
        const wb = this.book();
        const { sheet, r0, r1, c0, c1, patch: p } = a;
        const h = r1 - r0 + 1;
        const w = c1 - c0 + 1;
        const set = (path: string, value: string) => wb.setRangeStyle(sheet, r0, c0, h, w, path, value);
        if (p.bold !== undefined) set("font.b", String(p.bold));
        if (p.italic !== undefined) set("font.i", String(p.italic));
        if (p.underline !== undefined) set("font.u", String(p.underline));
        if (p.strike !== undefined) set("font.strike", String(p.strike));
        if (p.font_color) set("font.color", p.font_color);
        if (p.background_color !== undefined) set("fill.color", p.background_color ?? "");
        if (p.number_format) set("num_fmt", p.number_format);
        if (p.horizontal_align) set("alignment.horizontal", p.horizontal_align);
        if (p.vertical_align) set("alignment.vertical", p.vertical_align === "middle" ? "center" : p.vertical_align);
        if (p.wrap_text !== undefined) set("alignment.wrap_text", String(p.wrap_text));
        return h * w;
      }
      case "defineName": {
        const wb = this.book();
        wb.newDefinedName(a.name, String(a.formula).replace(/^=/, ""));
        return deltas(wb.evaluateEdited());
      }
      case "find":
        return JSON.parse(this.book().find(a.query, a.sheet ?? undefined, a.limit ?? undefined));
      // Reads for the agent tools (same calls the headless MCP session makes).
      case "cell":
        return this.book().cell(a.sheet, a.row, a.col);
      case "workbookIndex":
        return this.book().workbookIndex();
      case "queryRows":
        return this.book().queryRows(JSON.stringify(a.query));
      case "recalculated":
        // Formula cells this session changed, with current values (save-time cache refresh).
        return JSON.parse(this.book().recalculatedCells());
      case "sheetCells":
        // Bytes, so the script sandbox can take them without a copy.
        return new TextEncoder().encode(this.book().sheetCells(a.sheet));
      case "patch":
        return patch(a.base, a.patch);
      case "unload":
        this.wb?.free();
        this.wb = null;
        return true;
      default:
        throw new Error(`unknown engine command ${cmd}`);
    }
  }
}
