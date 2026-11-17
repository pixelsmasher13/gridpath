/**
 * Tool handlers over headless workbook sessions. Every tool takes `path`;
 * a session is opened on first use and kept for the process lifetime.
 *
 * Nothing touches the file until `save_workbook`. Writes accumulate as
 * pending batches (the same ChangeBatch the desktop app reviews), mirrored
 * to disk next to the workbook so the review link outlives this process;
 * `reject_batch` drops one and replays the rest onto a fresh model.
 *
 * The open workbooks form a workspace (`workbook/workspace.ts`): a formula
 * written into one may read another by file name, and an edit to the one
 * that is read recalculates the ones that read it.
 */
import fs from "node:fs";
import path from "node:path";
import { interpretToolCall } from "../core/toolToMutation";
import { expandCopy } from "../core/copyRange";
import { describeWorkbookPayload, findRowsInIndex, firstLabelIn } from "../core/workbookIndex";
import type { ChangeBatch, UniverMutation } from "../core/types";
import { isLinkSheetName } from "../workbook/links";
import { WorkbookSession, a1Of, loadEngine, type ReadCell } from "../workbook/session";
import { clearStored, saveStored } from "../workbook/store";
import { LinkError, Workspace, isMirrorBatch, linkFailed } from "../workbook/workspace";
import { runScript } from "./script";
import toolsJson from "./tools.json";

export type ToolSchema = { name: string; description: string; input_schema: any };

/** Tools the headless server offers (subset of the desktop agent's). */
const EXCLUDED = new Set(["done", "read_reference", "fetch_web", "edgar_lookup", "keep_pages", "read_source", "stage_data"]);

const PATH_PROP = {
  type: "string",
  description: "Path of the .xlsx to operate on (absolute, or ~/… for the user's home directory).",
};

export function toolSchemas(): ToolSchema[] {
  const base = (toolsJson as ToolSchema[])
    .filter((t) => !EXCLUDED.has(t.name))
    .map((t) => ({
      name: t.name,
      description: t.description,
      input_schema: {
        ...t.input_schema,
        properties: { path: PATH_PROP, ...(t.input_schema?.properties ?? {}) },
        required: ["path", ...((t.input_schema?.required as string[] | undefined) ?? [])],
      },
    }));
  const extra: ToolSchema[] = [
    {
      name: "list_workbooks",
      description:
        "List the Excel workbooks (.xlsx, .xlsm) in a folder: path, size, when it was last modified, whether it is open in this session with unsaved changes, and which other workbooks it reads or is read by. Use it to find the files to work across; any of them can then be described, searched and read with the other tools, and a formula written into one can read another (see the server instructions on linking workbooks).",
      input_schema: {
        type: "object",
        properties: { path: { type: "string", description: "The folder to list (absolute, or ~/… for the user's home directory). Its subfolders are not searched." } },
        required: ["path"],
      },
    },
    {
      name: "create_workbook",
      description:
        "Create a new, empty .xlsx file on disk, for example the workbook that will pull figures together from several others. Never overwrites: fails when the file already exists. After this, write into it like any other workbook.",
      input_schema: {
        type: "object",
        properties: {
          path: { type: "string", description: "Where to create it (absolute, or ~/…). `.xlsx` is added when missing." },
          sheets: { type: "array", items: { type: "string" }, description: "Sheet names, in order. Default: one sheet, Sheet1." },
        },
        required: ["path"],
      },
    },
    {
      name: "list_batches",
      description: "List this workbook's pending (unsaved) change batches with their mutation counts, and the review link.",
      input_schema: { type: "object", properties: { path: PATH_PROP }, required: ["path"] },
    },
    {
      name: "reject_batch",
      description: "Drop one pending batch. The remaining batches are replayed onto a fresh model, so later edits that depended on it may change.",
      input_schema: { type: "object", properties: { path: PATH_PROP, batch_id: { type: "string" } }, required: ["path", "batch_id"] },
    },
    {
      name: "save_workbook",
      description:
        "Write every pending batch into the .xlsx on disk through the surgical patcher: only the parts an edit touched are rewritten, everything else stays byte-identical. This is the ONLY tool that modifies the file. When the server runs with --review-required, this returns the review link instead and the user saves from there.",
      input_schema: {
        type: "object",
        properties: {
          path: PATH_PROP,
          as: { type: "string", description: "Optional: write the result to this path instead (a copy), leaving the original untouched. Pending batches stay pending against the original." },
        },
        required: ["path"],
      },
    },
  ];
  return [...base, ...extra];
}

export type HandlerOptions = {
  /** Directories the server may open files under. Default: process.cwd(). */
  allowRoots?: string[];
  /** Only the review tab may save; save_workbook returns the link. */
  reviewRequired?: boolean;
  /** The review server's base URL and per-process token, once it's up. */
  reviewUrl?: () => { url: string; token: string } | null;
};

/** The batches someone wrote: everything but the repeats of another workbook's row, column and sheet edits. */
const written = (s: WorkbookSession): ChangeBatch[] => s.batches.filter((b) => !isMirrorBatch(b));

export class ToolHandlers {
  private readonly workspace = new Workspace((abs) => this.allowed(abs));
  private readonly roots: string[];
  readonly reviewRequired: boolean;
  private readonly reviewBase: () => { url: string; token: string } | null;

  constructor(opts: HandlerOptions = {}) {
    const home = process.env.HOME ?? process.env.USERPROFILE ?? "";
    // Hosts don't reliably substitute home-directory placeholders (Claude
    // Desktop passed the bundle's `${HOME}/Documents` default literally, with
    // a leading slash). Accept every common spelling.
    const expand = (raw: string) => {
      let r = raw.trim().replace(/^\/+(?=(\$\{HOME\}|\$HOME|~))/, "");
      r = r.replace(/^\$\{HOME\}|^\$HOME|^%USERPROFILE%/, home);
      if (r === "~") return home;
      if (r.startsWith("~/")) return path.join(home, r.slice(2));
      return r;
    };
    // Default: the whole home directory. The review gate, not a folder list,
    // is what protects files; --allow is for people who want a tighter box.
    this.roots = (opts.allowRoots?.length ? opts.allowRoots : [home || process.cwd()]).map((r) => path.resolve(expand(r)));
    this.reviewRequired = !!opts.reviewRequired;
    this.reviewBase = opts.reviewUrl ?? (() => null);
  }

  private allowed(abs: string): boolean {
    return this.roots.some((r) => abs === r || abs.startsWith(r + path.sep));
  }

  private resolve(p: unknown): string {
    if (typeof p !== "string" || !p.trim()) throw new Error("`path` is required");
    const abs = path.resolve(p.startsWith("~/") ? path.join(process.env.HOME ?? "", p.slice(2)) : p);
    if (!this.allowed(abs)) {
      throw new Error(
        `refused: ${abs} is outside the allowed folders (${this.roots.join(", ")}). ` +
          "Do NOT copy or move the file to work around this — edits must land in the user's original file. " +
          "Ask the user to restart gridpath with `--allow <folder>` covering it.",
      );
    }
    return abs;
  }

  reviewUrlFor(workbookPath: string): string | null {
    const base = this.reviewBase();
    return base ? `${base.url}/review?path=${encodeURIComponent(workbookPath)}&token=${base.token}` : null;
  }

  /** Open (or reopen) a session; replays batches persisted on disk. */
  session(p: unknown): WorkbookSession {
    return this.workspace.session(this.resolve(p));
  }

  /** The batches of a workbook that someone wrote (what review lists and save counts). */
  pendingBatches(p: unknown): ChangeBatch[] {
    return written(this.session(p));
  }

  openWorkbooks(): Array<{ path: string; pending: number; review_url: string | null }> {
    return this.workspace.open().map((s) => ({ path: s.path, pending: written(s).length, review_url: this.reviewUrlFor(s.path) }));
  }

  async call(name: string, input: Record<string, unknown>): Promise<unknown> {
    try {
      return await this.dispatch(name, input);
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
  }

  /** Per-cell reject from the review tab: drop those cells' mutations from one batch, replay everything. */
  rejectCells(workbookPath: string, batchId: string, cells: Array<{ sheet: string; row: number; col: number }>): unknown {
    const s = this.session(workbookPath);
    const b = written(s).find((x) => x.id === batchId);
    if (!b) return { ok: false, error: `no batch ${batchId}` };
    const drop = new Set(cells.map((c) => `${c.sheet}!${c.row},${c.col}`));
    const keepMutations = b.mutations.flatMap((m): UniverMutation[] => {
      if (m.type === "set_cell") return drop.has(`${m.address.sheet}!${m.address.row},${m.address.col}`) ? [] : [m];
      if (m.type === "set_range") {
        // Blank out dropped cells (null preserves the existing cell in set_range semantics).
        const values = m.values.map((rowVals, dr) => rowVals.map((v, dc) => (drop.has(`${m.sheet}!${m.start_row + dr},${m.start_col + dc}`) ? null : v)));
        return values.some((r) => r.some((v) => v !== null && v !== undefined && v !== "")) ? [{ ...m, values }] : [];
      }
      if (m.type === "clear_range") {
        const kept = m.cells.filter((c) => !drop.has(`${m.sheet}!${c.row},${c.col}`));
        return kept.length ? [{ ...m, cells: kept }] : [];
      }
      if (m.type === "set_format") {
        const kept = m.cells.filter((c) => !drop.has(`${m.sheet}!${c.row},${c.col}`));
        return kept.length ? [{ ...m, cells: kept }] : [];
      }
      return [m];
    });
    const batches = s.batches
      .map((x) => (x.id === batchId ? { ...x, mutations: keepMutations } : x))
      .filter((x) => x.mutations.length > 0);
    const fresh = this.workspace.rebuild(s.path, batches);
    return { ok: true, batch_id: batchId, rejected_cells: cells.length, remaining_in_batch: keepMutations.length, batches: written(fresh).length };
  }

  /** Accept from the review tab: a status flip (save writes accepted + pending). */
  acceptBatch(workbookPath: string, batchId: string | null): unknown {
    const s = this.session(workbookPath);
    let n = 0;
    for (const b of s.batches) {
      if ((batchId === null || b.id === batchId) && b.status === "pending") {
        b.status = "accepted";
        n++;
      }
    }
    saveStored(s.path, s.batches);
    return { ok: true, accepted: n, pending: written(s).filter((b) => b.status === "pending").length };
  }

  /** Save from the review tab: bypasses --review-required (that's the point). */
  async saveFromReview(workbookPath: string, asPath?: string | null): Promise<unknown> {
    try {
      return this.doSave(this.session(workbookPath), asPath ?? undefined);
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
  }

  private doSave(s: WorkbookSession, asPath?: string): unknown {
    const count = written(s).length;
    // Rows, columns or sheets that moved in a workbook this one reads moved
    // its links too, and its file has to be told, even with nothing written here.
    const followed = s.batches.length - count;
    if (count === 0 && followed === 0) return { ok: true, saved: false, note: "nothing pending" };
    const target = asPath ? this.resolve(/\.xls[xm]$/i.test(asPath) ? asPath : `${asPath}.xlsx`) : s.path;
    // The workbooks this one reads go into the file as links, the way Excel writes them.
    const plan = this.workspace.savePlan(s.path, target);
    if (!plan.ok) return { ok: false, error: "save_source_first", reason: plan.reason, save_first: plan.save_first };
    const built = s.buildSave(plan.links);
    if (!built.ok) return { ok: false, error: built.links ? "links_not_saved" : "surgical_save_unavailable", reason: built.reason };
    const reads = plan.reads;
    const linked = reads.length
      ? [`Its links to ${reads.map((p) => path.basename(p)).join(", ")} are in the file: Excel opens it with the figures in place, and they are live again when it is opened here.`, ...plan.notes]
      : [];
    if (asPath) {
      // Save As: a patched copy next to (or wherever) the user asked; the
      // original and its pending batches are left exactly as they were.
      fs.mkdirSync(path.dirname(target), { recursive: true });
      fs.writeFileSync(target, built.bytes);
      return {
        ok: true,
        saved: true,
        copy: true,
        path: target,
        bytes: built.bytes.length,
        batches: count,
        ...(reads.length ? { reads } : {}),
        note: [`Saved a copy to ${target}. The original is untouched and its ${count} batch(es) are still pending.`, ...linked].join(" "),
      };
    }
    // Workbooks that read this one and are not open take up its moved rows first.
    this.workspace.beforeSave(s.path);
    // The ones whose files already point into this one now point at the old rows.
    const also = this.workspace.movedWith(s.path);
    fs.writeFileSync(s.path, built.bytes);
    clearStored(s.path);
    // The batches are now in the file. Drop the session; the next call reopens
    // it from disk for a fresh baseline (a reopen here would make every save
    // pay for a full load even when the agent is done).
    // Workbooks that read this one keep the values they have; they follow it
    // again from its next edit.
    this.workspace.close(s.path);
    const notes = [
      ...linked,
      ...(also.length
        ? [`Rows, columns or sheets moved in this workbook, and ${also.map((p) => path.basename(p)).join(", ")} link(s) to it from the file: save ${also.length > 1 ? "those" : "that one"} too (also_save), so the links in the file point at the new places.`]
        : []),
    ];
    return {
      ok: true,
      saved: true,
      path: s.path,
      bytes: built.bytes.length,
      batches: count,
      ...(reads.length ? { reads } : {}),
      ...(also.length ? { also_save: also } : {}),
      ...(notes.length ? { note: notes.join(" ") } : {}),
    };
  }

  private async dispatch(name: string, input: Record<string, unknown>): Promise<unknown> {
    switch (name) {
      case "list_workbooks":
        return this.listWorkbooks(input.path ?? input.folder);
      case "create_workbook":
        return this.createWorkbook(input.path, input.sheets);
      case "list_batches": {
        const s = this.session(input.path);
        return {
          ok: true,
          batches: written(s).map((b) => ({
            id: b.id,
            tool: b.prompt,
            status: b.status,
            mutations: b.mutations.length,
            created_at: b.created_at,
            ...(b.links?.length ? { reads: [...new Set(b.links.map((l) => l.source))] } : {}),
          })),
          review_url: this.reviewUrlFor(s.path),
        };
      }
      case "reject_batch": {
        const s = this.session(input.path);
        const id = String(input.batch_id ?? "");
        const b = s.batches.find((x) => x.id === id);
        if (!b) return { ok: false, error: `no batch ${id}` };
        if (b.link_origin) {
          return {
            ok: false,
            error: `batch ${id} is not an edit of this workbook: it follows batch ${b.link_origin.batch_id} of ${b.link_origin.path}, which this workbook reads. Reject that one (path = ${b.link_origin.path}, batch_id = ${b.link_origin.batch_id}).`,
          };
        }
        const readers = this.workspace.followersOf(s.path);
        const fresh = this.workspace.rebuild(s.path, s.batches.filter((x) => x.id !== id));
        return {
          ok: true,
          rejected: id,
          remaining: written(fresh).length,
          ...(readers.length ? { recalculated_workbooks: readers } : {}),
          review_url: this.reviewUrlFor(s.path),
        };
      }
      case "save_workbook": {
        const s = this.session(input.path);
        const asPath = typeof input.as === "string" && input.as.trim() ? input.as : undefined;
        const pending = written(s).length;
        if (this.reviewRequired && pending > 0 && !asPath) {
          return {
            ok: true,
            saved: false,
            pending_batches: pending,
            review_url: this.reviewUrlFor(s.path),
            note: "This server requires human review before saving. Give the user the review_url; they accept and save from there. Nothing has been written to the file.",
          };
        }
        return this.doSave(s, asPath);
      }
    }

    const s = this.session(input.path);
    const result = interpretToolCall(name, input);
    switch (result.kind) {
      case "describe_workbook": {
        const elsewhere = this.elsewhere(s, result.sheet);
        if (elsewhere) return elsewhere;
        const reads = this.workspace.linksOf(s.path);
        const readBy = this.workspace.readersOf(s.path);
        return {
          ...describeWorkbookPayload(s.workbookIndex(), result.sheet),
          ...(reads.length ? { linked_workbooks: reads } : {}),
          ...(readBy.length ? { read_by: readBy } : {}),
        };
      }
      case "query_rows": {
        const elsewhere = this.elsewhere(s, result.sheet);
        if (elsewhere) return elsewhere;
        try {
          return s.queryRows(result.sheet, result.query);
        } catch (e: any) {
          return { ok: false, error: String(e?.message ?? e) };
        }
      }
      case "find_rows": {
        const elsewhere = this.elsewhere(s, result.sheet);
        if (elsewhere) return elsewhere;
        const index = s.workbookIndex();
        const { matches, total } = findRowsInIndex(index, result.query, result.sheet, result.max_results);
        return {
          matches: matches.map((m: any) => {
            const out: any = { sheet: m.sheet, row: m.row, label: m.label };
            if (m.section !== null) out.section = m.section;
            if (m.column !== undefined) out.column = m.column;
            if (m.sampleCol !== null) {
              const c = s.cell(m.sheet, m.row - 1, m.sampleCol);
              if (c.value !== null || c.formula) {
                out.sample = c.formula ? `${c.cell} = ${c.formula} → ${c.value ?? ""}` : `${c.cell} = ${c.value}`;
              }
            }
            return out;
          }),
          total,
          truncated: total > matches.length,
        };
      }
      case "read": {
        const elsewhere = this.elsewhere(s, result.sheet);
        if (elsewhere) return elsewhere;
        const { cells, truncated } = s.readRange(result.sheet, result.range);
        return { sheet: result.sheet, range: result.range, cells, truncated };
      }
      case "mutations":
        return this.applyAndReadBack(s, name, result.mutations);
      case "copy": {
        const elsewhere = this.elsewhere(s, result.sheet);
        if (elsewhere) return elsewhere;
        const r = s.reader();
        const expansion = expandCopy(
          { getCell: (sh, row, col) => r.getCell(sh, row, col), getCellFormat: (sh, row, col) => r.getCellFormat(sh, row, col) ?? {} },
          result,
        );
        if (!expansion.ok) return expansion;
        const back = this.applyAndReadBack(s, name, expansion.mutations) as Record<string, unknown>;
        return { ...back, dest_range: expansion.dest_range, copied_cells: expansion.copied_cells };
      }
      case "script": {
        const run = await runScript(s.snapshot(), result.code);
        if (!run.ok) return { ok: false, error: run.error, script_logs: run.logs };
        const back = this.applyAndReadBack(s, name, run.mutations) as Record<string, unknown>;
        return { ...back, writes: run.writes, script_logs: run.logs };
      }
      case "ignored":
        return { ok: false, error: result.reason };
      default:
        return { ok: false, error: `${name} is not available in the headless server` };
    }
  }

  /**
   * A sheet argument that names another workbook's sheet (`[NOW.xlsx]Model`,
   * as links spell it): say where that sheet really is.
   */
  private elsewhere(s: WorkbookSession, sheet: unknown): { ok: false; error: string } | null {
    if (typeof sheet !== "string" || !isLinkSheetName(sheet)) return null;
    const from = this.workspace.sourceOf(s.path, sheet);
    return {
      ok: false,
      error:
        `${sheet} is not a sheet of ${path.basename(s.path)}; it is how formulas here refer to a sheet of another workbook. ` +
        (from ? `Read it there: path = ${from.workbook}, sheet = ${from.sheet}.` : "Read it in its own file."),
    };
  }

  private listWorkbooks(folder: unknown): unknown {
    const dir = this.resolve(folder);
    let names: string[];
    try {
      names = fs.readdirSync(dir);
    } catch {
      return { ok: false, error: `cannot read folder ${dir}` };
    }
    // `~$Book.xlsx` is the lock file Excel keeps next to an open workbook.
    const files = names.filter((n) => /\.xls[xm]$/i.test(n) && !n.startsWith("~$") && !n.startsWith(".")).sort((a, b) => a.localeCompare(b));
    const LIMIT = 200;
    const workbooks = files.slice(0, LIMIT).flatMap((n) => {
      const abs = path.join(dir, n);
      let st: fs.Stats;
      try {
        st = fs.statSync(abs);
      } catch {
        return [];
      }
      if (!st.isFile()) return [];
      const out: Record<string, unknown> = { path: abs, name: n, size_kb: Math.max(1, Math.round(st.size / 1024)), modified: st.mtime.toISOString() };
      if (this.workspace.isOpen(abs)) {
        out.open = true;
        out.pending_batches = written(this.workspace.session(abs)).length;
        const reads = this.workspace.linksOf(abs);
        if (reads.length) out.reads = reads;
        const readBy = this.workspace.readersOf(abs);
        if (readBy.length) out.read_by = readBy;
      }
      return [out];
    });
    return { ok: true, folder: dir, workbooks, ...(files.length > LIMIT ? { truncated: true, total: files.length } : {}) };
  }

  private createWorkbook(p: unknown, sheetsArg: unknown): unknown {
    if (typeof p !== "string" || !p.trim()) throw new Error("`path` is required");
    if (/\.(xlsm|xls|csv)$/i.test(p)) return { ok: false, error: "only .xlsx workbooks can be created" };
    const abs = this.resolve(/\.xlsx$/i.test(p) ? p : `${p}.xlsx`);
    const sheets = Array.isArray(sheetsArg) && sheetsArg.length ? sheetsArg.map((n) => String(n).trim()) : ["Sheet1"];
    const bad = sheets.find((n) => !n || n.length > 31 || /[[\]:*?/\\]/.test(n) || n.startsWith("'") || n.endsWith("'"));
    if (bad !== undefined) return { ok: false, error: `not a valid sheet name: "${bad}" (1 to 31 characters, none of [ ] : * ? / \\)` };
    if (new Set(sheets.map((n) => n.toUpperCase())).size !== sheets.length) return { ok: false, error: "sheet names must be different from each other" };
    const blank = loadEngine().blankXlsx;
    if (typeof blank !== "function") return { ok: false, error: "this engine build cannot create workbooks (run `npm run build:engine`)" };
    const bytes = blank(sheets);
    fs.mkdirSync(path.dirname(abs), { recursive: true });
    try {
      // `wx`: fail rather than replace a file that is already there.
      fs.writeFileSync(abs, bytes, { flag: "wx" });
    } catch (e: any) {
      if (e?.code === "EEXIST") return { ok: false, error: `${abs} already exists. Nothing was changed; open it with describe_workbook, or choose another name.` };
      throw e;
    }
    return { ok: true, created: true, path: abs, sheets, note: "An empty workbook on disk. Write into it like any other workbook." };
  }

  private applyAndReadBack(s: WorkbookSession, tool: string, mutations: UniverMutation[]): unknown {
    if (mutations.length === 0) return { ok: false, error: "no cells to write — check the arguments" };
    let applied: ReturnType<Workspace["apply"]>;
    try {
      applied = this.workspace.apply(s, tool, mutations);
    } catch (e) {
      if (e instanceof LinkError) return { ok: false, error: e.message, links: e.links.filter(linkFailed) };
      throw e;
    }
    const { batch, links, updated } = applied;

    // Read back what was written (same shape the desktop app returns).
    const touched = new Map<string, { sheet: string; row: number; col: number }>();
    const expected = new Set<string>();
    for (const m of batch.mutations) {
      if (m.type === "set_cell") {
        const k = `${m.address.sheet}!${m.address.row},${m.address.col}`;
        touched.set(k, m.address);
        if (m.new_value != null || m.new_formula != null) expected.add(k);
      } else if (m.type === "set_range") {
        m.values.forEach((rowVals, dr) =>
          rowVals.forEach((v, dc) => {
            if (v === null || v === undefined || v === "") return;
            const k = `${m.sheet}!${m.start_row + dr},${m.start_col + dc}`;
            touched.set(k, { sheet: m.sheet, row: m.start_row + dr, col: m.start_col + dc });
            expected.add(k);
          }),
        );
      }
    }
    const cells: Array<ReadCell & { sheet: string }> = [];
    let landed = 0;
    for (const [k, a] of touched) {
      const c = s.cell(a.sheet, a.row, a.col);
      cells.push({ sheet: a.sheet, ...c });
      if (expected.has(k) && (c.value !== null || c.formula)) landed++;
    }
    const rows = new Map<string, { sheet: string; row: number }>();
    for (const a of touched.values()) rows.set(`${a.sheet}!${a.row}`, a);
    const row_map = [...rows.values()]
      .map(({ sheet, row }) => {
        const label = firstLabelIn(Array.from({ length: 4 }, (_, col) => s.cell(sheet, row, col).value));
        return label ? { sheet, row: row + 1, label } : null;
      })
      .filter((x): x is { sheet: string; row: number; label: string } => x !== null)
      .sort((a, b) => a.row - b.row);
    const write_check = expected.size === 0 ? "ok" : landed === 0 ? "FAILED" : landed < expected.size ? "partial" : "ok";
    const review_url = this.reviewUrlFor(s.path);
    return {
      ok: write_check !== "FAILED",
      batch_id: batch.id,
      write_check,
      cells: cells.slice(0, 200),
      row_map,
      pending_batches: written(s).length,
      ...(links.length ? { links } : {}),
      ...(updated.length ? { recalculated_workbooks: updated } : {}),
      review_url,
      note: this.reviewRequired
        ? "Changes are pending in memory. The user reviews and saves them at review_url — give them that link."
        : "Changes are pending in memory. Call save_workbook to write them to the file, or give the user review_url to review first.",
    };
  }

  /** Test/diagnostic helper. */
  static a1 = a1Of;
}
