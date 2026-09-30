/**
 * Tool handlers over headless workbook sessions. Every tool takes `path`;
 * a session is opened on first use and kept for the process lifetime.
 *
 * Nothing touches the file until `save_workbook`. Writes accumulate as
 * pending batches (the same ChangeBatch the desktop app reviews), mirrored
 * to disk next to the workbook so the review link outlives this process;
 * `reject_batch` drops one and replays the rest onto a fresh model.
 */
import fs from "node:fs";
import path from "node:path";
import { interpretToolCall } from "../core/toolToMutation";
import { buildWorkbookIndex, describeWorkbookPayload, findRowsInIndex, firstLabelIn } from "../core/workbookIndex";
import type { ChangeBatch, UniverMutation } from "../core/types";
import { WorkbookSession, a1Of, type ReadCell } from "../workbook/session";
import { clearStored, loadStored, saveStored } from "../workbook/store";
import { runScript } from "./script";
import toolsJson from "./tools.json";

export type ToolSchema = { name: string; description: string; input_schema: any };

/** Tools the headless server offers (subset of the desktop agent's). */
const EXCLUDED = new Set(["done", "read_reference", "fetch_web", "edgar_lookup", "keep_pages", "read_source", "copy_range", "stage_data"]);

const PATH_PROP = {
  type: "string",
  description: "Absolute path of the .xlsx to operate on (or a path relative to the server's working directory).",
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
      input_schema: { type: "object", properties: { path: PATH_PROP }, required: ["path"] },
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

export class ToolHandlers {
  private readonly sessions = new Map<string, WorkbookSession>();
  private readonly roots: string[];
  private batchSeq = 0;
  readonly reviewRequired: boolean;
  private readonly reviewBase: () => { url: string; token: string } | null;

  constructor(opts: HandlerOptions = {}) {
    this.roots = (opts.allowRoots?.length ? opts.allowRoots : [process.cwd()]).map((r) => path.resolve(r));
    this.reviewRequired = !!opts.reviewRequired;
    this.reviewBase = opts.reviewUrl ?? (() => null);
  }

  private resolve(p: unknown): string {
    if (typeof p !== "string" || !p.trim()) throw new Error("`path` is required");
    const abs = path.resolve(p.startsWith("~/") ? path.join(process.env.HOME ?? "", p.slice(2)) : p);
    if (!this.roots.some((r) => abs === r || abs.startsWith(r + path.sep))) {
      throw new Error(`refused: ${abs} is outside the allowed roots (${this.roots.join(", ")})`);
    }
    return abs;
  }

  reviewUrlFor(workbookPath: string): string | null {
    const base = this.reviewBase();
    return base ? `${base.url}/review?path=${encodeURIComponent(workbookPath)}&token=${base.token}` : null;
  }

  /** Open (or reopen) a session; replays batches persisted on disk. */
  session(p: unknown): WorkbookSession {
    const abs = this.resolve(p);
    let s = this.sessions.get(abs);
    if (!s) {
      s = WorkbookSession.open(abs, fs.readFileSync(abs));
      const stored = loadStored(abs);
      if (stored) {
        for (const b of stored.batches) {
          s.apply({ id: b.id, prompt: b.prompt, justification: b.justification, mutations: b.mutations, created_at: b.created_at });
          const n = Number(b.id.replace(/^b/, ""));
          if (Number.isFinite(n) && n > this.batchSeq) this.batchSeq = n;
        }
      }
      this.sessions.set(abs, s);
    }
    return s;
  }

  openWorkbooks(): Array<{ path: string; pending: number; review_url: string | null }> {
    return [...this.sessions.values()].map((s) => ({ path: s.path, pending: s.batches.length, review_url: this.reviewUrlFor(s.path) }));
  }

  async call(name: string, input: Record<string, unknown>): Promise<unknown> {
    try {
      return await this.dispatch(name, input);
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
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
    return { ok: true, accepted: n, pending: s.batches.filter((b) => b.status === "pending").length };
  }

  /** Save from the review tab: bypasses --review-required (that's the point). */
  async saveFromReview(workbookPath: string): Promise<unknown> {
    try {
      return this.doSave(this.session(workbookPath));
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
  }

  private doSave(s: WorkbookSession): unknown {
    if (s.batches.length === 0) return { ok: true, saved: false, note: "nothing pending" };
    const built = s.buildSave();
    if (!built.ok) return { ok: false, error: "surgical_save_unavailable", reason: built.reason };
    fs.writeFileSync(s.path, built.bytes);
    clearStored(s.path);
    // Fresh baseline from the saved bytes; the batches are now in the file.
    this.sessions.set(s.path, WorkbookSession.open(s.path, built.bytes));
    return { ok: true, saved: true, path: s.path, bytes: built.bytes.length, batches: s.batches.length };
  }

  private async dispatch(name: string, input: Record<string, unknown>): Promise<unknown> {
    switch (name) {
      case "list_batches": {
        const s = this.session(input.path);
        return {
          ok: true,
          batches: s.batches.map((b) => ({ id: b.id, tool: b.prompt, status: b.status, mutations: b.mutations.length, created_at: b.created_at })),
          review_url: this.reviewUrlFor(s.path),
        };
      }
      case "reject_batch": {
        const s = this.session(input.path);
        const id = String(input.batch_id ?? "");
        if (!s.batches.some((b) => b.id === id)) return { ok: false, error: `no batch ${id}` };
        const keep = s.batches.filter((b) => b.id !== id);
        const fresh = WorkbookSession.open(s.path, s.originalBytes);
        for (const b of keep) fresh.apply({ id: b.id, prompt: b.prompt, justification: b.justification, mutations: b.mutations, created_at: b.created_at });
        this.sessions.set(s.path, fresh);
        saveStored(s.path, fresh.batches);
        return { ok: true, rejected: id, remaining: keep.length, review_url: this.reviewUrlFor(s.path) };
      }
      case "save_workbook": {
        const s = this.session(input.path);
        if (this.reviewRequired && s.batches.length > 0) {
          return {
            ok: true,
            saved: false,
            pending_batches: s.batches.length,
            review_url: this.reviewUrlFor(s.path),
            note: "This server requires human review before saving. Give the user the review_url; they accept and save from there. Nothing has been written to the file.",
          };
        }
        return this.doSave(s);
      }
    }

    const s = this.session(input.path);
    const result = interpretToolCall(name, input);
    switch (result.kind) {
      case "describe_workbook":
        return describeWorkbookPayload(buildWorkbookIndex(s.snapshot()), result.sheet);
      case "find_rows": {
        const index = buildWorkbookIndex(s.snapshot());
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
        const { cells, truncated } = s.readRange(result.sheet, result.range);
        return { sheet: result.sheet, range: result.range, cells, truncated };
      }
      case "mutations":
        return this.applyAndReadBack(s, name, result.mutations);
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

  private applyAndReadBack(s: WorkbookSession, tool: string, mutations: UniverMutation[]): unknown {
    if (mutations.length === 0) return { ok: false, error: "no cells to write — check the arguments" };
    const id = `b${++this.batchSeq}`;
    const batch: ChangeBatch = s.apply({ id, prompt: tool, justification: tool, mutations });
    saveStored(s.path, s.batches);

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
      pending_batches: s.batches.length,
      review_url,
      note: this.reviewRequired
        ? "Changes are pending in memory. The user reviews and saves them at review_url — give them that link."
        : "Changes are pending in memory. Call save_workbook to write them to the file, or give the user review_url to review first.",
    };
  }

  /** Test/diagnostic helper. */
  static a1 = a1Of;
}
