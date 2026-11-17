/**
 * The agent's reads of ANOTHER workbook of the workspace.
 *
 * The agent works in one tab, and its write tools change that tab's
 * workbook only. The read tools (`describe_workbook`, `find_rows`,
 * `read_range`, `query_rows`) also take `workbook`, the file name of
 * another workbook in the workspace. Such a read never touches a grid: the
 * Tauri side holds a model for the workbook (its tab's, or one it loads
 * without a tab) and the same engine commands the active tab uses are
 * asked under that model's key.
 *
 * What the agent finds this way it brings into the active workbook as a
 * link: `='[SNOW.xlsx]Model'!BV32`. Every result here carries the prefix
 * for that, so the agent does not have to get the quoting right itself.
 */
import { invoke } from "@tauri-apps/api/core";
import { describeWorkbookPayload, findRowsInIndex, type WorkbookIndex } from "./workbookIndex";
import { interpretToolCall } from "./toolToMutation";
import { colLetters } from "./reviewGroups";
import { stripImplicitIntersection } from "../../../../packages/gridpath/src/workbook/snapshot";
import { workbookName, type WorkspaceMember } from "../workspaceApi";

/** The read tools that take `workbook`. */
export const WORKSPACE_READ_TOOLS: ReadonlySet<string> = new Set(["describe_workbook", "find_rows", "read_range", "query_rows"]);

/** A workbook of the workspace as the agent names it. */
export interface WorkspaceBook {
  /** The file name: what the agent passes as `workbook`, and what a link names. */
  label: string;
  path: string;
}

/**
 * The workbooks the agent can read besides the one at `activePath`: every
 * model of the workspace shown whose file is there, plus one that is only
 * loaded because an open workbook reads it. A link names a file by its
 * name alone, so of two files with the same name only the first is offered.
 */
export function workspaceBooks(members: WorkspaceMember[], activePath: string | null): WorkspaceBook[] {
  const own = activePath ? workbookName(activePath).toLowerCase() : null;
  const seen = new Set<string>();
  const out: WorkspaceBook[] = [];
  for (const m of members) {
    if (m.state === "missing" || m.state === "failed") continue;
    if (activePath && m.path === activePath) continue;
    const key = m.name.toLowerCase();
    if (key === own || seen.has(key)) continue;
    seen.add(key);
    out.push({ label: m.name, path: m.path });
  }
  return out;
}

/** `'[SNOW.xlsx]Model'!`: what goes in front of a cell address to link to it. */
export function linkPrefix(label: string, sheet: string): string {
  return `'[${label.replace(/'/g, "''")}]${sheet.replace(/'/g, "''")}'!`;
}

interface Model {
  key: string;
  sheets: Array<{ name: string; link: boolean }>;
}

interface WindowCell {
  r: number;
  c: number;
  v?: unknown;
  d?: string;
  f?: string;
  fileSaved?: boolean;
}

const json = (raw: ArrayBuffer) => JSON.parse(new TextDecoder().decode(new Uint8Array(raw)));

/** The bounds of an A1 range ("B5:D20", "C7"), 1-based; null when it is not one. */
export function rangeBounds(range: string): { r0: number; c0: number; r1: number; c1: number } | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?$/.exec(range.trim());
  if (!m) return null;
  const col = (letters: string) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const ca = col(m[1]), ra = Number(m[2]);
  const cb = m[3] ? col(m[3]) : ca, rb = m[4] ? Number(m[4]) : ra;
  if (ra < 1 || rb < 1) return null;
  return { r0: Math.min(ra, rb), c0: Math.min(ca, cb), r1: Math.max(ra, rb), c1: Math.max(ca, cb) };
}

const MAX_READ_CELLS = 500;

/**
 * Answers one read tool call that names another workbook. `input` is the
 * tool's input as the agent sent it (with `workbook`); `books` is what
 * `workspaceBooks` gives for the active tab. Returns the tool result as a
 * JSON string; a failure is a result too (the agent reads why), never a
 * throw.
 */
export async function readWorkspaceWorkbook(name: string, input: any, books: WorkspaceBook[]): Promise<string> {
  const wanted = String(input?.workbook ?? "").trim();
  const book = books.find((b) => b.label.toLowerCase() === wanted.toLowerCase());
  if (!book) {
    return JSON.stringify({
      ok: false,
      error: "unknown_workbook",
      workbook: wanted,
      available: books.map((b) => b.label),
      message:
        books.length > 0
          ? "No workbook of that name is in the workspace. Use a name from `available`, exactly; leave `workbook` out for the active workbook."
          : "There are no other workbooks in the workspace. Leave `workbook` out to read the active workbook.",
    });
  }
  const { workbook: _named, ...rest } = input ?? {};
  const parsed = interpretToolCall(name, rest);
  if (parsed.kind === "ignored") return JSON.stringify({ ok: false, error: "bad_arguments", message: parsed.reason });

  const open = () => invoke<Model>("workspace_model", { path: book.path });
  const run = async (model: Model): Promise<object> => {
    const sheets = model.sheets.map((s, index) => ({ ...s, index })).filter((s) => !s.link);
    const sheetNamed = (sheetName: string) => sheets.find((s) => s.name === sheetName) ?? null;
    const unknownSheet = (sheetName: string) => ({
      ok: false,
      error: "unknown_sheet",
      workbook: book.label,
      sheet: sheetName,
      available: sheets.map((s) => s.name),
      message: `${book.label} has no sheet of that name. Use a name from \`available\`, exactly.`,
    });
    const window = async (sheet: number, r0: number, r1: number, c0: number, c1: number): Promise<WindowCell[]> => {
      const win = json(await invoke<ArrayBuffer>("calc_engine_sheet_window", { tab: model.key, sheet, r0, r1, c0, c1 }));
      return Array.isArray(win?.cells) ? win.cells : [];
    };
    const formulaOf = (c: WindowCell) => (typeof c.f === "string" && c.f ? stripImplicitIntersection(c.f) : null);

    if (parsed.kind === "read") {
      const sheet = sheetNamed(parsed.sheet);
      if (!sheet) return unknownSheet(parsed.sheet);
      const b = rangeBounds(parsed.range);
      if (!b) return { ok: false, error: "bad_range", message: `"${parsed.range}" is not an A1 range such as "B5:N40" or "B7".` };
      // At most 500 cells, row by row from the top-left, as the active
      // workbook's read_range gives them.
      const cols = Math.min(b.c1 - b.c0 + 1, MAX_READ_CELLS);
      const rows = Math.min(b.r1 - b.r0 + 1, Math.max(1, Math.floor(MAX_READ_CELLS / cols)));
      const truncated = cols * rows < (b.c1 - b.c0 + 1) * (b.r1 - b.r0 + 1);
      const cells = (await window(sheet.index, b.r0, b.r0 + rows - 1, b.c0, b.c0 + cols - 1))
        .sort((x, y) => x.r - y.r || x.c - y.c)
        .map((c) => {
          const out: { cell: string; value: unknown; display?: string; formula: string | null; file_saved?: unknown } = {
            cell: `${colLetters(c.c - 1)}${c.r}`,
            value: c.v ?? null,
            formula: formulaOf(c),
          };
          if (c.d !== undefined && c.d !== String(c.v)) out.display = c.d;
          if (c.fileSaved) out.file_saved = out.value;
          return out;
        })
        .filter((c) => c.value !== null || c.formula !== null);
      const prefix = linkPrefix(book.label, sheet.name);
      return {
        workbook: book.label,
        sheet: sheet.name,
        range: parsed.range,
        cells,
        truncated,
        link_prefix: prefix,
        link_note:
          `To bring one of these cells into the active workbook, write a formula that links to it: =${prefix}${cells[0]?.cell ?? "B2"} ` +
          `(a range: =SUM(${prefix}B5:D5)). Do not type the number.`,
      };
    }

    if (parsed.kind === "query_rows") {
      const sheet = sheetNamed(parsed.sheet);
      if (!sheet) return unknownSheet(parsed.sheet);
      try {
        const result = await invoke<any>("calc_engine_query_rows", { tab: model.key, query: { ...parsed.query, sheet: sheet.index } });
        return { workbook: book.label, link_prefix: linkPrefix(book.label, sheet.name), ...result };
      } catch (e) {
        if (String(e).includes("no workbook loaded")) throw e;
        return { ok: false, error: "bad_query", message: String(e) };
      }
    }

    if (parsed.kind === "describe_workbook" || parsed.kind === "find_rows") {
      if (parsed.sheet && !sheetNamed(parsed.sheet)) return unknownSheet(parsed.sheet);
      const index = json(await invoke<ArrayBuffer>("calc_engine_workbook_index", { tab: model.key })) as WorkbookIndex;
      if (parsed.kind === "describe_workbook") {
        return {
          workbook: book.label,
          link_form: `'[${book.label}]<sheet name>'!<cell>, e.g. =${linkPrefix(book.label, sheets[0]?.name ?? "Sheet1")}B2`,
          ...describeWorkbookPayload(index, parsed.sheet),
        };
      }
      const { matches, total } = findRowsInIndex(index, parsed.query, parsed.sheet, parsed.max_results);
      const found = await Promise.all(
        matches.map(async (m) => {
          const out: any = { sheet: m.sheet, row: m.row, label: m.label };
          if (m.section !== null) out.section = m.section;
          if (m.column !== undefined) out.column = m.column;
          const sheet = sheetNamed(m.sheet);
          if (sheet && m.sampleCol !== null) {
            const c = (await window(sheet.index, m.row, m.row, m.sampleCol + 1, m.sampleCol + 1))[0];
            if (c && ((c.v !== null && c.v !== undefined) || c.f)) {
              const addr = `${colLetters(m.sampleCol)}${m.row}`;
              const formula = formulaOf(c);
              out.sample = formula ? `${addr} = ${formula} → ${c.v ?? ""}` : `${addr} = ${c.v}`;
              out.sample_link = `=${linkPrefix(book.label, m.sheet)}${addr}`;
            }
          }
          return out;
        }),
      );
      return {
        workbook: book.label,
        matches: found,
        total,
        truncated: total > matches.length,
        link_note: "`sample_link` is the formula that links to the sample cell from the active workbook; change the column for another period.",
      };
    }

    return { ok: false, error: "bad_arguments", message: `${name} cannot read another workbook.` };
  };

  try {
    const model = await open();
    try {
      return JSON.stringify(await run(model));
    } catch (e) {
      // The model was let go between the two calls (a tab closed): once more.
      if (!String(e).includes("no workbook loaded")) throw e;
      return JSON.stringify(await run(await open()));
    }
  } catch (e) {
    return JSON.stringify({
      ok: false,
      error: "workbook_unavailable",
      workbook: book.label,
      message: `${book.label} could not be read: ${String(e)}`,
    });
  }
}
