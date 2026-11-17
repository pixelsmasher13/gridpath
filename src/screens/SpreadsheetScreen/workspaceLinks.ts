/**
 * Saving a workbook that reads other workbooks.
 *
 * Live, a link is spelled with the file's name: `'[NOW.xlsx]Model'!B23`. An
 * .xlsx spells it with a number, `[1]Model!B23`, and keeps one part per
 * workbook read: where that file is, the sheets read, and the last value of
 * every cell read (what Excel shows before it has opened the other file).
 *
 * The save patch is built from the cells this session wrote, in the live
 * spelling. The engine, which knows which workbook each link reads, gives
 * the file's side of it (`LinkSavePlan`), and this folds the two together.
 */
import { addXlfnPrefixes } from "./xlfnCompat";
import { stripImplicitIntersection } from "../../../packages/gridpath/src/workbook/snapshot";
import type { LinkSavePlan } from "./ironcalcEngine";

export interface LinkMergeResult {
  /** Formula cells of the patch respelled for the file. */
  respelled: number;
  /** Formula cells the session did not write, added because their links
   * were rewritten to follow another workbook. */
  added: number;
  /** Workbooks read, as link parts of the file. */
  parts: number;
}

/**
 * Folds `plan` into `patch` (the object `buildWorkbookPatch` returned), in
 * place.
 *
 * - A formula the session wrote that names another workbook gets the
 *   file's spelling and its current value.
 * - With `rewriteAll` (rows moved or a sheet was renamed in a workbook this
 *   one reads) every such formula is written, edited here or not.
 * - A cell the session cleared or overwrote with a plain value is left as
 *   the session has it.
 */
export function mergeLinkPlan(patch: any, plan: LinkSavePlan | null | undefined): LinkMergeResult {
  const out: LinkMergeResult = { respelled: 0, added: 0, parts: 0 };
  if (!patch || !plan) return out;

  if (plan.cells.length > 0) {
    const sheets: any[] = Array.isArray(patch.sheets) ? patch.sheets : [];
    const byName = new Map<string, any>(sheets.map((sp) => [sp.name, sp]));
    const indexed = new Map<string, Map<string, any>>();
    const cellsOf = (name: string): Map<string, any> | null => {
      let byKey = indexed.get(name);
      if (byKey) return byKey;
      const sp = byName.get(name);
      if (!sp) return null;
      byKey = new Map((sp.cells ?? []).map((cell: any) => [`${cell.r},${cell.c}`, cell]));
      indexed.set(name, byKey!);
      return byKey!;
    };
    for (const cell of plan.cells) {
      const key = `${cell.r},${cell.c}`;
      const written = cellsOf(cell.sheet)?.get(key);
      if (written) {
        if (typeof written.f === "string") {
          // The session's own text decides whether `@` is the user's.
          written.f = addXlfnPrefixes(written.f.includes("@") ? cell.f : stripImplicitIntersection(cell.f));
          if (cell.v !== undefined) written.v = cell.v;
          out.respelled++;
        } else if (plan.rewriteAll && written.cache !== undefined && !written.clear && written.v === undefined) {
          // Only its cached result was going to be refreshed.
          written.f = addXlfnPrefixes(stripImplicitIntersection(cell.f));
          written.v = cell.v ?? written.cache;
          delete written.cache;
          out.added++;
        }
        continue;
      }
      if (!plan.rewriteAll) continue;
      let sp = byName.get(cell.sheet);
      if (!sp) {
        sp = { name: cell.sheet };
        sheets.push(sp);
        byName.set(cell.sheet, sp);
      }
      const add: any = { r: cell.r, c: cell.c, f: addXlfnPrefixes(stripImplicitIntersection(cell.f)) };
      if (cell.v !== undefined) add.v = cell.v;
      (sp.cells ??= []).push(add);
      cellsOf(cell.sheet)?.set(key, add);
      out.added++;
    }
    if (sheets.length > 0) patch.sheets = sheets;
  }

  if (plan.externalLinks.length > 0) {
    patch.externalLinks = plan.externalLinks;
    out.parts = plan.externalLinks.length;
  }
  return out;
}
