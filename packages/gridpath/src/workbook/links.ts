/**
 * Links to other workbooks, as they are written in formula text:
 *
 *   [NOW.xlsx]Model!B23
 *   '[NOW.xlsx]Income Statement'!B23
 *   '/Users/me/Models/[NOW.xlsx]Model'!B23
 *   [1]Model!B23                              (inside a saved .xlsx)
 *
 * Text handling only. Which file a book name means, and keeping the values
 * current, is the workspace's job (`workspace.ts`).
 *
 * Inside a workbook that reads another one, the engine keeps a hidden copy
 * of each sheet it reads, named exactly like the link (`[NOW.xlsx]Model`).
 * Those names are never shown as sheets of the workbook.
 */

/** True for the name of a hidden copy of another workbook's sheet. */
export const isLinkSheetName = (name: string): boolean => name.startsWith("[");

export type FoundLink = {
  /** Offset of the prefix in the formula, and one past its `!`. */
  start: number;
  end: number;
  /** Folder written before the bracket (quoted form only), else "". */
  dir: string;
  /** Text between the brackets: a file name, or a number in a saved file. */
  book: string;
  sheet: string;
};

const NAME_CHAR = /[\p{L}\p{N}_.]/u;

/**
 * Every link prefix in a formula, in order. String literals are skipped, and
 * so are the brackets of table references (`Sales[Amount]`, `[@Amount]`):
 * only `[book]Sheet!` counts.
 */
export function findLinks(formula: string): FoundLink[] {
  const out: FoundLink[] = [];
  const n = formula.length;
  let i = 0;
  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2;
            continue;
          }
          break;
        }
        i++;
      }
      i++;
      continue;
    }
    if (ch === "'") {
      let j = i + 1;
      let name = "";
      let closed = false;
      while (j < n) {
        if (formula[j] === "'") {
          if (formula[j + 1] === "'") {
            name += "'";
            j += 2;
            continue;
          }
          closed = true;
          break;
        }
        name += formula[j];
        j++;
      }
      if (!closed) break;
      if (formula[j + 1] !== "!") {
        i = j + 1;
        continue;
      }
      const close = name.lastIndexOf("]");
      const open = close < 0 ? -1 : name.lastIndexOf("[", close);
      if (open >= 0 && close > open + 1 && close < name.length - 1) {
        out.push({ start: i, end: j + 2, dir: name.slice(0, open), book: name.slice(open + 1, close), sheet: name.slice(close + 1) });
      }
      i = j + 2;
      continue;
    }
    if (ch === "[") {
      const prev = i > 0 ? formula[i - 1] : "";
      const close = formula.indexOf("]", i + 1);
      // `Table[Column]` and `[[#Headers],[Column]]` are not links.
      if (close > i + 1 && !(prev && (NAME_CHAR.test(prev) || prev === "]"))) {
        const book = formula.slice(i + 1, close);
        if (!/^[@#]/.test(book) && !/[['"!]/.test(book)) {
          let k = close + 1;
          while (k < n && NAME_CHAR.test(formula[k])) k++;
          if (k > close + 1 && formula[k] === "!") {
            out.push({ start: i, end: k + 1, dir: "", book, sheet: formula.slice(close + 1, k) });
            i = k + 1;
            continue;
          }
        }
      }
    }
    i++;
  }
  return out;
}

/** The formula with each link prefix replaced by what `to` returns (null keeps it as written). */
export function rewriteLinks(formula: string, to: (link: FoundLink) => string | null): string {
  const found = findLinks(formula);
  let out = formula;
  for (let i = found.length - 1; i >= 0; i--) {
    const replacement = to(found[i]);
    if (replacement !== null) out = out.slice(0, found[i].start) + replacement + out.slice(found[i].end);
  }
  return out;
}

/** `[NOW.xlsx]Model`: the name of the hidden copy, and of the link in formulas. */
export const linkSheetName = (book: string, sheet: string): string => `[${book}]${sheet}`;

export function splitLinkSheetName(name: string): { book: string; sheet: string } | null {
  if (!name.startsWith("[")) return null;
  const close = name.indexOf("]");
  if (close < 2 || close === name.length - 1) return null;
  return { book: name.slice(1, close), sheet: name.slice(close + 1) };
}

/** A sheet name as a formula prefix, always quoted: `'[NOW.xlsx]Model'!`. */
export const quotedPrefix = (sheetName: string): string => `'${sheetName.replace(/'/g, "''")}'!`;

/**
 * A link as an .xlsx file holds it: the other workbook by its number in the
 * file, `[1]Model!`, quoted when the sheet name is more than letters, digits,
 * `_` and `.`. That is the rule the engine prints sheet names by, so a
 * formula nobody changed reads the same as when the file was opened.
 */
export const numberedPrefix = (n: number, sheet: string): string =>
  /^[\p{Alphabetic}\p{N}_.]+$/u.test(sheet) ? `[${n}]${sheet}!` : `'[${n}]${sheet.replace(/'/g, "''")}'!`;

/** The link written the way people read it: `[NOW.xlsx]Model`, with its folder when it had one. */
export const linkText = (l: FoundLink): string => `${l.dir}[${l.book}]${l.sheet}`;
