/**
 * Column-translation of formulas — the primitive behind fill-run compression
 * in the agent's workbook preview.
 *
 * A quarterly model row is one formula filled across 50–100 columns; the
 * preview used to spell out every translated copy (measured: 294KB of
 * formula text on one real model). `shiftFormulaCols(f, 1)` computes what
 * Excel's fill-right produces, letting the capture detect "cells C..CX are
 * the same formula" and emit one line instead of a hundred.
 *
 * Mirrors the Rust translator in xlsx_patch/refs.rs: relative column letters
 * shift, absolute ($) parts and row numbers don't, string literals and
 * quoted sheet names are opaque, identifiers (LOG10, names) never match.
 */

const MAX_COL = 16_384; // XFD, 1-based

export function colLettersToIndex(letters: string): number | null {
  if (!letters || letters.length > 3) return null;
  let n = 0;
  for (const ch of letters) {
    const c = ch.toUpperCase().charCodeAt(0);
    if (c < 65 || c > 90) return null;
    n = n * 26 + (c - 64);
  }
  return n <= MAX_COL ? n - 1 : null;
}

export function colIndexToLetters(c: number): string {
  let n = c;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

const IDENT = /[A-Za-z0-9_.]/;

/**
 * Shift every RELATIVE column reference in `formula` by `dc`. Returns null
 * when the formula can't be shifted (a reference would leave the sheet).
 * Rows are never touched (fill-right only moves columns).
 */
export function shiftFormulaCols(formula: string, dc: number): string | null {
  let out = "";
  let i = 0;
  const n = formula.length;

  while (i < n) {
    const ch = formula[i];

    // Double-quoted string literal ("" escapes a quote) — opaque.
    if (ch === '"') {
      const start = i;
      i++;
      while (i < n) {
        if (formula[i] === '"') {
          if (formula[i + 1] === '"') {
            i += 2;
            continue;
          }
          i++;
          break;
        }
        i++;
      }
      out += formula.slice(start, i);
      continue;
    }

    // Quoted sheet name ('' escapes a quote) — opaque.
    if (ch === "'") {
      const start = i;
      i++;
      while (i < n) {
        if (formula[i] === "'") {
          if (formula[i + 1] === "'") {
            i += 2;
            continue;
          }
          i++;
          break;
        }
        i++;
      }
      out += formula.slice(start, i);
      continue;
    }

    const prevIsIdent = i > 0 && IDENT.test(formula[i - 1]);
    if (!prevIsIdent && (ch === "$" || /[A-Za-z]/.test(ch))) {
      const shifted = tryShiftRef(formula, i, dc);
      if (shifted === undefined) {
        // Not a reference — consume the whole identifier so its interior
        // can't be re-scanned as one.
        const start = i;
        while (i < n && (IDENT.test(formula[i]) || formula[i] === "$")) i++;
        if (i === start) i++;
        out += formula.slice(start, i);
        continue;
      }
      if (shifted === null) return null; // out of bounds
      out += shifted.text;
      i = shifted.end;
      continue;
    }

    out += ch;
    i += 1;
  }

  return out;
}

/**
 * Try to read a cell ref (or column-only range start) at `pos`. Returns:
 *   undefined — not a reference; caller treats it as an identifier
 *   null      — reference exists but shifting leaves the sheet
 *   {text, end} — the shifted text and the index just past the ref
 */
function tryShiftRef(
  s: string,
  pos: number,
  dc: number,
): { text: string; end: number } | null | undefined {
  let i = pos;
  const colAbs = s[i] === "$";
  if (colAbs) i++;
  const lettersStart = i;
  while (i < s.length && /[A-Za-z]/.test(s[i]) && i - lettersStart < 3) i++;
  const letters = s.slice(lettersStart, i);
  if (!letters) return undefined;

  // Column-only range: A:B, $A:$B
  if (s[i] === ":") {
    let j = i + 1;
    const abs2 = s[j] === "$";
    if (abs2) j++;
    const l2Start = j;
    while (j < s.length && /[A-Za-z]/.test(s[j]) && j - l2Start < 3) j++;
    const letters2 = s.slice(l2Start, j);
    const nextCh = s[j];
    if (
      letters2 &&
      !(nextCh !== undefined && (/[0-9$]/.test(nextCh) || IDENT.test(nextCh) || nextCh === "("))
    ) {
      const c1 = colLettersToIndex(letters);
      const c2 = colLettersToIndex(letters2);
      if (c1 === null || c2 === null) return undefined;
      const n1 = colAbs ? c1 : c1 + dc;
      const n2 = abs2 ? c2 : c2 + dc;
      if (n1 < 0 || n2 < 0 || n1 >= MAX_COL || n2 >= MAX_COL) return null;
      const text =
        (colAbs ? "$" : "") + colIndexToLetters(n1) + ":" + (abs2 ? "$" : "") + colIndexToLetters(n2);
      return { text, end: j };
    }
  }

  const rowAbs = s[i] === "$";
  let j = rowAbs ? i + 1 : i;
  const digitsStart = j;
  while (j < s.length && /[0-9]/.test(s[j])) j++;
  if (j === digitsStart) return undefined; // letters with no row → identifier
  const after = s[j];
  if (after !== undefined && (IDENT.test(after) || after === "(")) return undefined;

  const col = colLettersToIndex(letters);
  if (col === null) return undefined;
  const row = Number(s.slice(digitsStart, j));
  if (!Number.isFinite(row) || row < 1 || row > 1_048_576) return undefined;

  const newCol = colAbs ? col : col + dc;
  if (newCol < 0 || newCol >= MAX_COL) return null;

  const text = (colAbs ? "$" : "") + colIndexToLetters(newCol) + (rowAbs ? "$" : "") + row;
  return { text, end: j };
}

/**
 * Round a number for PREVIEW display: ~6 significant digits, exact integers
 * kept as-is. Full-precision floats (`13077.743432019966`) are pure token
 * waste in a preview the agent treats as read-only context.
 */
export function previewNumber(v: number): string {
  if (!Number.isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  return String(Number(v.toPrecision(6)));
}
