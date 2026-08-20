import { colIndexToLetters, shiftFormulaCols } from "../formulaRuns";

/**
 * Read-side structural index of the active workbook, built from the same
 * Univer snapshot `captureWorkbookContext` uses.
 *
 * The prompt-time preview answers "what does the top of each sheet look
 * like"; this index answers "where IS something" — on sheets whose preview
 * is truncated the agent previously had to page through `read_range` probes
 * (each one a full model turn) to find a row. `describe_workbook` (the
 * section map) and `find_rows` (label search) are served from here in one
 * tool call.
 *
 * Also the canonical home of helpers that had drifted into three private
 * copies (row-label inference, content hashing, date-format detection) —
 * captureContext and the review UI import them from here now.
 *
 * Refresh strategy: memoized on the same content hash as the preview memo.
 * The hash covers formulas + stored values but NOT styles, so a pure
 * formatting change can serve a stale bold/section reading until the next
 * content edit — the preview memo makes the same trade, and section
 * structure moving on a format-only edit is rare enough to accept.
 */

/** A non-empty row, as the agent addresses it (1-indexed). */
export type RowEntry = {
  row: number;
  label: string | null;
  /** Count of value-only cells (no formula). */
  values: number;
  /** Count of formula cells. */
  formulas: number;
  /** Any bold cell in the row (resolved through the snapshot's style map). */
  bold: boolean;
  /**
   * 0-indexed column of the row's best "show me one evaluated cell" sample:
   * the rightmost formula cell, else the rightmost value cell. The index
   * stores the ADDRESS only — callers read the live grid at result time so
   * samples are never stale.
   */
  sampleCol: number;
};

export type SheetSection = {
  /** 1-indexed, inclusive. */
  startRow: number;
  endRow: number;
  title: string | null;
};

/**
 * One maximal run of column-translated formulas on a row (runs of length 1
 * are single formula cells). `formula` is the leftmost cell's text; the
 * rest of the run is recoverable via `shiftFormulaCols`.
 */
export type FormulaRun = {
  row: number; // 1-indexed
  startCol: number; // 0-indexed
  endCol: number;
  formula: string;
};

export type SheetIndex = {
  name: string;
  hidden: boolean;
  usedRange: string | null;
  /** 1-indexed row of the detected column-header (period) row, or null. */
  headerRow: number | null;
  /** Column letter → header text, e.g. { "BU": "FY2026E" }. */
  headers: Record<string, string>;
  rows: RowEntry[];
  sections: SheetSection[];
  formulaRuns: FormulaRun[];
  formulaRunsTruncated: boolean;
};

export type WorkbookIndex = {
  sheets: SheetIndex[];
  definedNames: Array<{ name: string; ref: string }>;
};

/** Formula-run entries kept per sheet — backstop against pathological sheets. */
const MAX_FORMULA_RUNS = 20_000;
/** Sections kept per sheet at build time. */
const MAX_SECTIONS = 500;
/** Header cells kept per sheet. */
const MAX_HEADERS = 150;
/** Rows examined when hunting for the header row. */
const HEADER_SCAN_ROWS = 50;
/** Row labels are clipped to this length everywhere they're inferred. */
const MAX_LABEL_CHARS = 60;

// ---------------------------------------------------------------------------
// Shared helpers (single source of truth — captureContext and the review UI
// import these instead of keeping private copies).
// ---------------------------------------------------------------------------

/**
 * Is this cell value usable as a human row label? Unifies three prior
 * heuristics (preview truncation index, review-modal region labels, row_map
 * readback): a string of ≥2 chars that doesn't lead with number/formula/
 * currency punctuation — rejects flag columns ("Y"), years ("2024A"),
 * numbers, and formula-ish text.
 */
export function isLabelText(v: unknown): v is string {
  if (typeof v !== "string") return false;
  const s = v.trim();
  if (s.length < 2) return false;
  return !/^[-+#0-9.($%=]/.test(s);
}

/** First label-worthy value in column order, trimmed and clipped. */
export function firstLabelIn(values: unknown[]): string | null {
  for (const v of values) {
    if (isLabelText(v)) return v.trim().slice(0, MAX_LABEL_CHARS);
  }
  return null;
}

/**
 * Does this number format render a date/time? (y/m/d/h tokens outside
 * quoted literals — "General", "#,##0.00", "0.0%" have none.)
 */
export function isDateFormat(nf: string): boolean {
  const withoutLiterals = nf.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "");
  return /[ymdh]/i.test(withoutLiterals);
}

/** Excel 1900-system serial → "YYYY-MM-DD" (UTC math, no timezone drift). */
export function serialToIsoDate(serial: number): string | null {
  if (!Number.isFinite(serial) || serial < 1 || serial > 2958465) return null; // 9999-12-31
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
  return new Date(ms).toISOString().slice(0, 10);
}

/**
 * Content hash of a snapshot: AUTHORED content only — formula text for
 * formula cells, stored values for literal cells. Never the computed value
 * of a formula cell: in a Univer snapshot that IS `cell.v`, and the worker
 * engine clears/rewrites it during recalc, so hashing it makes the key
 * drift with recalc timing and zero edits (a base captured mid-recalc then
 * "changed" once the engine settled — one such no-edit drift re-billed a
 * 181K-token cache write on a "thank you" turn). Byte-stable across recalc
 * jitter — the prompt-cache split and both memos (preview + index) key off
 * it.
 */
export function snapshotContentKey(snapshot: any): string {
  // Two FNV-1a passes with different seeds ≈ 64-bit key; collisions would
  // show a stale preview after an edit, so cheap-but-wide matters.
  let h1 = 0x811c9dc5;
  let h2 = 0xcbf29ce4;
  const mix = (s: string) => {
    for (let i = 0; i < s.length; i++) {
      const ch = s.charCodeAt(i);
      h1 ^= ch;
      h1 = Math.imul(h1, 0x01000193);
      h2 = Math.imul(h2 ^ ch, 0x01000193) + 7;
    }
  };
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets);
  for (const id of order) {
    const sheet = snapshot.sheets[id];
    if (!sheet) continue;
    // Visibility is part of the key: the preview renders hidden sheets
    // collapsed, so a mid-session hide/unhide must invalidate the memos
    // even though no cell content changed. NUL separator — it cannot
    // appear in a sheet name.
    mix(sheet.hidden === 1 ? `${sheet.name}\u0000hidden` : `${sheet.name}`);
    const cellData = sheet.cellData ?? {};
    for (const r of Object.keys(cellData)) {
      const row = cellData[r];
      for (const c of Object.keys(row)) {
        const cell = row[c];
        if (!cell) continue;
        const hasFormula = typeof cell.f === "string" && cell.f.trim() !== "";
        mix(hasFormula ? `${r},${c}=${cell.f}` : `${r},${c}=${cell.v ?? ""}`);
      }
    }
  }
  return `${(h1 >>> 0).toString(16)}-${(h2 >>> 0).toString(16)}:${order.length}`;
}

// ---------------------------------------------------------------------------
// Index builder
// ---------------------------------------------------------------------------

/**
 * Resolve a cell's style object: `s` is either an id into the workbook
 * style map or an inline style object.
 */
function styleOf(styles: any, cell: any): any | null {
  const s = cell?.s;
  if (!s) return null;
  if (typeof s === "string") return styles?.[s] ?? null;
  return typeof s === "object" ? s : null;
}

/** Header-cell rendering: strings verbatim, years as text, date serials as ISO. */
function headerTextOf(v: unknown, style: any): string | null {
  if (typeof v === "string" && v.trim()) return v.trim().slice(0, 40);
  if (typeof v === "number") {
    const pattern = style?.n?.pattern;
    if (typeof pattern === "string" && isDateFormat(pattern)) {
      const iso = serialToIsoDate(v);
      if (iso) return iso;
    }
    if (Number.isInteger(v) && v >= 1900 && v <= 2100) return String(v);
  }
  return null;
}

const normF = (f: unknown): string | null =>
  typeof f === "string" && f.trim() ? f.trim().replace(/^=/, "") : null;

export function buildWorkbookIndex(snapshot: any): WorkbookIndex {
  const styles = snapshot.styles ?? {};
  const order: string[] = snapshot.sheetOrder ?? Object.keys(snapshot.sheets ?? {});
  const sheets: SheetIndex[] = [];

  for (const sheetId of order) {
    const sheet = snapshot.sheets?.[sheetId];
    if (!sheet) continue;
    const cellData = sheet.cellData ?? {};
    const rowKeys = Object.keys(cellData)
      .map(Number)
      .sort((a, b) => a - b);

    const rows: RowEntry[] = [];
    const formulaRuns: FormulaRun[] = [];
    let formulaRunsTruncated = false;
    let maxUsedRow = -1;
    let maxUsedCol = -1;

    for (const r of rowKeys) {
      const rowCells = cellData[r];
      if (!rowCells) continue;
      const cols = Object.keys(rowCells)
        .map(Number)
        .sort((a, b) => a - b);

      let values = 0;
      let formulas = 0;
      let bold = false;
      let lastValueCol = -1;
      let lastFormulaCol = -1;
      const labelCandidates: unknown[] = [];

      for (const c of cols) {
        const cell = rowCells[c];
        if (!cell) continue;
        const hasF = cell.f !== undefined && cell.f !== null && cell.f !== "";
        const hasV = cell.v !== undefined && cell.v !== null && cell.v !== "";
        if (!hasF && !hasV) continue;
        if (r > maxUsedRow) maxUsedRow = r;
        if (c > maxUsedCol) maxUsedCol = c;
        if (hasF) {
          formulas++;
          lastFormulaCol = c;
        } else {
          values++;
          lastValueCol = c;
        }
        // Formula cells' cached values count as label candidates too — a
        // label pulled via formula (`=Assumptions!A5`) is still the label.
        if (hasV) labelCandidates.push(cell.v);
        if (!bold && styleOf(styles, cell)?.bl === 1) bold = true;
      }

      if (values + formulas === 0) continue;
      rows.push({
        row: r + 1,
        label: firstLabelIn(labelCandidates),
        values,
        formulas,
        bold,
        sampleCol: lastFormulaCol >= 0 ? lastFormulaCol : lastValueCol,
      });

      // Fill-run compression over this row's formula cells: consecutive
      // columns whose formulas are exact column-translations collapse to
      // one entry (same primitive as the preview's run detection).
      let i = 0;
      while (i < cols.length) {
        const f0 = normF(rowCells[cols[i]]?.f);
        if (!f0) {
          i++;
          continue;
        }
        let j = i;
        let prevF = f0;
        while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) {
          const nextF = normF(rowCells[cols[j + 1]]?.f);
          if (!nextF) break;
          const expected = shiftFormulaCols(prevF, 1);
          if (expected === null || expected !== nextF) break;
          prevF = nextF;
          j++;
        }
        if (formulaRuns.length < MAX_FORMULA_RUNS) {
          formulaRuns.push({
            row: r + 1,
            startCol: cols[i],
            endCol: cols[j],
            formula: `=${f0}`,
          });
        } else {
          formulaRunsTruncated = true;
        }
        i = j + 1;
      }
    }

    // Header (period) row: topmost row where the cells beyond the label
    // column are mostly text / years / date-formatted — the "FY2019A …
    // Q4-2031E" row of a financial model. Titles in row 1 don't match
    // (single cell); data rows don't match (numbers dominate).
    let headerRow: number | null = null;
    let headers: Record<string, string> = {};
    let scanned = 0;
    for (const r of rowKeys) {
      if (scanned++ >= HEADER_SCAN_ROWS) break;
      const rowCells = cellData[r];
      if (!rowCells) continue;
      let nonEmpty = 0;
      let headerLike = 0;
      const candidate: Record<string, string> = {};
      for (const key of Object.keys(rowCells)) {
        const c = Number(key);
        if (c === 0) continue; // label column
        const cell = rowCells[c];
        if (!cell) continue;
        const hasContent =
          (cell.v !== undefined && cell.v !== null && cell.v !== "") ||
          (cell.f !== undefined && cell.f !== null && cell.f !== "");
        if (!hasContent) continue;
        nonEmpty++;
        const text = headerTextOf(cell.v, styleOf(styles, cell));
        if (text !== null) {
          headerLike++;
          if (Object.keys(candidate).length < MAX_HEADERS) {
            candidate[colIndexToLetters(c)] = text;
          }
        }
      }
      if (headerLike >= 3 && headerLike >= 0.6 * nonEmpty) {
        headerRow = r + 1;
        headers = candidate;
        break;
      }
    }

    // The header row's "label" is usually just its first period cell
    // ("FY2019A") — noise as a row label or section title. The row is
    // already reported via header_row/headers, so drop that label; a real
    // label-column caption ("Fiscal year end") isn't in headers and stays.
    if (headerRow !== null) {
      const headerEntry = rows.find((e) => e.row === headerRow);
      if (
        headerEntry &&
        headerEntry.label !== null &&
        Object.values(headers).includes(headerEntry.label)
      ) {
        headerEntry.label = null;
      }
    }

    // Sections: contiguous row runs split on ≥2 blank rows, or on a
    // header-styled row (bold label with no data cells) which titles the
    // section it starts. Content-bearing bold rows ("Total revenue") do
    // NOT split — bold totals are ubiquitous inside sections.
    const sections: SheetSection[] = [];
    let current: SheetSection | null = null;
    for (const entry of rows) {
      const isHeaderStyled =
        entry.bold && entry.formulas === 0 && entry.values === 1 && entry.label !== null;
      const gap = current !== null ? entry.row - current.endRow : 0;
      if (current === null || gap > 2 || isHeaderStyled) {
        if (sections.length >= MAX_SECTIONS) {
          current = null;
          break;
        }
        current = { startRow: entry.row, endRow: entry.row, title: entry.label };
        sections.push(current);
      } else {
        current.endRow = entry.row;
        if (current.title === null && entry.label !== null) current.title = entry.label;
      }
    }

    sheets.push({
      name: String(sheet.name ?? sheetId),
      hidden: sheet.hidden === 1,
      usedRange:
        maxUsedRow >= 0 ? `A1:${colIndexToLetters(maxUsedCol)}${maxUsedRow + 1}` : null,
      headerRow,
      headers,
      rows,
      sections,
      formulaRuns,
      formulaRunsTruncated,
    });
  }

  return { sheets, definedNames: parseDefinedNames(snapshot) };
}

/**
 * Defined names from the snapshot's resources block (Univer's defined-name
 * plugin serializes `{ id: { name, formulaOrRefString, … } }`). Best-effort:
 * absent or unrecognized shapes yield [].
 */
function parseDefinedNames(snapshot: any): Array<{ name: string; ref: string }> {
  const out: Array<{ name: string; ref: string }> = [];
  const resources = snapshot?.resources;
  if (!Array.isArray(resources)) return out;
  for (const res of resources) {
    if (typeof res?.name !== "string" || !res.name.includes("DEFINED_NAME")) continue;
    try {
      const data = typeof res.data === "string" ? JSON.parse(res.data) : res.data;
      if (!data || typeof data !== "object") continue;
      for (const entry of Object.values(data) as any[]) {
        const name = entry?.name;
        const ref = entry?.formulaOrRefString;
        if (typeof name === "string" && name && typeof ref === "string" && ref) {
          out.push({ name, ref });
        }
      }
    } catch {
      /* unrecognized resource payload — skip */
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Memo
// ---------------------------------------------------------------------------

const indexMemo = new Map<string, { key: string; index: WorkbookIndex }>();

/**
 * Memoized index for a workbook path. Rebuilds lazily on the first call
 * after the snapshot's CONTENT changed (same key as the preview memo) —
 * no event wiring, no invalidation bugs possible.
 */
export function getWorkbookIndex(path: string, snapshot: any): WorkbookIndex | null {
  if (!snapshot || !snapshot.sheets) return null;
  const key = snapshotContentKey(snapshot);
  const memo = indexMemo.get(path);
  if (memo && memo.key === key) return memo.index;
  const index = buildWorkbookIndex(snapshot);
  indexMemo.set(path, { key, index });
  return index;
}

// ---------------------------------------------------------------------------
// Tool payloads
// ---------------------------------------------------------------------------

/** Sections shipped per sheet in a describe_workbook reply. */
const DESCRIBE_MAX_SECTIONS = 200;
/** Defined names shipped in a describe_workbook reply. */
const DESCRIBE_MAX_NAMES = 100;

/**
 * The `describe_workbook` tool reply: the orientation map (used range,
 * header row, section table-of-contents per sheet) WITHOUT row-level
 * detail — that's what find_rows / read_range are for.
 */
export function describeWorkbookPayload(index: WorkbookIndex, sheetFilter: string | null): object {
  const wanted = sheetFilter
    ? index.sheets.filter((s) => s.name === sheetFilter)
    : index.sheets;
  if (sheetFilter && wanted.length === 0) {
    return {
      error: `no sheet named "${sheetFilter}"`,
      sheets: index.sheets.map((s) => s.name),
    };
  }
  const payload: any = {
    sheets: wanted.map((s) => {
      const sections = s.sections.slice(0, DESCRIBE_MAX_SECTIONS).map((sec) => ({
        rows: sec.startRow === sec.endRow ? `${sec.startRow}` : `${sec.startRow}-${sec.endRow}`,
        title: sec.title,
      }));
      const entry: any = {
        name: s.name,
        used_range: s.usedRange,
        hidden: s.hidden,
        header_row: s.headerRow,
        headers: s.headers,
        sections,
      };
      if (s.sections.length > DESCRIBE_MAX_SECTIONS) {
        entry.sections_truncated = true;
        entry.section_count = s.sections.length;
      }
      return entry;
    }),
  };
  if (index.definedNames.length > 0) {
    payload.defined_names = index.definedNames.slice(0, DESCRIBE_MAX_NAMES);
    if (index.definedNames.length > DESCRIBE_MAX_NAMES) {
      payload.defined_names_truncated = true;
      payload.defined_name_count = index.definedNames.length;
    }
  }
  return payload;
}

export type FindRowsMatch = {
  sheet: string;
  /** 1-indexed. */
  row: number;
  label: string;
  /** Title of the section containing the row, when known. */
  section: string | null;
  /** Set on header-cell matches: the matched column's letter. */
  column?: string;
  /** 0-indexed sample column (row matches only) — caller reads it live. */
  sampleCol: number | null;
};

/**
 * Match score: exact > prefix > substring > word-initial fuzzy ("tot rev"
 * matches "Total revenue"). 0 = no match.
 */
function matchScore(label: string, query: string, tokens: string[]): number {
  const l = label.toLowerCase();
  if (l === query) return 4;
  if (l.startsWith(query)) return 3;
  if (l.includes(query)) return 2;
  const words = l.split(/[^a-z0-9]+/).filter(Boolean);
  if (
    tokens.length > 0 &&
    tokens.every((t) => words.some((w) => w.startsWith(t)))
  ) {
    return 1;
  }
  return 0;
}

function sectionTitleAt(sheet: SheetIndex, row: number): string | null {
  for (const sec of sheet.sections) {
    if (row >= sec.startRow && row <= sec.endRow) return sec.title;
  }
  return null;
}

/**
 * Search row labels and header cells across the index. Returns matches
 * best-first (score, then sheet order, then row). Samples are NOT
 * attached here — the caller reads `sampleCol` from the live grid so
 * values are current at result time.
 */
export function findRowsInIndex(
  index: WorkbookIndex,
  query: string,
  sheetFilter: string | null,
  maxResults: number,
): { matches: FindRowsMatch[]; total: number } {
  const q = query.trim().toLowerCase();
  const tokens = q.split(/\s+/).filter(Boolean);
  const scored: Array<{ score: number; order: number; match: FindRowsMatch }> = [];
  let order = 0;
  for (const sheet of index.sheets) {
    if (sheetFilter && sheet.name !== sheetFilter) continue;
    for (const entry of sheet.rows) {
      if (entry.label === null) continue;
      const score = matchScore(entry.label, q, tokens);
      if (score === 0) continue;
      scored.push({
        score,
        order: order++,
        match: {
          sheet: sheet.name,
          row: entry.row,
          label: entry.label,
          section: sectionTitleAt(sheet, entry.row),
          sampleCol: entry.sampleCol >= 0 ? entry.sampleCol : null,
        },
      });
    }
    if (sheet.headerRow !== null) {
      for (const [col, text] of Object.entries(sheet.headers)) {
        const score = matchScore(text, q, tokens);
        if (score === 0) continue;
        scored.push({
          score,
          order: order++,
          match: {
            sheet: sheet.name,
            row: sheet.headerRow,
            label: text,
            section: null,
            column: col,
            sampleCol: null,
          },
        });
      }
    }
  }
  scored.sort((a, b) => b.score - a.score || a.order - b.order);
  return {
    matches: scored.slice(0, maxResults).map((s) => s.match),
    total: scored.length,
  };
}
