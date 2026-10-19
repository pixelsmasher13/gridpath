// SPDX-License-Identifier: Apache-2.0
// Ships in the gridpath package (packages/gridpath); licensed under the repository root LICENSE.
/**
 * Viewport grid — the big-model mode. The engine is the only copy of the
 * workbook; this component holds a tile cache of the cells near the
 * viewport and draws them on a canvas. Memory is flat in workbook size:
 * a 2.3M-cell sheet costs what a 20k-cell one does.
 *
 * Deliberately small editing surface for the first cut: selection with
 * mouse and keyboard, a formula bar, in-cell typing of values and
 * formulas, copy as TSV, and the agent's batches through the same handle
 * methods Univer serves. Writes go straight to the engine (journaled and
 * coalesced by the engine mirror); the engine's computed values come back
 * through the registered applier, which invalidates the affected tiles.
 *
 * Opt-in: localStorage gridpath.viewportGrid = "1" (see SpreadsheetScreen).
 */
import React, { useCallback, useEffect, useImperativeHandle, useMemo, useRef, useState } from "react";
import styled from "styled-components";
import type { GridEngine } from "../../../../packages/grid/src/engine";
import {
  DEFAULT_COL_PX,
  DEFAULT_ROW_PX,
  HEADER_H,
  HEADER_W,
  SheetGeometry,
  StyleBook,
  TILE_COLS,
  TILE_ROWS,
  TileCache,
  cellAlign,
  cellText,
  colLetters,
  borderStroke,
  fontString,
  outlineGroups,
  outlineGutter,
  OUTLINE_STEP,
  parseA1,
  tileKey,
  type EngineLayout,
  type EngineWindowCell,
  type GridCell,
  type ResolvedStyle,
  type SizeOverrides,
} from "../viewportModel";
import type { ManualSheetOp, UniverGridHandle } from "./UniverGrid";
import { buildWorkbookPatch, shiftFormulaA1, type BaselineCell } from "../surgicalPatch";
import type { CellFormatShape } from "../types";
import type { CellReader } from "../agent/captureContext";
import { stripImplicitIntersection as stripImplicitIntersection } from "../../../../packages/gridpath/src/workbook/snapshot";
import type { AutoFilterColumn } from "../../../../packages/gridpath/src/core/types";

/** Session state of one sheet's AutoFilter (0-based bounds; header row = r0).
 * Columns are keyed by absolute column index. */
type ColFilter = {
  values?: Set<string>;
  blank?: boolean;
  custom?: { and: boolean; items: Array<{ op: string; val: string }> };
  /** Criteria the model cannot evaluate: rows keep their file state. */
  other?: boolean;
};
type FilterState = { r0: number; c0: number; r1: number; c1: number; columns: Map<number, ColFilter> };

/** Excel's custom-filter comparison on one cell. */
function colFilterMatches(f: ColFilter, text: string, num: number): boolean {
  if (f.other) return true;
  if (f.values) return text === "" ? !!f.blank : f.values.has(text);
  if (f.custom) {
    const test = (op: string, val: string) => {
      const vn = Number(val);
      const numeric = Number.isFinite(num) && val.trim() !== "" && Number.isFinite(vn);
      if (numeric) {
        switch (op) {
          case "greaterThan": return num > vn;
          case "greaterThanOrEqual": return num >= vn;
          case "lessThan": return num < vn;
          case "lessThanOrEqual": return num <= vn;
          case "notEqual": return num !== vn;
          default: return num === vn;
        }
      }
      const a = text.toLowerCase(), b = val.toLowerCase();
      if (op === "equal" || op === "notEqual") {
        const re = new RegExp(`^${b.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`);
        const eq = re.test(a);
        return op === "equal" ? eq : !eq;
      }
      const cmp = a.localeCompare(b, undefined, { sensitivity: "base" });
      switch (op) {
        case "greaterThan": return cmp > 0;
        case "greaterThanOrEqual": return cmp >= 0;
        case "lessThan": return cmp < 0;
        case "lessThanOrEqual": return cmp <= 0;
        default: return cmp === 0;
      }
    };
    const results = f.custom.items.map((i) => test(i.op, i.val));
    return f.custom.and ? results.every(Boolean) : results.some(Boolean);
  }
  return true;
}

/** Value-list ordering: numbers first (numerically), then text, blanks last. */
function filterValueCompare(a: string, b: string): number {
  if (a === "") return b === "" ? 0 : 1;
  if (b === "") return -1;
  const na = Number(a.replace(/[,%$]/g, "")), nb = Number(b.replace(/[,%$]/g, ""));
  const fa = Number.isFinite(na), fb = Number.isFinite(nb);
  if (fa && fb) return na - nb;
  if (fa !== fb) return fa ? -1 : 1;
  return a.localeCompare(b, undefined, { sensitivity: "base" });
}

/** Excel's standard tab colours (theme accents and a few basics). */
const TAB_COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#8EAADB", "#F4B183", "#C00000", "#FF0000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0", "#000000"];

/** 100% on a Mac matches Excel for Mac's 72 DPI rendering (see viewScale). */
const PLATFORM_SCALE = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform) ? 0.75 : 1;

/** Preview region per sheet for the agent's opening context (0-based extents). */
const PREVIEW_ROWS = 512;
const PREVIEW_ROWS_HIDDEN = 256;
const PREVIEW_COLS = 128;
/** Below the region, the first columns are fetched so the truncation index
 * (row labels below the fold) reads real cells, as on the Univer path. */
const PREVIEW_LABEL_COLS = 8;

/** The viewport grid is the default for xlsx tabs; "0" brings Univer back
 * for one release while the remaining gaps close. */
export function viewportGridEnabled(): boolean {
  try {
    return localStorage.getItem("gridpath.viewportGrid") !== "0";
  } catch {
    return true;
  }
}

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

const Root = styled.div`
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  background: #fff;
  color: #1f1f1f;
  font-family: Calibri, Arial, sans-serif;
  user-select: none;
`;
const FormulaBar = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  height: 30px;
  padding: 0 8px;
  border-bottom: 1px solid #d9d9d9;
  background: #fafafa;
  font-size: 13px;
`;
const NameBox = styled.div`
  width: 90px;
  padding: 3px 6px;
  border: 1px solid #d0d0d0;
  border-radius: 3px;
  background: #fff;
  font-variant-numeric: tabular-nums;
`;
const FormulaInput = styled.input`
  flex: 1;
  height: 22px;
  border: 1px solid #d0d0d0;
  border-radius: 3px;
  padding: 0 6px;
  font: inherit;
  font-family: Menlo, Consolas, monospace;
  font-size: 12px;
  background: #fff;
`;
/** Grid area: the canvas is a fixed overlay underneath a transparent
 * scroll container. The canvas never moves (the scroll offset is applied
 * when drawing, rounded to whole pixels), so WebKit never has to resample
 * it at a fractional trackpad scroll position — a sticky canvas inside the
 * scroller rendered visibly soft on Retina. */
const GridArea = styled.div`
  position: relative;
  flex: 1;
  min-height: 0;
`;
const Scroller = styled.div`
  position: absolute;
  inset: 0;
  overflow: auto;
  outline: none;
  background: transparent;
`;
const Canvas = styled.canvas`
  position: absolute;
  top: 0;
  left: 0;
  display: block;
  pointer-events: none;
`;
/** Outline (group) gutters: drawn on their own canvas over the grid area's
 * left/top strips; the scroller and main canvas are offset to make room. */
const OutlineCanvas = styled.canvas`
  position: absolute;
  top: 0;
  left: 0;
  display: block;
  pointer-events: none;
`;
const OutlineStrip = styled.div`
  position: absolute;
  cursor: pointer;
`;
const Spacer = styled.div`
  position: absolute;
  top: 0;
  left: 0;
  width: 1px;
  height: 1px;
  pointer-events: none;
`;
const CellEditor = styled.input`
  position: absolute;
  z-index: 3;
  border: 2px solid #1a73e8;
  padding: 0 3px;
  margin: 0;
  outline: none;
  font: inherit;
  font-size: 13px;
  background: #fff;
  box-sizing: border-box;
`;
const TabBar = styled.div`
  display: flex;
  align-items: stretch;
  height: 28px;
  border-top: 1px solid #d9d9d9;
  background: #f3f3f3;
  font-size: 12px;
`;
const TabArrows = styled.div`
  display: flex;
  align-items: center;
  padding: 0 4px;
  gap: 2px;
  border-right: 1px solid #ddd;
  button { width: 22px; height: 22px; border: none; background: transparent; cursor: pointer; color: #555; font: inherit; border-radius: 3px; }
  button:hover { background: #e4e4e4; }
  button:disabled { color: #bbb; cursor: default; background: transparent; }
`;
const Tabs = styled.div`
  display: flex;
  align-items: stretch;
  flex: 1;
  min-width: 0;
  overflow-x: auto;
  scrollbar-width: none;
  &::-webkit-scrollbar { display: none; }
`;
/** Text colour that reads on a tab fill (white on dark colours). */
function tabTextOn(color: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(color.trim());
  if (!m) return "#222";
  const n = parseInt(m[1], 16);
  const lum = (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
  return lum < 0.6 ? "#fff" : "#222";
}
/** Excel's tab colours: an inactive tab is filled with its colour; the
 * active tab is white with a colour stripe along its bottom edge. */
const Tab = styled.button<{ $active: boolean; $hidden: boolean; $color?: string | null }>`
  position: relative;
  border: none;
  border-right: 1px solid #ddd;
  box-shadow: ${(p) => (p.$color && p.$active ? `inset 0 -3px 0 ${p.$color}` : "none")};
  padding: 0 12px;
  background: ${(p) => (p.$active ? "#fff" : p.$color ? p.$color : "transparent")};
  color: ${(p) => (p.$hidden ? "#999" : p.$color && !p.$active ? tabTextOn(p.$color) : "#222")};
  font: inherit;
  font-weight: ${(p) => (p.$active ? 600 : 400)};
  white-space: nowrap;
  cursor: pointer;
  &:hover { background: ${(p) => (p.$color && !p.$active ? p.$color : "#fff")}; filter: ${(p) => (p.$color && !p.$active ? "brightness(1.08)" : "none")}; }
`;
const Menu = styled.div`
  position: absolute;
  z-index: 5;
  min-width: 200px;
  padding: 4px 0;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 6px;
  box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
  font-size: 13px;
`;
const MenuItem = styled.button`
  display: block;
  width: 100%;
  padding: 6px 14px;
  border: none;
  background: transparent;
  text-align: left;
  font: inherit;
  cursor: pointer;
  &:hover { background: #eef3fb; }
  &:disabled { color: #aaa; cursor: default; }
`;
const MenuSep = styled.div`
  height: 1px;
  margin: 4px 0;
  background: #e6e6e6;
`;
const FindBar = styled.div`
  position: absolute;
  top: 34px;
  right: 16px;
  z-index: 6;
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 8px;
  background: #fff;
  border: 1px solid #d0d0d0;
  border-radius: 6px;
  box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
  font-size: 12px;
  input { width: 220px; height: 24px; border: 1px solid #d0d0d0; border-radius: 3px; padding: 0 6px; font: inherit; }
  button { border: 1px solid #d0d0d0; background: #fff; border-radius: 3px; padding: 2px 8px; font: inherit; cursor: pointer; }
`;
const TabInput = styled.input`
  width: 120px;
  margin: 2px 6px;
  font: inherit;
  font-size: 12px;
  padding: 2px 4px;
`;
const ZoomCtl = styled.span`
  display: inline-flex;
  align-items: center;
  margin-left: 10px;
  gap: 2px;
  button { border: 1px solid #d0d0d0; background: #fff; border-radius: 3px; min-width: 22px; height: 20px; padding: 0 5px; font: inherit; font-size: 11px; cursor: pointer; color: #444; }
  button:hover { background: #f0f0f0; }
`;
const Status = styled.div`
  margin-left: auto;
  padding: 0 10px;
  display: flex;
  align-items: center;
  color: #666;
  font-size: 11px;
  white-space: nowrap;
`;

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

interface Selection {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  /** Active cell (keyboard focus) — one corner of the range. */
  ar: number;
  ac: number;
}

interface Editing {
  row: number;
  col: number;
  text: string;
  /** Where typing happens: the in-cell box takes focus, the formula bar keeps it. */
  origin: "cell" | "bar";
}

const COLORS = {
  grid: "#e2e2e2",
  header: "#f4f4f4",
  headerText: "#555",
  headerSel: "#dbe6f7",
  selFill: "rgba(26,115,232,0.08)",
  /** Cells an agent changed that the user has not accepted yet. */
  pending: "rgba(251,188,4,0.28)",
  selBorder: "#1a73e8",
  frozen: "#8c8c8c",
};

export const ViewportGrid = React.memo(
  React.forwardRef<
    UniverGridHandle,
    {
      tabId: string;
      workbookPath: string | null;
      /** The workbook's engine; one stable object for the grid's lifetime. */
      engine: GridEngine;
      onUserEdit?: () => void;
      onManualSheetOp?: (op: ManualSheetOp) => void;
    }
  >(function ViewportGrid({ tabId, engine, onUserEdit, onManualSheetOp }, ref) {
    const scrollerRef = useRef<HTMLDivElement | null>(null);
    const canvasRef = useRef<HTMLCanvasElement | null>(null);
    const outlineCanvasRef = useRef<HTMLCanvasElement | null>(null);
    const gridAreaRef = useRef<HTMLDivElement | null>(null);
    /** Gutter sizes in CSS px for the active sheet (0 without groups). */
    const [gutter, setGutter] = useState({ w: 0, h: 0 });
    const gutterRef = useRef(gutter);
    type OutlineHit = { x0: number; y0: number; x1: number; y1: number; run: () => void };
    const outlineHitsRef = useRef<OutlineHit[]>([]);
    /** AutoFilter per sheet index; mirror entries per sheet name for the save. */
    const filtersRef = useRef<Map<number, FilterState>>(new Map());
    /** Sheets created this session whose layout refresh has not landed yet. */
    const pendingSheetsRef = useRef<Set<string>>(new Set());
    const autoFilterEditsRef = useRef<Map<string, { range: string | null; columns?: AutoFilterColumn[] }>>(new Map());
    const filterHitsRef = useRef<Array<{ x0: number; y0: number; x1: number; y1: number; col: number }>>([]);
    const [filterMenu, setFilterMenu] = useState<null | { sheet: number; col: number; x: number; y: number; search: string; items: Array<{ text: string; checked: boolean }>; loading: boolean }>(null);
    const spacerRef = useRef<HTMLDivElement | null>(null);
    const editorRef = useRef<HTMLInputElement | null>(null);

    const layoutRef = useRef<EngineLayout | null>(null);
    const geomsRef = useRef<Map<number, SheetGeometry>>(new Map());
    const stylesRef = useRef(new StyleBook());
    const tilesRef = useRef(new TileCache());
    const activeRef = useRef(0);
    const selRef = useRef<Selection>({ r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 });
    /** Additional selected ranges (Cmd+click / Cmd+drag); `selRef` is the active one. */
    const extraRanges = useRef<Array<{ r0: number; c0: number; r1: number; c1: number }>>([]);
    const scrollPosRef = useRef<Map<number, { x: number; y: number }>>(new Map());
    const drawReq = useRef(0);
    const contentVersionRef = useRef(0);
    const readyRef = useRef<{ promise: Promise<boolean>; resolve: (v: boolean) => void }>();
    /** Blocks of tiles to fetch, each one engine window (≤ 65,536 cells). */
    const fetchQueue = useRef<Array<{ sheet: number; tr0: number; tr1: number; tc0: number; tc1: number }>>([]);
    const lastScroll = useRef<{ x: number; y: number; dx: number; dy: number }>({ x: 0, y: 0, dx: 0, dy: 0 });
    const inflight = useRef(0);
    const sourceBytesRef = useRef<Uint8Array | null>(null);
    /** Cells written this session (user + agent), for the surgical save. */
    const editsRef = useRef<Map<string, Map<string, { f?: string; v?: unknown }>>>(new Map());
    const dragging = useRef(false);
    /** Last pointer position while dragging, for auto-scroll past the edges. */
    const dragPos = useRef<{ x: number; y: number } | null>(null);
    const autoScrollReq = useRef(0);
    /** Header-edge drag in progress. */
    const resizing = useRef<{ kind: "col" | "row"; index: number; start: number; size: number } | null>(null);
    /** Session-local sizes per sheet index (px; 0 = hidden). */
    const sizeOverridesRef = useRef<Map<number, SizeOverrides>>(new Map());
    /** Sizes changed this session, for the save (sheet name → index → px). */
    const sizeEditsRef = useRef<{ cols: Map<string, Map<number, number>>; rows: Map<string, Map<number, number>> }>({ cols: new Map(), rows: new Map() });
    /** Format overrides the session applied (agent + toolbar), by sheet name and "r,c". */
    const formatOverlayRef = useRef<Map<string, Map<string, CellFormatShape>>>(new Map());
    /** Every format change this session (agent, toolbar, keyboard, paste,
     * undo/redo), merged per cell, for the save. The overlay above is only
     * a display bridge: it drops each entry once the engine has the style. */
    const formatEditsRef = useRef<Map<string, Map<string, CellFormatShape>>>(new Map());
    /** Pending agent changes to tint, by sheet name and row*1e6+col (display only). */
    const highlightsRef = useRef<Map<string, Set<number>>>(new Map());
    /** Cells written or recalculated this session, per sheet index (row0*1e6+col0),
     * so the agent's preview delta covers changes outside the preview region. */
    const dirtyRef = useRef<Map<number, Set<number>>>(new Map());
    const calcHealthRef = useRef<any>(undefined);
    /** Undo stack of cell writes: each entry restores the inputs it replaced. */
    /** General undo stack: every entry knows how to undo and redo itself.
     * Entries are replayed strictly LIFO, so coordinates recorded before a
     * later row/column insert stay valid when that insert is undone first. */
    type UndoEntry = { label: string; undo: () => void; redo: () => void };
    const undoRef = useRef<UndoEntry[]>([]);
    const redoRef = useRef<UndoEntry[]>([]);
    /** While non-null, pushed entries collect here and become one composite. */
    const groupRef = useRef<UndoEntry[] | null>(null);
    const pushUndo = (entry: UndoEntry) => {
      if (groupRef.current) {
        groupRef.current.push(entry);
        return;
      }
      undoRef.current.push(entry);
      if (undoRef.current.length > 200) undoRef.current.shift();
      redoRef.current = [];
    };
    const withUndoGroup = (label: string, fn: () => void) => {
      if (groupRef.current) { fn(); return; }
      const entries: UndoEntry[] = [];
      groupRef.current = entries;
      try { fn(); } finally { groupRef.current = null; }
      if (entries.length === 1) pushUndo(entries[0]);
      else if (entries.length > 1) pushUndo({ label, undo: () => { for (let i = entries.length - 1; i >= 0; i--) entries[i].undo(); }, redo: () => { for (const e of entries) e.redo(); } });
    };

    const [sheetNames, setSheetNames] = useState<Array<{ name: string; hidden: boolean; color?: string | null }>>([]);
    const [active, setActive] = useState(0);
    const [nameBox, setNameBox] = useState("A1");
    const [formulaText, setFormulaText] = useState("");
    const [editing, setEditing] = useState<Editing | null>(null);
    const [status, setStatus] = useState("");
    const [stats, setStats] = useState("");
    const [calculating, setCalculating] = useState(false);
    const [menu, setMenu] = useState<{ x: number; y: number; header?: "row" | "col" } | null>(null);
    const [tabMenu, setTabMenu] = useState<{ x: number; y: number; index: number } | null>(null);
    const [renaming, setRenaming] = useState<{ index: number; text: string } | null>(null);
    const [colorPick, setColorPick] = useState<number | null>(null);
    // Any click outside an open menu, or Escape, closes it (Excel's behaviour).
    useEffect(() => {
      if (!menu && !tabMenu && colorPick === null && !filterMenu) return;
      const onDown = (ev: MouseEvent) => {
        if ((ev.target as HTMLElement | null)?.closest?.("[data-vp-menu]")) return;
        setMenu(null); setTabMenu(null); setColorPick(null); setFilterMenu(null);
      };
      const onKey = (ev: KeyboardEvent) => { if (ev.key === "Escape") { setMenu(null); setTabMenu(null); setColorPick(null); setFilterMenu(null); } };
      document.addEventListener("mousedown", onDown, true);
      document.addEventListener("keydown", onKey, true);
      return () => { document.removeEventListener("mousedown", onDown, true); document.removeEventListener("keydown", onKey, true); };
    }, [menu, tabMenu, colorPick, filterMenu]);
    const setTabColor = (index: number, color: string | null) => {
      const l = layoutRef.current?.sheets[index];
      if (!l) return;
      l.color = color;
      setSheetNames(layoutRef.current!.sheets.map((x) => ({ name: x.name, hidden: x.hidden, color: x.color })));
      onManualSheetOp?.({ kind: "tabColor", name: l.name, color });
      onUserEdit?.();
    };
    /** Excel keeps hidden sheets out of the tab bar; this reveals them greyed, for a look. */
    const [showHidden, setShowHidden] = useState(false);
    const tabsRef = useRef<HTMLDivElement | null>(null);
    /** View zoom (1 = 100%). Sheet geometry stays in sheet px; the canvas
     * transform, the spacer and every screen↔sheet conversion apply it,
     * multiplied by the platform's idea of 100%: Excel for Mac renders at
     * 72 DPI (an 11pt font is 11px, a 15pt row 15px), Windows and the
     * sheet geometry here at 96 DPI, so on a Mac 100% is drawn at 0.75. */
    const zoomRef = useRef(1);
    const viewScale = () => zoomRef.current * PLATFORM_SCALE;
    /** Zoom per sheet, seeded from the file's saved zoom, kept as the user changes it. */
    const sheetZoomRef = useRef<Map<number, number>>(new Map());
    const zoomForSheet = (sheet: number) => sheetZoomRef.current.get(sheet) ?? Math.max(0.5, Math.min(2, (layoutRef.current?.sheets[sheet]?.zoom ?? 100) / 100));
    const [zoom, setZoom] = useState(1);
    const setZoomLevelRef = useRef<(next: number) => void>(() => {});
    const setZoomLevel = (next: number) => {
      const z = Math.max(0.5, Math.min(2, Math.round(next * 20) / 20));
      const sc = scrollerRef.current;
      const prev = zoomRef.current;
      if (z === prev) return;
      zoomRef.current = z;
      sheetZoomRef.current.set(activeRef.current, z);
      setZoom(z);
      updateSpacer();
      if (sc) { sc.scrollLeft = (sc.scrollLeft / prev) * z; sc.scrollTop = (sc.scrollTop / prev) * z; }
      scheduleDraw();
    };
    setZoomLevelRef.current = setZoomLevel;
    const [tabScroll, setTabScroll] = useState({ left: 0, max: 0 });
    const measureTabs = () => {
      const el = tabsRef.current;
      if (el) setTabScroll({ left: el.scrollLeft, max: Math.max(0, el.scrollWidth - el.clientWidth) });
    };
    const setSheetVisibility = (index: number, hidden: boolean) => {
      const l = layoutRef.current?.sheets[index];
      if (!l) return;
      const visible = sheetNames.filter((x, i) => !x.hidden || i === index).length;
      if (hidden && visible <= 1) return;
      l.hidden = hidden;
      setSheetNames(layoutRef.current!.sheets.map((x) => ({ name: x.name, hidden: x.hidden, color: x.color })));
      onManualSheetOp?.({ kind: "visibility", name: l.name, hidden });
      if (hidden && index === activeRef.current) {
        const next = layoutRef.current!.sheets.findIndex((x, i) => !x.hidden && i !== index);
        if (next >= 0) activateSheet(next);
      }
      bumpVersion();
      onUserEdit?.();
    };
    const [find, setFind] = useState<{ query: string; hits: Array<{ sheet: number; row: number; column: number; text: string }>; at: number; scope: "sheet" | "workbook"; replace: string } | null>(null);
    const findInputRef = useRef<HTMLInputElement | null>(null);
    /** Span of the reference last inserted by clicking while editing a formula. */
    const refPick = useRef<{ start: number; end: number; anchor: { r: number; c: number } } | null>(null);
    const refDrag = useRef(false);
    /** Ranges the formula being edited references on the active sheet (for outlines). */
    const editRefsRef = useRef<Array<{ r0: number; c0: number; r1: number; c1: number }>>([]);
    const barRef = useRef<HTMLInputElement | null>(null);
    /** What the last in-grid copy held, so a paste back into the grid keeps
     * formulas (shifted) and formats; matched against the clipboard text. */
    const clipRef = useRef<{ text: string; r0: number; c0: number; cells: Array<{ dr: number; dc: number; input: string; format: CellFormatShape | null }> } | null>(null);
    /** Border overrides the session applied, by sheet name and "r,c". */
    const borderOverlayRef = useRef<Map<string, Map<string, { top?: boolean; bottom?: boolean; left?: boolean; right?: boolean }>>>(new Map());
    /** Notes by sheet name and "r,c" (display only; the save mirror carries them). */
    const notesRef = useRef<Map<string, Map<string, string>>>(new Map());
    const [hoverNote, setHoverNote] = useState<{ x: number; y: number; text: string } | null>(null);
    /** Defined names: layout's plus the session's (name → ref). */
    const namesRef = useRef<Map<string, string>>(new Map());
    /** Sheet-snapshot captures in flight (see getSheetSnapshot / deleteSheet). */
    const capturesRef = useRef<Array<Promise<unknown>>>([]);
    /** Fill-handle drag: the source range and the current target extent. */
    const filling = useRef<{ r0: number; c0: number; r1: number; c1: number; toR: number; toC: number } | null>(null);
    const structuralRef = useRef<(op: string, sheetName: string, index1: number, count: number) => boolean>(() => false);
    const prefetchRef = useRef<(sheet: number, r0: number, r1: number, c0: number, c1: number) => Promise<void>>(async () => {});
    const applyFormatRef = useRef<(sheetName: string, sr: number, sc: number, er: number, ec: number, format: CellFormatShape, history?: "record" | "skip") => void>(() => {});
    /** Row/column ops the user made from the grid (the agent's come via the mirror). */
    const manualRowColOpsRef = useRef<Array<{ kind: "insertRows" | "deleteRows" | "insertColumns" | "deleteColumns"; sheet: string; before?: number; start?: number; count: number }>>([]);

    if (!readyRef.current) {
      let resolve: (v: boolean) => void = () => {};
      const promise = new Promise<boolean>((r) => (resolve = r));
      readyRef.current = { promise, resolve };
    }

    const geom = (sheet: number): SheetGeometry | null => {
      const g = geomsRef.current.get(sheet);
      if (g) return g;
      const l = layoutRef.current?.sheets[sheet];
      if (!l) return null;
      const ng = new SheetGeometry(l, sizeOverridesRef.current.get(sheet));
      geomsRef.current.set(sheet, ng);
      return ng;
    };

    const sheetNameOf = (sheet: number): string => layoutRef.current?.sheets[sheet]?.name ?? "";

    const overridesFor = (sheet: number): SizeOverrides => {
      let o = sizeOverridesRef.current.get(sheet);
      if (!o) sizeOverridesRef.current.set(sheet, (o = { rows: new Map(), cols: new Map() }));
      return o;
    };

    /** Set a column's or row's px size for display, optionally recording it for the save. */
    const setSize = (sheet: number, kind: "col" | "row", index: number, px: number, record: boolean, history: "record" | "skip" = "record") => {
      const o = overridesFor(sheet);
      if (record && history === "record") {
        const g0 = geom(sheet);
        const before = g0 ? (kind === "col" ? g0.colWidth(index) : g0.rowHeight(index)) : px;
        const after = Math.max(0, Math.round(px));
        if (before !== after) pushUndo({ label: `${kind} size`, undo: () => setSize(sheet, kind, index, before, true, "skip"), redo: () => setSize(sheet, kind, index, after, true, "skip") });
      }
      (kind === "col" ? o.cols : o.rows).set(index, Math.max(0, Math.round(px)));
      geomsRef.current.delete(sheet);
      if (record) {
        const name = layoutRef.current?.sheets[sheet]?.name;
        if (name) {
          const store = kind === "col" ? sizeEditsRef.current.cols : sizeEditsRef.current.rows;
          let per = store.get(name);
          if (!per) store.set(name, (per = new Map()));
          per.set(index, Math.max(0, Math.round(px)));
        }
      }
      if (sheet === activeRef.current) updateSpacer();
      bumpVersion();
      scheduleDraw();
    };

    const currentMerges = (sheet: number): string[] => {
      const l = layoutRef.current?.sheets[sheet];
      return [...(sizeOverridesRef.current.get(sheet)?.merges ?? l?.merges ?? [])];
    };
    const setMerges = (sheet: number, list: string[], history: "record" | "skip") => {
      if (history === "record") {
        const before = currentMerges(sheet);
        pushUndo({ label: "merge", undo: () => setMerges(sheet, before, "skip"), redo: () => setMerges(sheet, list, "skip") });
      }
      overridesFor(sheet).merges = list;
      geomsRef.current.delete(sheet);
      bumpVersion();
      scheduleDraw();
    };
    const setFrozen = (sheet: number, rows: number, cols: number, history: "record" | "skip") => {
      if (history === "record") {
        const g0 = geom(sheet);
        const before = { rows: g0?.frozenRows ?? 0, cols: g0?.frozenCols ?? 0 };
        if (before.rows !== rows || before.cols !== cols) pushUndo({ label: "freeze", undo: () => setFrozen(sheet, before.rows, before.cols, "skip"), redo: () => setFrozen(sheet, rows, cols, "skip") });
      }
      overridesFor(sheet).frozen = { rows, cols };
      geomsRef.current.delete(sheet);
      if (sheet === activeRef.current) updateSpacer();
      scheduleDraw();
    };

    const overlayOf = (sheet: number, r: number, c: number): CellFormatShape | undefined => {
      const name = layoutRef.current?.sheets[sheet]?.name;
      if (!name) return undefined;
      const per = formatOverlayRef.current.get(name);
      return per && per.size > 0 ? per.get(`${r},${c}`) : undefined;
    };

    /** Conditional formatting wins over the cell's own style for what it sets. */
    const withCf = (st: ResolvedStyle, cf: GridCell["cf"]): ResolvedStyle => {
      if (!cf) return st;
      return {
        ...st,
        fill: cf.fill !== undefined ? cf.fill : st.fill,
        color: cf.color ? cf.color : st.color,
        bold: cf.bold ?? st.bold,
        italic: cf.italic ?? st.italic,
        underline: cf.underline ?? st.underline,
        strike: cf.strike ?? st.strike,
      };
    };

    /** A resolved engine style with the session's format overlay applied. */
    const styleWithOverlay = (st: ResolvedStyle, ov: CellFormatShape | undefined): ResolvedStyle => {
      if (!ov) return st;
      return {
        ...st,
        bold: ov.bold ?? st.bold,
        italic: ov.italic ?? st.italic,
        underline: ov.underline ?? st.underline,
        strike: ov.strike ?? st.strike,
        color: ov.font_color ?? st.color,
        fill: ov.background_color === undefined ? st.fill : ov.background_color && ov.background_color !== "" ? ov.background_color : null,
        sizePt: ov.font_size ?? st.sizePt,
        family: ov.font_family ? `"${ov.font_family}"` : st.family,
        align: ov.horizontal_align ?? st.align,
        vertical: ov.vertical_align ?? st.vertical,
        wrap: ov.wrap_text ?? st.wrap,
      };
    };

    const sheetIndexOf = (name: string): number => {
      const i = layoutRef.current?.sheets.findIndex((s) => s.name === name) ?? -1;
      if (i >= 0) return i;
      return engine.sheetNames().indexOf(name);
    };

    const bumpVersion = () => {
      contentVersionRef.current++;
    };

    // ----- drawing --------------------------------------------------------

    const scheduleDraw = useCallback(() => {
      if (drawReq.current) return;
      drawReq.current = requestAnimationFrame(() => {
        drawReq.current = 0;
        draw();
      });
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const updateSpacer = useCallback(() => {
      const g = geom(activeRef.current);
      const sp = spacerRef.current;
      if (!g || !sp) return;
      const z = viewScale();
      sp.style.width = `${(g.totalWidth + HEADER_W + 200) * z}px`;
      sp.style.height = `${(g.totalHeight + HEADER_H + 200) * z}px`;
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    /** Visible 0-based cell rectangle in the scrolling (non-frozen) pane. */
    const visibleRect = () => {
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!sc || !g) return null;
      const z = viewScale();
      const x0 = Math.round(sc.scrollLeft) / z;
      const y0 = Math.round(sc.scrollTop) / z;
      const w = sc.clientWidth / z - HEADER_W;
      const h = sc.clientHeight / z - HEADER_H;
      const r0 = g.rowAt(y0);
      const r1 = Math.min(g.rowCount - 1, g.rowAt(y0 + h) + 1);
      const c0 = g.colAt(x0);
      const c1 = Math.min(g.colCount - 1, g.colAt(x0 + w) + 1);
      return { r0, r1, c0, c1, x0, y0, w, h };
    };

    const requestTiles = useCallback(() => {
      const g = geom(activeRef.current);
      const vis = visibleRect();
      if (!g || !vis) return;
      const sheet = activeRef.current;
      // Margin: half a tile each way, plus a full viewport ahead in the
      // direction the user is scrolling, so the fetch lands before the
      // cells are visible.
      const ls = lastScroll.current;
      const dx = vis.x0 - ls.x, dy = vis.y0 - ls.y;
      lastScroll.current = { x: vis.x0, y: vis.y0, dx: dx || ls.dx, dy: dy || ls.dy };
      const visRows = vis.r1 - vis.r0 + 1, visCols = vis.c1 - vis.c0 + 1;
      const ahead = { up: 0, down: 0, left: 0, right: 0 };
      if (lastScroll.current.dy > 0) ahead.down = visRows; else if (lastScroll.current.dy < 0) ahead.up = visRows;
      if (lastScroll.current.dx > 0) ahead.right = visCols; else if (lastScroll.current.dx < 0) ahead.left = visCols;
      const r0 = Math.max(0, vis.r0 - TILE_ROWS / 2 - ahead.up);
      const r1 = Math.min(g.rowCount - 1, vis.r1 + TILE_ROWS / 2 + ahead.down);
      const c0 = Math.max(0, vis.c0 - TILE_COLS / 2 - ahead.left);
      const c1 = Math.min(g.colCount - 1, vis.c1 + TILE_COLS / 2 + ahead.right);
      const wanted = tilesRef.current.missing(sheet, r0, r1, c0, c1);
      if (g.frozenRows > 0) wanted.push(...tilesRef.current.missing(sheet, 0, g.frozenRows - 1, c0, c1));
      if (g.frozenCols > 0) wanted.push(...tilesRef.current.missing(sheet, r0, r1, 0, g.frozenCols - 1));
      // Tiles entirely beyond the used range are empty — mark loaded, no fetch.
      const seen = new Set<string>();
      const need: Array<{ tr: number; tc: number }> = [];
      for (const t of wanted) {
        if (seen.has(t.key)) continue;
        seen.add(t.key);
        if (t.tr * TILE_ROWS >= g.layout.maxRow || t.tc * TILE_COLS >= g.layout.maxCol) {
          tilesRef.current.putTile(t.key, []);
          continue;
        }
        need.push(t);
      }
      // One engine window per block: vertical runs of tiles in a tile
      // column, then adjacent columns with the same run merged, while the
      // block stays within the window limit. A sideways crossing on a tall
      // viewport is one request instead of one per tile row.
      const MAX_TILES = Math.floor(65_536 / (TILE_ROWS * TILE_COLS));
      need.sort((a, b) => a.tc - b.tc || a.tr - b.tr);
      const runs: Array<{ tc0: number; tc1: number; tr0: number; tr1: number }> = [];
      for (const t of need) {
        const last = runs[runs.length - 1];
        if (last && last.tc0 === t.tc && last.tc1 === t.tc && last.tr1 === t.tr - 1 && last.tr1 - last.tr0 + 2 <= MAX_TILES) last.tr1 = t.tr;
        else runs.push({ tc0: t.tc, tc1: t.tc, tr0: t.tr, tr1: t.tr });
      }
      const blocks: typeof runs = [];
      for (const run of runs) {
        const last = blocks[blocks.length - 1];
        const h = run.tr1 - run.tr0 + 1;
        if (last && last.tr0 === run.tr0 && last.tr1 === run.tr1 && last.tc1 === run.tc0 - 1 && (last.tc1 - last.tc0 + 2) * h <= MAX_TILES) last.tc1 = run.tc0;
        else blocks.push({ ...run });
      }
      for (const b of blocks) {
        for (let tr = b.tr0; tr <= b.tr1; tr++) for (let tc = b.tc0; tc <= b.tc1; tc++) tilesRef.current.markPending(tileKey(sheet, tr, tc));
        fetchQueue.current.push({ sheet, ...b });
      }
      pump();
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    /** Resolves when the visible tiles are loaded (or after `timeoutMs`). */
    const firstTilesLanded = (timeoutMs: number) =>
      new Promise<void>((resolve) => {
        const t0 = performance.now();
        const tick = () => {
          requestTiles();
          if ((inflight.current === 0 && fetchQueue.current.length === 0) || performance.now() - t0 > timeoutMs) resolve();
          else setTimeout(tick, 16);
        };
        tick();
      });

    /** One tile from the engine; the tile is marked pending until it lands.
     * A tile already loaded stays readable (with any values applied in
     * place) until the fresh window replaces it. */
    const fetchTile = async (sheet: number, tr: number, tc: number): Promise<void> => {
      const key = tileKey(sheet, tr, tc);
      tilesRef.current.markPending(key);
      const r0 = tr * TILE_ROWS + 1;
      const c0 = tc * TILE_COLS + 1;
      try {
        const win = await engine.fetchWindow(sheet, r0, r0 + TILE_ROWS - 1, c0, c0 + TILE_COLS - 1);
        if (win && Array.isArray(win.cells)) {
          stylesRef.current.merge(win.styles ?? {});
          tilesRef.current.putTile(key, win.cells);
        } else tilesRef.current.clearPending(key);
      } catch {
        tilesRef.current.clearPending(key);
      }
    };

    /** One engine window covering a block of tiles, split into tiles on arrival. */
    const fetchBlock = async (b: { sheet: number; tr0: number; tr1: number; tc0: number; tc1: number }): Promise<void> => {
      const keys: string[] = [];
      for (let tr = b.tr0; tr <= b.tr1; tr++) for (let tc = b.tc0; tc <= b.tc1; tc++) keys.push(tileKey(b.sheet, tr, tc));
      try {
        const win = await engine.fetchWindow(b.sheet, b.tr0 * TILE_ROWS + 1, (b.tr1 + 1) * TILE_ROWS, b.tc0 * TILE_COLS + 1, (b.tc1 + 1) * TILE_COLS);
        if (!win || !Array.isArray(win.cells)) {
          for (const k of keys) tilesRef.current.clearPending(k);
          return;
        }
        stylesRef.current.merge(win.styles ?? {});
        const buckets = new Map<string, EngineWindowCell[]>();
        for (const k of keys) buckets.set(k, []);
        for (const c of win.cells) {
          const k = tileKey(b.sheet, Math.floor((c.r - 1) / TILE_ROWS), Math.floor((c.c - 1) / TILE_COLS));
          buckets.get(k)?.push(c);
        }
        for (const [k, cells] of buckets) tilesRef.current.putTile(k, cells);
      } catch {
        for (const k of keys) tilesRef.current.clearPending(k);
      }
    };

    const pump = () => {
      while (inflight.current < 2 && fetchQueue.current.length > 0) {
        // Most recently requested first: closest to where the user is now.
        const job = fetchQueue.current.pop()!;
        if (job.sheet !== activeRef.current) {
          for (let tr = job.tr0; tr <= job.tr1; tr++) for (let tc = job.tc0; tc <= job.tc1; tc++) tilesRef.current.clearPending(tileKey(job.sheet, tr, tc));
          continue;
        }
        inflight.current++;
        void fetchBlock(job).finally(() => {
          inflight.current--;
          scheduleDraw();
          pump();
        });
      }
    };

    const draw = () => {
      const canvas = canvasRef.current;
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!canvas || !sc || !g) return;
      const dpr = window.devicePixelRatio || 1;
      const z = viewScale();
      const cssW = sc.clientWidth, cssH = sc.clientHeight;
      if (canvas.width !== Math.round(cssW * dpr) || canvas.height !== Math.round(cssH * dpr)) {
        canvas.width = Math.round(cssW * dpr);
        canvas.height = Math.round(cssH * dpr);
        canvas.style.width = `${cssW}px`;
        canvas.style.height = `${cssH}px`;
      }
      const W = cssW / z, H = cssH / z;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;
      ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
      ctx.clearRect(0, 0, W, H);
      ctx.fillStyle = "#fff";
      ctx.fillRect(0, 0, W, H);

      const sheet = activeRef.current;
      filterHitsRef.current = [];
      const scrollX = Math.round(sc.scrollLeft) / z;
      const scrollY = Math.round(sc.scrollTop) / z;
      const fy = g.rowTop[g.frozenRows]; // frozen band height in sheet px
      const fx = g.colLeft[g.frozenCols];
      const sel = selRef.current;
      const tiles = tilesRef.current;
      const styles = stylesRef.current;

      // Draw one pane: rows [rA, rB) × cols [cA, cB) with the given pixel offsets.
      const pane = (
        rowFrom: number, rowTo: number, colFrom: number, colTo: number,
        offX: number, offY: number, clipX: number, clipY: number, clipW: number, clipH: number,
      ) => {
        if (clipW <= 0 || clipH <= 0) return;
        ctx.save();
        ctx.beginPath();
        ctx.rect(clipX, clipY, clipW, clipH);
        ctx.clip();
        // Gridlines.
        if (g.layout.showGridLines) {
          ctx.strokeStyle = COLORS.grid;
          ctx.lineWidth = 1;
          ctx.beginPath();
          for (let r = rowFrom; r <= rowTo; r++) {
            const y = Math.round(g.rowTop[r] - offY) + 0.5;
            if (y < clipY - 1 || y > clipY + clipH + 1) continue;
            ctx.moveTo(clipX, y);
            ctx.lineTo(clipX + clipW, y);
          }
          for (let c = colFrom; c <= colTo; c++) {
            const x = Math.round(g.colLeft[c] - offX) + 0.5;
            if (x < clipX - 1 || x > clipX + clipW + 1) continue;
            ctx.moveTo(x, clipY);
            ctx.lineTo(x, clipY + clipH);
          }
          ctx.stroke();
        }
        // Cells, two passes: every fill and border first, then every text,
        // so text that overflows its cell is never painted over by the
        // neighbour's fill. Overflow follows Excel: left-aligned text spills
        // right, right-aligned spills left, centered spills both ways, and
        // only into cells with no content.
        const rects: Array<{ r: number; c: number; x: number; y: number; w: number; h: number; merge: boolean }> = [];
        for (let r = rowFrom; r < rowTo; r++) {
          const rh = g.rowHeight(r);
          if (rh === 0) continue;
          const y = g.rowTop[r] - offY;
          for (let c = colFrom; c < colTo; c++) {
            const cw = g.colWidth(c);
            if (cw === 0) continue;
            const x = g.colLeft[c] - offX;
            const merge = g.mergeOf(r, c);
            if (merge && (merge.r0 !== r || merge.c0 !== c)) continue;
            let w = cw;
            let h = rh;
            if (merge) {
              w = g.colLeft[merge.c1 + 1] - g.colLeft[merge.c0];
              h = g.rowTop[merge.r1 + 1] - g.rowTop[merge.r0];
            }
            rects.push({ r, c, x, y, w, h, merge: !!merge });
            const cell = tiles.cell(sheet, r, c);
            let st = withCf(styleWithOverlay(styles.get(cell?.s), overlayOf(sheet, r, c)), cell?.cf);
            const bo = borderOverlayRef.current.get(sheetNameOf(sheet))?.get(`${r},${c}`);
            if (bo) {
              const line = { style: "thin", color: "#000000" };
              const borders = { ...st.borders };
              for (const side of ["top", "bottom", "left", "right"] as const) {
                if (bo[side] === true) borders[side] = line;
                else if (bo[side] === false) delete borders[side];
              }
              st = { ...st, borders };
            }
            if (st.fill) {
              ctx.fillStyle = st.fill;
              ctx.fillRect(x, y, w, h);
            } else if (merge) {
              ctx.fillStyle = "#fff";
              ctx.fillRect(x + 1, y + 1, w - 1, h - 1);
            }
            const db = cell?.cf?.dataBar;
            if (db && typeof cell?.v === "number") {
              // Data bar: from the axis toward the value, inside the cell.
              const axisX = x + 1 + (w - 2) * Math.max(0, Math.min(1, db.axis));
              const span = (w - 2) * Math.max(0, Math.min(1, Math.abs(db.value)));
              const neg = cell.v < 0;
              const bx = neg ? axisX - span : axisX;
              ctx.fillStyle = (neg ? db.negative : db.positive) ?? "#638EC6";
              ctx.globalAlpha = db.gradient ? 0.55 : 0.8;
              ctx.fillRect(bx, y + 2, span, h - 4);
              ctx.globalAlpha = 1;
            }
            const bd = st.borders;
            if (bd.top || bd.bottom || bd.left || bd.right) {
              const edge = (line: { style: string; color: string }, x1: number, y1: number, x2: number, y2: number, horizontal: boolean) => {
                const { width, dash, double } = borderStroke(line.style);
                ctx.strokeStyle = line.color;
                ctx.setLineDash(dash);
                if (double) {
                  ctx.lineWidth = 1;
                  const d = horizontal ? [0, 2] : [0, 2];
                  for (const off of [-1, 1]) {
                    ctx.beginPath();
                    ctx.moveTo(x1 + (horizontal ? 0 : off), y1 + (horizontal ? off : 0));
                    ctx.lineTo(x2 + (horizontal ? 0 : off), y2 + (horizontal ? off : 0));
                    ctx.stroke();
                  }
                  void d;
                } else {
                  ctx.lineWidth = width;
                  ctx.beginPath();
                  const snap = width % 2 === 1 ? 0.5 : 0;
                  ctx.moveTo(horizontal ? x1 : Math.round(x1) + snap, horizontal ? Math.round(y1) + snap : y1);
                  ctx.lineTo(horizontal ? x2 : Math.round(x2) + snap, horizontal ? Math.round(y2) + snap : y2);
                  ctx.stroke();
                }
                ctx.setLineDash([]);
              };
              if (bd.bottom) edge(bd.bottom, x, y + h, x + w, y + h, true);
              if (bd.top) edge(bd.top, x, y, x + w, y, true);
              if (bd.left) edge(bd.left, x, y, x, y + h, false);
              if (bd.right) edge(bd.right, x + w, y, x + w, y + h, false);
            }
            if (notesRef.current.get(sheetNameOf(sheet))?.has(`${r},${c}`)) {
              ctx.fillStyle = "#d93025";
              ctx.beginPath();
              ctx.moveTo(x + w - 7, y);
              ctx.lineTo(x + w, y);
              ctx.lineTo(x + w, y + 7);
              ctx.closePath();
              ctx.fill();
            }
          }
        }
        const hasContent = (r: number, c: number) => {
          const n = tiles.cell(sheet, r, c);
          return !!n && cellText(n) !== "";
        };
        // Free space to the right/left of a cell: empty neighbours' widths.
        const freeRight = (r: number, c: number, limit: number) => {
          let extra = 0;
          for (let cc = c + 1; cc < g.colCount && extra < limit; cc++) {
            if (hasContent(r, cc) || g.mergeOf(r, cc)) break;
            extra += g.colWidth(cc);
          }
          return extra;
        };
        const freeLeft = (r: number, c: number, limit: number) => {
          let extra = 0;
          for (let cc = c - 1; cc >= 0 && extra < limit; cc--) {
            if (hasContent(r, cc) || g.mergeOf(r, cc)) break;
            extra += g.colWidth(cc);
          }
          return extra;
        };
        ctx.textBaseline = "alphabetic";
        ctx.textAlign = "left";
        for (const rc of rects) {
          const { r, c, x, y, w, h } = rc;
          const cell = tiles.cell(sheet, r, c);
          const text = cellText(cell);
          if (text === "") continue;
          if (cell?.cf?.dataBar && cell.cf.dataBar.showValue === false) continue;
          const st = withCf(styleWithOverlay(styles.get(cell?.s), overlayOf(sheet, r, c)), cell?.cf);
          const align = cellAlign(cell, st);
          ctx.font = fontString(st);
          ctx.fillStyle = st.color;
          const metrics = ctx.measureText(text);
          const tw = metrics.width;
          const pad = 2;
          // Clip region: the cell, widened into empty neighbours when the
          // text does not fit (text only — numbers never overflow).
          let clipX = x;
          let clipW = w;
          const overflow = tw + 2 * pad - w;
          const isText = typeof cell?.v === "string" && !st.wrap && !rc.merge;
          if (isText && overflow > 0) {
            if (align === "left") {
              clipW += freeRight(r, c, overflow);
            } else if (align === "right") {
              const l = freeLeft(r, c, overflow);
              clipX -= l;
              clipW += l;
            } else {
              const half = overflow / 2;
              const l = Math.min(freeLeft(r, c, half), half);
              const rr = Math.min(freeRight(r, c, half), half);
              clipX -= l;
              clipW += l + rr;
            }
          }
          let tx: number;
          if (align === "right") tx = x + w - pad - tw;
          else if (align === "center") tx = x + (w - tw) / 2;
          else tx = x + pad;
          // Pixel-snapped baseline: Univer's cells sit text on the bottom
          // (Excel's default) with a 2px inset.
          const asc = metrics.actualBoundingBoxAscent || st.sizePt;
          const desc = metrics.actualBoundingBoxDescent || st.sizePt * 0.25;
          let ty: number;
          if (st.vertical === "top") ty = y + pad + asc;
          else if (st.vertical === "middle") ty = y + (h + asc - desc) / 2;
          else ty = y + h - pad - desc;
          tx = Math.round(tx);
          ty = Math.round(ty);
          ctx.save();
          ctx.beginPath();
          ctx.rect(clipX, y, clipW, h);
          ctx.clip();
          if (st.wrap && tw > w - 2 * pad && typeof cell?.v === "string") {
            // Word-wrap inside the cell; lines stack from the top, clipped to the cell.
            const lines = wrapText(ctx, text, w - 2 * pad);
            const lineH = asc + desc + 2;
            // Text taller than the cell starts at the top, as Excel shows it.
            const fits = lines.length * lineH <= h - pad;
            let ly = !fits || st.vertical === "top" ? y + pad + asc : st.vertical === "middle" ? y + (h - lines.length * lineH) / 2 + asc : y + h - pad - desc - (lines.length - 1) * lineH;
            for (const line of lines) {
              const lw = ctx.measureText(line).width;
              const lx = align === "right" ? x + w - pad - lw : align === "center" ? x + (w - lw) / 2 : x + pad;
              ctx.fillText(line, Math.round(lx), Math.round(ly));
              ly += lineH;
            }
          } else {
            ctx.fillText(text, tx, ty);
          }
          if (st.strike || st.underline) {
            ctx.strokeStyle = st.color;
            ctx.lineWidth = 1;
            ctx.beginPath();
            if (st.strike) { const sy = Math.round(ty - asc / 2) + 0.5; ctx.moveTo(tx, sy); ctx.lineTo(tx + tw, sy); }
            if (st.underline) { const uy = Math.round(ty + 1) + 0.5; ctx.moveTo(tx, uy); ctx.lineTo(tx + tw, uy); }
            ctx.stroke();
          }
          ctx.restore();
        }
        // AutoFilter dropdown buttons on the header row.
        const fs = filtersRef.current.get(sheet);
        if (fs && fs.r0 >= rowFrom && fs.r0 < rowTo) {
          const rh = g.rowHeight(fs.r0);
          for (let c = Math.max(fs.c0, colFrom); c <= Math.min(fs.c1, colTo - 1); c++) {
            const cw = g.colWidth(c);
            if (cw < 20 || rh < 15) continue;
            const bx = Math.round(g.colLeft[c + 1] - offX) - 17, by = Math.round(g.rowTop[fs.r0 + 1] - offY) - 16;
            const active = fs.columns.has(c);
            ctx.fillStyle = active ? COLORS.selBorder : "#fff";
            ctx.strokeStyle = active ? COLORS.selBorder : "#9a9a9a";
            ctx.lineWidth = 1;
            ctx.fillRect(bx + 0.5, by + 0.5, 15, 14);
            ctx.strokeRect(bx + 0.5, by + 0.5, 15, 14);
            ctx.fillStyle = active ? "#fff" : "#444";
            ctx.beginPath();
            if (active) {
              // funnel
              ctx.moveTo(bx + 4, by + 4); ctx.lineTo(bx + 12, by + 4); ctx.lineTo(bx + 9, by + 8); ctx.lineTo(bx + 9, by + 11); ctx.lineTo(bx + 7, by + 11); ctx.lineTo(bx + 7, by + 8);
            } else {
              ctx.moveTo(bx + 4, by + 5); ctx.lineTo(bx + 12, by + 5); ctx.lineTo(bx + 8, by + 10);
            }
            ctx.closePath();
            ctx.fill();
            filterHitsRef.current.push({ x0: bx * z, y0: by * z, x1: (bx + 16) * z, y1: (by + 15) * z, col: c });
          }
        }

        // Pending agent changes, under the selection.
        const marks = highlightsRef.current.get(g.layout.name);
        if (marks && marks.size) {
          ctx.fillStyle = COLORS.pending;
          const tint = (r: number, c: number) => {
            const w = g.colLeft[c + 1] - g.colLeft[c], h = g.rowTop[r + 1] - g.rowTop[r];
            if (w > 0 && h > 0) ctx.fillRect(g.colLeft[c] - offX, g.rowTop[r] - offY, w, h);
          };
          // Walk whichever is smaller: the marks, or the visible cells.
          if (marks.size < (rowTo - rowFrom) * (colTo - colFrom)) {
            for (const k of marks) {
              const r = Math.floor(k / 1_000_000), c = k % 1_000_000;
              if (r >= rowFrom && r < rowTo && c >= colFrom && c < colTo) tint(r, c);
            }
          } else {
            for (let r = rowFrom; r < rowTo; r++) for (let c = colFrom; c < colTo; c++) if (marks.has(r * 1_000_000 + c)) tint(r, c);
          }
        }

        // Selection.
        const sr0 = Math.min(sel.r0, sel.r1), sr1 = Math.max(sel.r0, sel.r1);
        const sc0 = Math.min(sel.c0, sel.c1), sc1 = Math.max(sel.c0, sel.c1);
        if (sr1 >= rowFrom && sr0 < rowTo && sc1 >= colFrom && sc0 < colTo) {
          const x = g.colLeft[sc0] - offX;
          const y = g.rowTop[sr0] - offY;
          const w = g.colLeft[sc1 + 1] - g.colLeft[sc0];
          const h = g.rowTop[sr1 + 1] - g.rowTop[sr0];
          if (sr0 !== sr1 || sc0 !== sc1) {
            ctx.fillStyle = COLORS.selFill;
            ctx.fillRect(x, y, w, h);
          }
          ctx.strokeStyle = COLORS.selBorder;
          ctx.lineWidth = 2;
          ctx.strokeRect(x + 1, y + 1, w - 1, h - 1);
          // Fill handle.
          ctx.fillStyle = "#fff";
          ctx.fillRect(x + w - 5, y + h - 5, 7, 7);
          ctx.fillStyle = COLORS.selBorder;
          ctx.fillRect(x + w - 4, y + h - 4, 5, 5);
        }
        for (const ex of extraRanges.current) {
          if (ex.r1 < rowFrom || ex.r0 >= rowTo || ex.c1 < colFrom || ex.c0 >= colTo) continue;
          const x = g.colLeft[ex.c0] - offX, y = g.rowTop[ex.r0] - offY;
          const w = g.colLeft[Math.min(ex.c1 + 1, g.colCount)] - g.colLeft[ex.c0], h = g.rowTop[Math.min(ex.r1 + 1, g.rowCount)] - g.rowTop[ex.r0];
          ctx.fillStyle = COLORS.selFill;
          ctx.fillRect(x, y, w, h);
          ctx.strokeStyle = COLORS.selBorder;
          ctx.lineWidth = 1;
          ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
        }
        const fl = filling.current;
        if (fl) {
          const fr0 = Math.min(fl.r0, fl.toR), fr1 = Math.max(fl.r1, fl.toR), fc0 = Math.min(fl.c0, fl.toC), fc1 = Math.max(fl.c1, fl.toC);
          if (fr1 >= rowFrom && fr0 < rowTo && fc1 >= colFrom && fc0 < colTo) {
            ctx.strokeStyle = "#555";
            ctx.lineWidth = 1;
            ctx.setLineDash([3, 2]);
            ctx.strokeRect(g.colLeft[fc0] - offX + 0.5, g.rowTop[fr0] - offY + 0.5, g.colLeft[fc1 + 1] - g.colLeft[fc0], g.rowTop[fr1 + 1] - g.rowTop[fr0]);
            ctx.setLineDash([]);
          }
        }
        // Ranges referenced by the formula being edited.
        editRefsRef.current.forEach((ref, i) => {
          if (ref.r1 < rowFrom || ref.r0 >= rowTo || ref.c1 < colFrom || ref.c0 >= colTo) return;
          const x = g.colLeft[Math.min(ref.c0, g.colCount - 1)] - offX;
          const y = g.rowTop[Math.min(ref.r0, g.rowCount - 1)] - offY;
          const w = g.colLeft[Math.min(ref.c1 + 1, g.colCount)] - g.colLeft[Math.min(ref.c0, g.colCount - 1)];
          const h = g.rowTop[Math.min(ref.r1 + 1, g.rowCount)] - g.rowTop[Math.min(ref.r0, g.rowCount - 1)];
          ctx.strokeStyle = REF_COLORS[i % REF_COLORS.length];
          ctx.lineWidth = 2;
          ctx.setLineDash([4, 2]);
          ctx.strokeRect(x + 1, y + 1, w - 1, h - 1);
          ctx.setLineDash([]);
        });
        ctx.restore();
      };

      const vis = visibleRect()!;
      const bodyX = HEADER_W;
      const bodyY = HEADER_H;
      // Scrolling pane (bottom-right), frozen rows (top-right), frozen cols
      // (bottom-left), corner (top-left). At scroll 0 the frozen band and
      // the scrolling pane are contiguous, so they draw as one pane and
      // text may overflow across the split, as in Excel.
      const rows0 = scrollY === 0 ? 0 : g.frozenRows;
      const cols0 = scrollX === 0 ? 0 : g.frozenCols;
      const rowEnd = Math.min(g.rowCount, vis.r1 + 2);
      const colEnd = Math.min(g.colCount, vis.c1 + 2);
      const rowStart = Math.max(rows0, vis.r0);
      const colStart = Math.max(cols0, vis.c0);
      const pfx = cols0 > 0 ? fx : 0;
      const pfy = rows0 > 0 ? fy : 0;
      pane(rowStart, rowEnd, colStart, colEnd, scrollX - bodyX, scrollY - bodyY, bodyX + pfx, bodyY + pfy, W - bodyX - pfx, H - bodyY - pfy);
      if (rows0 > 0) pane(0, rows0, colStart, colEnd, scrollX - bodyX, -bodyY, bodyX + pfx, bodyY, W - bodyX - pfx, fy);
      if (cols0 > 0) pane(rowStart, rowEnd, 0, cols0, -bodyX, scrollY - bodyY, bodyX, bodyY + pfy, fx, H - bodyY - pfy);
      if (rows0 > 0 && cols0 > 0) pane(0, rows0, 0, cols0, -bodyX, -bodyY, bodyX, bodyY, fx, fy);
      // Frozen divider lines.
      if (g.frozenRows > 0 || g.frozenCols > 0) {
        ctx.strokeStyle = COLORS.frozen;
        ctx.lineWidth = 1;
        ctx.beginPath();
        if (g.frozenRows > 0) { ctx.moveTo(0, bodyY + fy + 0.5); ctx.lineTo(W, bodyY + fy + 0.5); }
        if (g.frozenCols > 0) { ctx.moveTo(bodyX + fx + 0.5, 0); ctx.lineTo(bodyX + fx + 0.5, H); }
        ctx.stroke();
      }

      // Headers.
      ctx.fillStyle = COLORS.header;
      ctx.fillRect(0, 0, W, HEADER_H);
      ctx.fillRect(0, 0, HEADER_W, H);
      ctx.font = "12px Arial, sans-serif";
      ctx.textBaseline = "middle";
      ctx.textAlign = "center";
      const selR0 = Math.min(sel.r0, sel.r1), selR1 = Math.max(sel.r0, sel.r1);
      const selC0 = Math.min(sel.c0, sel.c1), selC1 = Math.max(sel.c0, sel.c1);
      const headerCols = (from: number, to: number, offX: number, clipX: number, clipW: number) => {
        ctx.save();
        ctx.beginPath();
        ctx.rect(clipX, 0, clipW, HEADER_H);
        ctx.clip();
        for (let c = from; c < to; c++) {
          const cw = g.colWidth(c);
          if (cw === 0) continue;
          const x = g.colLeft[c] - offX;
          if (c >= selC0 && c <= selC1) {
            ctx.fillStyle = COLORS.headerSel;
            ctx.fillRect(x, 0, cw, HEADER_H);
          }
          ctx.strokeStyle = COLORS.grid;
          ctx.beginPath();
          ctx.moveTo(Math.round(x + cw) + 0.5, 0);
          ctx.lineTo(Math.round(x + cw) + 0.5, HEADER_H);
          ctx.stroke();
          ctx.fillStyle = COLORS.headerText;
          if (cw > 14) ctx.fillText(colLetters(c), x + cw / 2, HEADER_H / 2 + 0.5);
        }
        ctx.restore();
      };
      const headerRows = (from: number, to: number, offY: number, clipY: number, clipH: number) => {
        ctx.save();
        ctx.beginPath();
        ctx.rect(0, clipY, HEADER_W, clipH);
        ctx.clip();
        for (let r = from; r < to; r++) {
          const rh = g.rowHeight(r);
          if (rh === 0) continue;
          const y = g.rowTop[r] - offY;
          if (r >= selR0 && r <= selR1) {
            ctx.fillStyle = COLORS.headerSel;
            ctx.fillRect(0, y, HEADER_W, rh);
          }
          ctx.strokeStyle = COLORS.grid;
          ctx.beginPath();
          ctx.moveTo(0, Math.round(y + rh) + 0.5);
          ctx.lineTo(HEADER_W, Math.round(y + rh) + 0.5);
          ctx.stroke();
          ctx.fillStyle = COLORS.headerText;
          if (rh > 10) ctx.fillText(String(r + 1), HEADER_W / 2, y + rh / 2 + 0.5);
        }
        ctx.restore();
      };
      const hfx = cols0 > 0 ? fx : 0;
      const hfy = rows0 > 0 ? fy : 0;
      headerCols(colStart, colEnd, scrollX - bodyX, bodyX + hfx, W - bodyX - hfx);
      if (cols0 > 0) headerCols(0, cols0, -bodyX, bodyX, fx);
      headerRows(rowStart, rowEnd, scrollY - bodyY, bodyY + hfy, H - bodyY - hfy);
      if (rows0 > 0) headerRows(0, rows0, -bodyY, bodyY, fy);
      ctx.strokeStyle = "#cfcfcf";
      ctx.beginPath();
      ctx.moveTo(0, HEADER_H + 0.5);
      ctx.lineTo(W, HEADER_H + 0.5);
      ctx.moveTo(HEADER_W + 0.5, 0);
      ctx.lineTo(HEADER_W + 0.5, H);
      ctx.stroke();

      drawOutline(g, scrollX, scrollY, z, dpr, rows0 > 0 ? fy : 0, cols0 > 0 ? fx : 0);

      // Kick fetches for anything we just drew blank.
      requestTiles();
    };

    /** Group gutters: level buttons in the corner, a bar per group with a
     * +/- button at its summary row/column (Excel's outline symbols). */
    const drawOutline = (g: SheetGeometry, scrollX: number, scrollY: number, z: number, dpr: number, hfy: number, hfx: number) => {
      const area = gridAreaRef.current;
      const oc = outlineCanvasRef.current;
      const gwS = outlineGutter(g.maxRowLevel), ghS = outlineGutter(g.maxColLevel);
      const gw = Math.round(gwS * z), gh = Math.round(ghS * z);
      if (gw !== gutterRef.current.w || gh !== gutterRef.current.h) {
        gutterRef.current = { w: gw, h: gh };
        setGutter(gutterRef.current);
      }
      outlineHitsRef.current = [];
      if (!area || !oc) return;
      const cssW = area.clientWidth, cssH = area.clientHeight;
      if (gw === 0 && gh === 0) { if (oc.width !== 0) { oc.width = 0; oc.height = 0; } return; }
      if (oc.width !== Math.round(cssW * dpr) || oc.height !== Math.round(cssH * dpr)) {
        oc.width = Math.round(cssW * dpr); oc.height = Math.round(cssH * dpr);
        oc.style.width = `${cssW}px`; oc.style.height = `${cssH}px`;
      }
      const ctx = oc.getContext("2d");
      if (!ctx) return;
      const W = cssW / z, H = cssH / z;
      ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
      ctx.clearRect(0, 0, W, H);
      ctx.fillStyle = COLORS.header;
      if (gw) ctx.fillRect(0, 0, gwS, H);
      if (gh) ctx.fillRect(0, 0, W, ghS);
      ctx.strokeStyle = "#cfcfcf";
      ctx.beginPath();
      if (gw) { ctx.moveTo(gwS - 0.5, 0); ctx.lineTo(gwS - 0.5, H); }
      if (gh) { ctx.moveTo(0, ghS - 0.5); ctx.lineTo(W, ghS - 0.5); }
      ctx.stroke();
      const sheet = activeRef.current;
      const hits = outlineHitsRef.current;
      const hit = (x: number, y: number, w: number, h: number, run: () => void) => hits.push({ x0: x * z, y0: y * z, x1: (x + w) * z, y1: (y + h) * z, run });
      const box = (cx: number, cy: number, glyph: "+" | "-" | string) => {
        ctx.fillStyle = "#fff";
        ctx.strokeStyle = "#8c8c8c";
        ctx.lineWidth = 1;
        ctx.fillRect(Math.round(cx - 5) + 0.5, Math.round(cy - 5) + 0.5, 10, 10);
        ctx.strokeRect(Math.round(cx - 5) + 0.5, Math.round(cy - 5) + 0.5, 10, 10);
        ctx.fillStyle = "#333";
        ctx.font = glyph.length === 1 && (glyph === "+" || glyph === "-") ? "bold 10px Arial, sans-serif" : "9px Arial, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(glyph === "-" ? "−" : glyph, Math.round(cx), Math.round(cy) + 0.5);
      };
      // Row positions in this canvas (sheet px): frozen rows never scroll.
      const rowY = (r: number) => ghS + HEADER_H + (scrollY > 0 && r >= g.frozenRows ? g.rowTop[r] - scrollY : g.rowTop[r]);
      const colX = (c: number) => gwS + HEADER_W + (scrollX > 0 && c >= g.frozenCols ? g.colLeft[c] - scrollX : g.colLeft[c]);
      const setRows = (from: number, to: number, hidden: boolean, label: string) => {
        withUndoGroup(label, () => { for (let r = from; r <= to; r++) setSize(sheet, "row", r, hidden ? 0 : g.naturalRowPx(r), true); });
        onUserEdit?.();
      };
      const setCols = (from: number, to: number, hidden: boolean, label: string) => {
        withUndoGroup(label, () => { for (let c = from; c <= to; c++) setSize(sheet, "col", c, hidden ? 0 : g.naturalColPx(c), true); });
        onUserEdit?.();
      };
      const showLevel = (axis: "row" | "col", n: number) => {
        const levels = axis === "row" ? g.rowLevel : g.colLevel;
        withUndoGroup(`outline level ${n}`, () => {
          for (let i = 0; i < levels.length; i++) {
            const l = levels[i];
            if (l === 0) continue;
            const hidden = l >= n;
            const cur = axis === "row" ? g.rowHeight(i) : g.colWidth(i);
            if (hidden !== (cur === 0)) setSize(sheet, axis, i, hidden ? 0 : axis === "row" ? g.naturalRowPx(i) : g.naturalColPx(i), true);
          }
        });
        onUserEdit?.();
      };
      ctx.lineWidth = 1;
      if (gw) {
        // Level buttons across the top of the row gutter.
        for (let n = 1; n <= g.maxRowLevel + 1; n++) {
          const cx = (n - 1) * OUTLINE_STEP + OUTLINE_STEP / 2 + 2, cy = ghS + HEADER_H / 2;
          box(cx, cy, String(n));
          hit(cx - 6, cy - 6, 12, 12, () => showLevel("row", n));
        }
        ctx.save();
        ctx.beginPath();
        ctx.rect(0, ghS + HEADER_H, gwS, H);
        ctx.clip();
        for (let L = 1; L <= g.maxRowLevel; L++) {
          const x = (L - 1) * OUTLINE_STEP + OUTLINE_STEP / 2 + 2;
          for (const run of outlineGroups(g.rowLevel, L)) {
            const summary = g.summaryBelow ? run.end + 1 : run.start - 1;
            const collapsed = g.rowTop[run.end + 1] === g.rowTop[run.start];
            const clipTop = ghS + HEADER_H + (run.start >= g.frozenRows ? hfy : 0);
            ctx.save();
            ctx.beginPath();
            ctx.rect(0, clipTop, gwS, H);
            ctx.clip();
            if (!collapsed) {
              const y0 = rowY(run.start), y1 = rowY(run.end + 1);
              if (y1 >= clipTop && y0 <= H) {
                ctx.strokeStyle = "#555";
                ctx.beginPath();
                ctx.moveTo(Math.round(x) + 0.5, y0 + 1);
                ctx.lineTo(Math.round(x) + 0.5, y1 - 1);
                // tick toward the summary row
                const ty = g.summaryBelow ? y1 - 1 : y0 + 1;
                ctx.moveTo(Math.round(x) + 0.5, Math.round(ty) + 0.5);
                ctx.lineTo(Math.round(x) + 5, Math.round(ty) + 0.5);
                ctx.stroke();
              }
            }
            if (summary >= 0 && summary < g.rowCount && g.rowHeight(summary) > 0) {
              const cy = rowY(summary) + g.rowHeight(summary) / 2;
              if (cy >= clipTop && cy <= H) {
                box(x, cy, collapsed ? "+" : "-");
                hit(x - 6, cy - 6, 12, 12, () => setRows(run.start, run.end, !collapsed, collapsed ? "expand group" : "collapse group"));
              }
            }
            ctx.restore();
          }
        }
        ctx.restore();
      }
      if (gh) {
        for (let n = 1; n <= g.maxColLevel + 1; n++) {
          const cy = (n - 1) * OUTLINE_STEP + OUTLINE_STEP / 2 + 2, cx = gwS + HEADER_W / 2;
          box(cx, cy, String(n));
          hit(cx - 6, cy - 6, 12, 12, () => showLevel("col", n));
        }
        ctx.save();
        ctx.beginPath();
        ctx.rect(gwS + HEADER_W, 0, W, ghS);
        ctx.clip();
        for (let L = 1; L <= g.maxColLevel; L++) {
          const y = (L - 1) * OUTLINE_STEP + OUTLINE_STEP / 2 + 2;
          for (const run of outlineGroups(g.colLevel, L)) {
            const summary = g.summaryRight ? run.end + 1 : run.start - 1;
            const collapsed = g.colLeft[run.end + 1] === g.colLeft[run.start];
            const clipLeft = gwS + HEADER_W + (run.start >= g.frozenCols ? hfx : 0);
            ctx.save();
            ctx.beginPath();
            ctx.rect(clipLeft, 0, W, ghS);
            ctx.clip();
            if (!collapsed) {
              const x0 = colX(run.start), x1 = colX(run.end + 1);
              if (x1 >= clipLeft && x0 <= W) {
                ctx.strokeStyle = "#555";
                ctx.beginPath();
                ctx.moveTo(x0 + 1, Math.round(y) + 0.5);
                ctx.lineTo(x1 - 1, Math.round(y) + 0.5);
                const tx = g.summaryRight ? x1 - 1 : x0 + 1;
                ctx.moveTo(Math.round(tx) + 0.5, Math.round(y) + 0.5);
                ctx.lineTo(Math.round(tx) + 0.5, Math.round(y) + 5);
                ctx.stroke();
              }
            }
            if (summary >= 0 && summary < g.colCount && g.colWidth(summary) > 0) {
              const cx = colX(summary) + g.colWidth(summary) / 2;
              if (cx >= clipLeft && cx <= W) {
                box(cx, y, collapsed ? "+" : "-");
                hit(cx - 6, y - 6, 12, 12, () => setCols(run.start, run.end, !collapsed, collapsed ? "expand group" : "collapse group"));
              }
            }
            ctx.restore();
          }
        }
        ctx.restore();
      }
    };

    // ----- AutoFilter -------------------------------------------------------

    /** Session filter state from the file's AutoFilters (fresh open only). */
    const initFilters = (layout: EngineLayout) => {
      filtersRef.current.clear();
      autoFilterEditsRef.current.clear();
      layout.sheets.forEach((l, i) => {
        const af = l.autoFilter;
        if (!af) return;
        const columns = new Map<number, ColFilter>();
        for (const c of af.columns) {
          const f: ColFilter = {};
          if (c.values) { f.values = new Set(c.values); f.blank = c.blank; }
          if (c.custom) f.custom = { and: c.custom.and, items: c.custom.items.map((it) => ({ op: it.op, val: it.val })) };
          if (c.other) f.other = true;
          if (f.values || f.custom || f.other) columns.set(af.c0 - 1 + c.col, f);
        }
        filtersRef.current.set(i, { r0: af.r0 - 1, c0: af.c0 - 1, r1: af.r1 - 1, c1: af.c1 - 1, columns });
      });
    };

    const filterRangeA1 = (fs: FilterState) => `${colLetters(fs.c0)}${fs.r0 + 1}:${colLetters(fs.c1)}${fs.r1 + 1}`;

    /** Queue the sheet's filter (range + criteria) for the save mirror. */
    const recordFilterEdit = (sheet: number) => {
      const name = sheetNameOf(sheet);
      if (!name) return;
      const fs = filtersRef.current.get(sheet);
      if (!fs) { autoFilterEditsRef.current.set(name, { range: null }); return; }
      const columns: AutoFilterColumn[] = [];
      for (const [c, f] of fs.columns) {
        if (f.other) continue;
        const col: AutoFilterColumn = { colId: c - fs.c0 };
        if (f.values) { col.values = [...f.values]; if (f.blank) col.blank = true; }
        if (f.custom) col.custom = { and: f.custom.and, items: f.custom.items.map((i) => ({ operator: i.op, val: i.val })) };
        columns.push(col);
      }
      autoFilterEditsRef.current.set(name, { range: filterRangeA1(fs), columns });
    };

    const cellNum = (cell: GridCell | undefined) => (cell && typeof cell.v === "number" ? cell.v : NaN);

    const rowPassesFilters = (sheet: number, fs: FilterState, r: number, skipCol = -1) => {
      for (const [c, f] of fs.columns) {
        if (c === skipCol) continue;
        const cell = tilesRef.current.cell(sheet, r, c);
        if (!colFilterMatches(f, cellText(cell), cellNum(cell))) return false;
      }
      return true;
    };

    /** Set (or clear with null) one column's criteria, then hide/show the
     * range's rows to match every column's filter. Undoable as one step. */
    const applyFilter = async (sheet: number, col: number, next: ColFilter | null) => {
      const fs = filtersRef.current.get(sheet);
      if (!fs) return;
      const prev = fs.columns.get(col);
      const setCol = (f: ColFilter | undefined) => { if (f) fs.columns.set(col, f); else fs.columns.delete(col); };
      for (const c of fs.columns.keys()) await prefetchRef.current(sheet, fs.r0 + 1, fs.r1, c, c);
      if (next) await prefetchRef.current(sheet, fs.r0 + 1, fs.r1, col, col);
      setCol(next ?? undefined);
      const g0 = geom(sheet);
      if (!g0) return;
      withUndoGroup("filter", () => {
        pushUndo({ label: "filter", undo: () => { setCol(prev); recordFilterEdit(sheet); scheduleDraw(); }, redo: () => { setCol(next ?? undefined); recordFilterEdit(sheet); scheduleDraw(); } });
        for (let r = fs.r0 + 1; r <= fs.r1; r++) {
          const show = rowPassesFilters(sheet, fs, r);
          const cur = g0.rowHeight(r);
          if (show && cur === 0) setSize(sheet, "row", r, g0.naturalRowPx(r), true);
          else if (!show && cur !== 0) setSize(sheet, "row", r, 0, true);
        }
      });
      recordFilterEdit(sheet);
      onUserEdit?.();
      scheduleDraw();
    };

    const openFilterMenu = (col: number, x: number, y: number) => {
      const sheet = activeRef.current;
      const fs = filtersRef.current.get(sheet);
      if (!fs) return;
      setFilterMenu({ sheet, col, x, y, search: "", items: [], loading: true });
      void (async () => {
        await prefetchRef.current(sheet, fs.r0 + 1, fs.r1, col, col);
        for (const c of fs.columns.keys()) if (c !== col) await prefetchRef.current(sheet, fs.r0 + 1, fs.r1, c, c);
        const cur = fs.columns.get(col);
        const seen = new Map<string, number>();
        for (let r = fs.r0 + 1; r <= fs.r1; r++) {
          if (!rowPassesFilters(sheet, fs, r, col)) continue; // Excel lists what the other filters leave
          const cell = tilesRef.current.cell(sheet, r, col);
          const t = cellText(cell);
          if (!seen.has(t)) seen.set(t, cellNum(cell));
        }
        const items = [...seen.keys()].sort(filterValueCompare).map((text) => ({ text, checked: cur ? colFilterMatches(cur, text, seen.get(text) ?? NaN) : true }));
        setFilterMenu((m) => (m && m.col === col && m.sheet === sheet ? { ...m, items, loading: false } : m));
      })();
    };

    /** Excel's current region: grow from the cell while the bordering rows/
     * columns have content (from loaded tiles; the screen is always loaded). */
    const currentRegion = (sheet: number, r: number, c: number, g: SheetGeometry) => {
      const has = (rr: number, cc: number) => cellText(tilesRef.current.cell(sheet, rr, cc)) !== "";
      let r0 = r, r1 = r, c0 = c, c1 = c;
      for (let n = 0; n < 20_000; n++) {
        let grew = false;
        const rowHas = (rr: number) => { for (let cc = c0; cc <= c1; cc++) if (has(rr, cc)) return true; return false; };
        const colHas = (cc: number) => { for (let rr = r0; rr <= r1; rr++) if (has(rr, cc)) return true; return false; };
        if (r0 > 0 && rowHas(r0 - 1)) { r0--; grew = true; }
        if (r1 < g.rowCount - 1 && rowHas(r1 + 1)) { r1++; grew = true; }
        if (c0 > 0 && colHas(c0 - 1)) { c0--; grew = true; }
        if (c1 < g.colCount - 1 && colHas(c1 + 1)) { c1++; grew = true; }
        if (!grew) break;
      }
      return { r0, c0, r1, c1 };
    };

    /** Excel's Filter toggle: remove the sheet's AutoFilter (showing its
     * rows again) or create one over the selection / current region. */
    const toggleFilterRange = (sheet: number, sr: number, sc: number, er: number, ec: number): boolean => {
      const g = geom(sheet);
      if (!g) return false;
      const existing = filtersRef.current.get(sheet);
      if (existing) {
        const snapshot = existing;
        withUndoGroup("remove filter", () => {
          pushUndo({ label: "remove filter", undo: () => { filtersRef.current.set(sheet, snapshot); recordFilterEdit(sheet); scheduleDraw(); }, redo: () => { filtersRef.current.delete(sheet); recordFilterEdit(sheet); scheduleDraw(); } });
          for (let r = existing.r0 + 1; r <= existing.r1; r++) if (g.rowHeight(r) === 0) setSize(sheet, "row", r, g.naturalRowPx(r), true);
        });
        filtersRef.current.delete(sheet);
        recordFilterEdit(sheet);
        onUserEdit?.();
        scheduleDraw();
        return false;
      }
      const single = sr === er && sc === ec;
      const rect = single ? currentRegion(sheet, sr, sc, g) : { r0: Math.min(sr, er), c0: Math.min(sc, ec), r1: Math.max(sr, er), c1: Math.max(sc, ec) };
      const fs: FilterState = { ...rect, columns: new Map() };
      filtersRef.current.set(sheet, fs);
      pushUndo({ label: "filter", undo: () => { filtersRef.current.delete(sheet); recordFilterEdit(sheet); scheduleDraw(); }, redo: () => { filtersRef.current.set(sheet, fs); recordFilterEdit(sheet); scheduleDraw(); } });
      recordFilterEdit(sheet);
      onUserEdit?.();
      scheduleDraw();
      if (single) {
        // The region may continue past the loaded tiles: extend it once more
        // after fetching the rows below.
        void (async () => {
          const l = layoutRef.current?.sheets[sheet];
          const bottom = Math.min(g.rowCount - 1, Math.max(fs.r1, (l?.usedRow ?? 1) - 1));
          if ((bottom - fs.r1) * (fs.c1 - fs.c0 + 1) > 400_000) return;
          await prefetchRef.current(sheet, fs.r1, bottom, fs.c0, fs.c1);
          const again = currentRegion(sheet, sr, sc, g);
          if (again.r1 !== fs.r1 || again.c1 !== fs.c1 || again.r0 !== fs.r0 || again.c0 !== fs.c0) { Object.assign(fs, again); recordFilterEdit(sheet); scheduleDraw(); }
        })();
      }
      return true;
    };

    const toggleFilterAtSelection = () => {
      const s = selRef.current;
      toggleFilterRange(activeRef.current, Math.min(s.r0, s.r1), Math.min(s.c0, s.c1), Math.max(s.r0, s.r1), Math.max(s.c0, s.c1));
    };

    const onOutlineMouseDown = (e: React.MouseEvent) => {
      const area = gridAreaRef.current;
      if (!area) return;
      const rect = area.getBoundingClientRect();
      const x = e.clientX - rect.left, y = e.clientY - rect.top;
      const h = outlineHitsRef.current.find((t) => x >= t.x0 && x <= t.x1 && y >= t.y0 && y <= t.y1);
      e.preventDefault();
      e.stopPropagation();
      if (h) h.run();
      scrollerRef.current?.focus();
    };

    // ----- formula references while editing --------------------------------

    const REF_COLORS = ["#1a73e8", "#e8710a", "#9334e6", "#188038", "#d93025", "#007b83"];

    /** A1 references in a formula that point at the active sheet. */
    const parseRefs = (text: string): Array<{ r0: number; c0: number; r1: number; c1: number }> => {
      const out: Array<{ r0: number; c0: number; r1: number; c1: number }> = [];
      if (!text.startsWith("=")) return out;
      const active = layoutRef.current?.sheets[activeRef.current]?.name;
      const re = /(?:'((?:[^']|'')+)'!|([A-Za-z0-9_.]+)!)?\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?/g;
      let m: RegExpExecArray | null;
      while ((m = re.exec(text)) && out.length < 64) {
        const sheet = m[1]?.replace(/''/g, "'") ?? m[2];
        if (sheet && sheet !== active) continue;
        // Not part of a longer identifier (e.g. LOG10, a name like TAX2024).
        const before = text[m.index - 1];
        if (before && /[A-Za-z0-9_.]/.test(before)) continue;
        const a = parseA1(`${m[3]}${m[4]}`);
        const b = m[5] ? parseA1(`${m[5]}${m[6]}`) : a;
        if (!a || !b) continue;
        out.push({ r0: Math.min(a.row0, b.row0), c0: Math.min(a.col0, b.col0), r1: Math.max(a.row0, b.row0), c1: Math.max(a.col0, b.col0) });
      }
      return out;
    };

    useEffect(() => {
      measureTabs();
      const ro = tabsRef.current ? new ResizeObserver(measureTabs) : null;
      if (ro && tabsRef.current) ro.observe(tabsRef.current);
      return () => ro?.disconnect();
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [sheetNames, showHidden]);

    useEffect(() => {
      editRefsRef.current = editing ? parseRefs(editing.text) : [];
      scheduleDraw();
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [editing?.text, editing === null]);

    /** Can a click insert a reference at the caret? After an operator, an
     * opening paren, a separator, or right after a reference we inserted. */
    const refInsertable = (text: string, caret: number): boolean => {
      if (!text.startsWith("=")) return false;
      if (refPick.current && caret === refPick.current.end) return true;
      const head = text.slice(0, caret).replace(/\s+$/, "");
      return /[=(,;+\-*\/^&<>:]$/.test(head);
    };

    const activeInput = () => (editing?.origin === "bar" ? barRef.current : editorRef.current);

    const refText = (a: { r: number; c: number }, b: { r: number; c: number }) => {
      const r0 = Math.min(a.r, b.r), r1 = Math.max(a.r, b.r), c0 = Math.min(a.c, b.c), c1 = Math.max(a.c, b.c);
      const one = `${colLetters(c0)}${r0 + 1}`;
      return r0 === r1 && c0 === c1 ? one : `${one}:${colLetters(c1)}${r1 + 1}`;
    };

    /** Insert or replace the picked reference in the editor text. */
    const putRef = (hit: { r: number; c: number }, extendFrom?: { r: number; c: number }) => {
      if (!editing) return;
      const input = activeInput();
      const text = editing.text;
      const pick = refPick.current;
      const start = pick ? pick.start : Math.min(input?.selectionStart ?? text.length, text.length);
      const end = pick ? pick.end : start;
      const anchor = extendFrom ?? pick?.anchor ?? hit;
      const ref = refText(anchor, hit);
      const next = text.slice(0, start) + ref + text.slice(end);
      refPick.current = { start, end: start + ref.length, anchor };
      setEditing({ ...editing, text: next });
      requestAnimationFrame(() => {
        const el = activeInput();
        if (el) { el.focus(); el.setSelectionRange(start + ref.length, start + ref.length); }
      });
    };

    // ----- selection + formula bar --------------------------------------

    const refreshFormulaBar = useCallback(() => {
      const s = selRef.current;
      setNameBox(`${colLetters(s.ac)}${s.ar + 1}`);
      const cell = tilesRef.current.cell(activeRef.current, s.ar, s.ac);
      setFormulaText(cell?.f ?? (cell ? rawInput(cell) : ""));
    }, []);

    const setSelection = (r: number, c: number, extend = false) => {
      const g = geom(activeRef.current);
      if (!g) return;
      r = Math.max(0, Math.min(g.rowCount - 1, r));
      c = Math.max(0, Math.min(g.colCount - 1, c));
      const s = selRef.current;
      selRef.current = extend ? { ...s, r1: r, c1: c } : { r0: r, c0: c, r1: r, c1: c, ar: r, ac: c };
      refreshFormulaBar();
      setStats(selectionStats());
      scheduleDraw();
    };

    const scrollIntoView = (r: number, c: number) => {
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!sc || !g) return;
      r = Math.max(0, Math.min(g.rowCount - 1, r));
      c = Math.max(0, Math.min(g.colCount - 1, c));
      const x0 = g.colLeft[c], x1 = g.colLeft[c + 1];
      const y0 = g.rowTop[r], y1 = g.rowTop[r + 1];
      const fx = g.colLeft[g.frozenCols], fy = g.rowTop[g.frozenRows];
      const z = viewScale();
      const viewW = sc.clientWidth / z - HEADER_W - fx;
      const viewH = sc.clientHeight / z - HEADER_H - fy;
      const sl = sc.scrollLeft / z, st = sc.scrollTop / z;
      if (c >= g.frozenCols) {
        if (x0 - fx < sl) sc.scrollLeft = Math.max(0, x0 - fx) * z;
        else if (x1 - fx > sl + viewW) sc.scrollLeft = (x1 - fx - viewW) * z;
      }
      if (r >= g.frozenRows) {
        if (y0 - fy < st) sc.scrollTop = Math.max(0, y0 - fy) * z;
        else if (y1 - fy > st + viewH) sc.scrollTop = (y1 - fy - viewH) * z;
      }
    };

    const cellAtPoint = (px: number, py: number): { r: number; c: number } | null => {
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!sc || !g) return null;
      const rect = sc.getBoundingClientRect();
      const z = viewScale();
      const x = (px - rect.left) / z;
      const y = (py - rect.top) / z;
      if (x < HEADER_W || y < HEADER_H) return null;
      const fx = g.colLeft[g.frozenCols], fy = g.rowTop[g.frozenRows];
      const sheetX = x - HEADER_W < fx ? x - HEADER_W : x - HEADER_W + Math.round(sc.scrollLeft) / z;
      const sheetY = y - HEADER_H < fy ? y - HEADER_H : y - HEADER_H + Math.round(sc.scrollTop) / z;
      return { r: g.rowAt(sheetY), c: g.colAt(sheetX) };
    };

    // ----- editing ---------------------------------------------------------

    /** The edit in progress, mirrored in a ref: the editor's blur fires after
     * a commit has already unmounted it, with a stale closure, and must not
     * commit (and re-select) a second time. */
    const editingRef = useRef<Editing | null>(null);
    editingRef.current = editing;
    const commitEdit = (text: string, move: "down" | "right" | "up" | "left" | "none") => {
      const e = editingRef.current;
      editingRef.current = null;
      refPick.current = null;
      refDrag.current = false;
      setEditing(null);
      const sheet = activeRef.current;
      const name = layoutRef.current?.sheets[sheet]?.name;
      if (!e || !name) return;
      const input = text;
      writeCells(name, [{ row: e.row, col: e.col, value: input === "" ? null : input }], "typed");
      onUserEdit?.();
      if (move === "down") setSelection(e.row + 1, e.col);
      else if (move === "right") setSelection(e.row, e.col + 1);
      else if (move === "up") setSelection(e.row - 1, e.col);
      else if (move === "left") setSelection(e.row, e.col - 1);
      else setSelection(e.row, e.col);
      scrollerRef.current?.focus();
    };

    const startEdit = (initial?: string) => {
      const s = selRef.current;
      const cell = tilesRef.current.cell(activeRef.current, s.ar, s.ac);
      const text = initial ?? (cell?.f ?? (cell ? rawInput(cell) : ""));
      scrollIntoView(s.ar, s.ac);
      setEditing({ row: s.ar, col: s.ac, text, origin: "cell" });
    };

    const markDirty = (sheet: number, row: number, col: number) => {
      let set = dirtyRef.current.get(sheet);
      if (!set) dirtyRef.current.set(sheet, (set = new Set()));
      if (set.size < 50_000) set.add(row * 1_000_000 + col);
    };

    /** The agent-facing format of a cell: engine style plus the session overlay. */
    const formatOf = (styleId: number | undefined, overlay: CellFormatShape | undefined): CellFormatShape => {
      const st = stylesRef.current.get(styleId);
      const raw = stylesRef.current.rawOf(styleId);
      const f: CellFormatShape = {};
      if (st.bold) f.bold = true;
      if (st.italic) f.italic = true;
      if (st.underline) f.underline = true;
      if (st.strike) f.strike = true;
      // The same facts the Univer grid reported for a styled cell (font
      // family, size and colour whenever the file's style carries them), so
      // the agent's preview bytes match across grids.
      const rawFont = (raw as any)?.font as { name?: string; sz?: number; color?: string } | undefined;
      if (rawFont?.color) f.font_color = rawFont.color;
      if (st.fill) f.background_color = st.fill;
      if (rawFont?.sz) f.font_size = rawFont.sz;
      if (rawFont?.name) f.font_family = rawFont.name;
      if (st.align) f.horizontal_align = st.align;
      if (st.vertical !== "bottom") f.vertical_align = st.vertical;
      if (st.wrap) f.wrap_text = true;
      const nf = (raw as any)?.num_fmt;
      if (typeof nf === "string" && nf.toLowerCase() !== "general") f.number_format = nf;
      return { ...f, ...(overlay ?? {}) };
    };

    /** Shared write path: engine (journaled, coalesced), local tile, edit log. */
    const writeCells = (
      sheetName: string,
      cells: Array<{ row: number; col: number; value: string | number | boolean | null }>,
      label: string,
      history: "record" | "skip" = "record",
    ): boolean => {
      // A sheet the layout does not know yet (created a moment ago, op still
      // queued) is fine: the engine mirror resolves names in-queue and adds
      // missing sheets itself. Only the local tile update needs the index.
      const sheet = sheetIndexOf(sheetName);
      const undoCells: Array<{ row: number; col: number; before: string; after: string }> = [];
      const wire = cells.map((c) => {
        const input = c.value === null || c.value === undefined ? "" : typeof c.value === "boolean" ? (c.value ? "TRUE" : "FALSE") : String(c.value);
        if (sheet >= 0) {
          markDirty(sheet, c.row, c.col);
          // Writes past the used range extend it (Ctrl+End, previews, autofit).
          const l = layoutRef.current?.sheets[sheet];
          if (l && input !== "") {
            if (c.row + 1 > l.maxRow) { l.maxRow = c.row + 1; geomsRef.current.delete(sheet); }
            if (c.col + 1 > l.maxCol) { l.maxCol = c.col + 1; geomsRef.current.delete(sheet); }
            if (c.row + 1 > (l.usedRow ?? 0)) l.usedRow = c.row + 1;
            if (c.col + 1 > (l.usedCol ?? 0)) l.usedCol = c.col + 1;
          }
          const prev = tilesRef.current.cell(sheet, c.row, c.col);
          // Only cells in loaded tiles can be undone (we know what they held).
          if (history === "record" && (prev || tilesRef.current.has(tileKey(sheet, Math.floor(c.row / TILE_ROWS), Math.floor(c.col / TILE_COLS)))))
            undoCells.push({ row: c.row, col: c.col, before: prev?.f ?? (prev ? rawInput(prev) : ""), after: input });
          tilesRef.current.setLocal(sheet, c.row, c.col, input);
        }
        let perSheet = editsRef.current.get(sheetName);
        if (!perSheet) editsRef.current.set(sheetName, (perSheet = new Map()));
        const entry: { f?: string; v?: unknown } = {};
        if (input.startsWith("=")) entry.f = input;
        else if (input !== "") entry.v = coerceForPatch(input);
        perSheet.set(`${c.row},${c.col}`, entry);
        return { row: c.row + 1, column: c.col + 1, value: input };
      });
      engine.writeCells(sheetName, `${sheetName} ${label}`, wire);
      if (history === "record" && undoCells.length > 0) {
        pushUndo({
          label,
          undo: () => writeCells(sheetName, undoCells.map((c) => ({ row: c.row, col: c.col, value: c.before === "" ? null : c.before })), "undo", "skip"),
          redo: () => writeCells(sheetName, undoCells.map((c) => ({ row: c.row, col: c.col, value: c.after === "" ? null : c.after })), "redo", "skip"),
        });
      }
      bumpVersion();
      scheduleDraw();
      return true;
    };

    const undo = () => {
      const entry = undoRef.current.pop();
      if (!entry) return;
      entry.undo();
      redoRef.current.push(entry);
      onUserEdit?.();
    };
    const redo = () => {
      const entry = redoRef.current.pop();
      if (!entry) return;
      entry.redo();
      undoRef.current.push(entry);
      onUserEdit?.();
    };

    /** Paste TSV/plain text at the active cell (Excel's clipboard text form). */
    const pasteText = (text: string, html?: string, mode: "normal" | "values" = "normal") => {
      const name = layoutRef.current?.sheets[activeRef.current]?.name;
      if (!name || text === "") return;
      const s = selRef.current;
      const rows = text.replace(/\r\n?/g, "\n").replace(/\n$/, "").split("\n");
      const cells: Array<{ row: number; col: number; value: string | null }> = [];
      const formats: Array<{ row: number; col: number; format: CellFormatShape }> = [];
      const own = clipRef.current;
      if (mode === "values") {
        rows.forEach((line, i) => line.split("\t").forEach((v, j) => cells.push({ row: s.ar + i, col: s.ac + j, value: v === "" ? null : v })));
      } else if (own && own.text === text.replace(/\r\n?/g, "\n").replace(/\n$/, "")) {
        // Pasting what this grid copied: formulas shift with the move, formats come along.
        const dr = s.ar - own.r0, dc = s.ac - own.c0;
        for (const c of own.cells) {
          const row = s.ar + c.dr, col = s.ac + c.dc;
          const value = c.input.startsWith("=") ? `=${shiftFormulaA1(c.input.slice(1), dr, dc) ?? c.input.slice(1)}` : c.input;
          cells.push({ row, col, value: value === "" ? null : value });
          if (c.format) formats.push({ row, col, format: c.format });
        }
      } else {
        rows.forEach((line, i) => line.split("\t").forEach((v, j) => cells.push({ row: s.ar + i, col: s.ac + j, value: v === "" ? null : v })));
        // Excel and other sheets put an HTML table on the clipboard with
        // inline styles; keep bold/italic, colours, alignment and the
        // number format where present.
        if (html) for (const f of formatsFromHtmlTable(html)) formats.push({ row: s.ar + f.dr, col: s.ac + f.dc, format: f.format });
      }
      if (cells.length === 0 || cells.length > 50_000) return;
      withUndoGroup("paste", () => {
        writeCells(name, cells, "paste");
        for (const f of formats.slice(0, 20_000)) applyFormatRef.current(name, f.row, f.col, f.row, f.col, f.format);
      });
      selRef.current = { r0: s.ar, c0: s.ac, r1: s.ar + rows.length - 1, c1: s.ac + Math.max(...rows.map((l) => l.split("\t").length)) - 1, ar: s.ar, ac: s.ac };
      onUserEdit?.();
      refreshFormulaBar();
      scheduleDraw();
    };

    // ----- engine deltas → invalidate tiles -------------------------------

    useEffect(() => {
      engine.registerApplier((sheetName, cells) => {
        const sheet = sheetIndexOf(sheetName);
        if (sheet < 0) return false;
        // Raw values land in place at once (reads see them immediately);
        // the touched loaded tiles then refetch for display strings.
        tilesRef.current.applyValues(sheet, cells);
        for (const c of cells) markDirty(sheet, c.row0, c.col0);
        const touched = new Set<string>();
        for (const c of cells) {
          const tr = Math.floor(c.row0 / TILE_ROWS), tc = Math.floor(c.col0 / TILE_COLS);
          const key = tileKey(sheet, tr, tc);
          if (tilesRef.current.has(key) && !tilesRef.current.isPending(key)) touched.add(`${tr}:${tc}`);
        }
        for (const k of touched) {
          const [tr, tc] = k.split(":").map(Number);
          void fetchTile(sheet, tr, tc).then(scheduleDraw);
        }
        bumpVersion();
        scheduleDraw();
        if (sheet === activeRef.current) refreshFormulaBar();
        return true;
      });
      return () => engine.unregisterApplier();
    }, [tabId, scheduleDraw, refreshFormulaBar]);

    // ----- DOM wiring ------------------------------------------------------

    useEffect(() => {
      const sc = scrollerRef.current;
      if (!sc) return;
      const onScroll = () => scheduleDraw();
      sc.addEventListener("scroll", onScroll, { passive: true });
      const ro = new ResizeObserver(() => scheduleDraw());
      ro.observe(sc);
      // Zoom like Excel: trackpad pinch (WebKit sends gesture events, and
      // browsers send wheel + ctrlKey), Cmd+scroll on a Mac, Ctrl+scroll
      // elsewhere. Pinch deltas are continuous; zoom moves in 5% steps.
      let inGesture = false;
      let gestureBase = 1;
      let wheelAccum = 0;
      const onWheel = (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) return;
        e.preventDefault();
        if (inGesture) return;
        wheelAccum += -e.deltaY;
        const threshold = e.ctrlKey ? 8 : 40; // pinch deltas are small
        while (Math.abs(wheelAccum) >= threshold) {
          setZoomLevelRef.current(zoomRef.current + (wheelAccum > 0 ? 0.05 : -0.05));
          wheelAccum -= Math.sign(wheelAccum) * threshold;
        }
      };
      const onGestureStart = (e: Event) => { inGesture = true; gestureBase = zoomRef.current; e.preventDefault(); };
      const onGestureChange = (e: Event) => { e.preventDefault(); const scale = (e as unknown as { scale: number }).scale; if (scale > 0) setZoomLevelRef.current(gestureBase * scale); };
      const onGestureEnd = (e: Event) => { inGesture = false; e.preventDefault(); };
      sc.addEventListener("wheel", onWheel, { passive: false });
      sc.addEventListener("gesturestart", onGestureStart);
      sc.addEventListener("gesturechange", onGestureChange);
      sc.addEventListener("gestureend", onGestureEnd);
      return () => {
        sc.removeEventListener("scroll", onScroll);
        sc.removeEventListener("wheel", onWheel);
        sc.removeEventListener("gesturestart", onGestureStart);
        sc.removeEventListener("gesturechange", onGestureChange);
        sc.removeEventListener("gestureend", onGestureEnd);
        ro.disconnect();
      };
    }, [scheduleDraw]);

    useEffect(() => {
      if (editing && editing.origin === "cell" && editorRef.current) {
        const el = editorRef.current;
        el.focus();
        el.setSelectionRange(el.value.length, el.value.length);
      }
    }, [editing]);

    /** Header hit: which column/row header the point is on, and whether it is on a size edge. */
    const headerHit = (px: number, py: number): { kind: "col" | "row"; index: number; edge: number | null } | null => {
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!sc || !g) return null;
      const rect = sc.getBoundingClientRect();
      const z = viewScale();
      const x = (px - rect.left) / z;
      const y = (py - rect.top) / z;
      const fx = g.colLeft[g.frozenCols], fy = g.rowTop[g.frozenRows];
      if (y < HEADER_H && x >= HEADER_W) {
        const sheetX = sc.scrollLeft === 0 || x - HEADER_W < fx ? x - HEADER_W : x - HEADER_W + sc.scrollLeft / z;
        const c = g.colAt(sheetX);
        let edge: number | null = null;
        if (Math.abs(sheetX - g.colLeft[c + 1]) <= 4) edge = c;
        else if (c > 0 && Math.abs(sheetX - g.colLeft[c]) <= 4) edge = c - 1;
        return { kind: "col", index: c, edge };
      }
      if (x < HEADER_W && y >= HEADER_H) {
        const sheetY = sc.scrollTop === 0 || y - HEADER_H < fy ? y - HEADER_H : y - HEADER_H + sc.scrollTop / z;
        const r = g.rowAt(sheetY);
        let edge: number | null = null;
        if (Math.abs(sheetY - g.rowTop[r + 1]) <= 4) edge = r;
        else if (r > 0 && Math.abs(sheetY - g.rowTop[r]) <= 4) edge = r - 1;
        return { kind: "row", index: r, edge };
      }
      return null;
    };

    const onMouseDown = (e: React.MouseEvent) => {
      if (menu) setMenu(null);
      if (tabMenu) setTabMenu(null);
      if (filterMenu) setFilterMenu(null);
      if (e.button === 2) return; // context menu handled by onContextMenu
      if (editing) {
        const input = activeInput();
        const caret = input?.selectionStart ?? editing.text.length;
        const hit = cellAtPoint(e.clientX, e.clientY);
        if (hit && refInsertable(editing.text, caret)) {
          e.preventDefault(); // keep the caret in the editor
          refDrag.current = true;
          refPick.current = refPick.current && caret === refPick.current.end ? refPick.current : null;
          putRef(hit, hit);
          return;
        }
        commitEdit(input?.value ?? editing.text, "none");
      }
      const g = geom(activeRef.current);
      const hh = headerHit(e.clientX, e.clientY);
      if (hh && g) {
        if (hh.edge !== null) {
          const size = hh.kind === "col" ? g.colWidth(hh.edge) : g.rowHeight(hh.edge);
          resizing.current = { kind: hh.kind, index: hh.edge, start: hh.kind === "col" ? e.clientX : e.clientY, size };
          e.preventDefault();
          return;
        }
        // Click on a header selects the whole column / row.
        if (!(e.metaKey || e.ctrlKey)) extraRanges.current = [];
        else { const s = selRef.current; extraRanges.current.push({ r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1) }); }
        if (hh.kind === "col") selRef.current = { r0: 0, c0: hh.index, r1: g.rowCount - 1, c1: hh.index, ar: 0, ac: hh.index };
        else selRef.current = { r0: hh.index, c0: 0, r1: hh.index, c1: g.colCount - 1, ar: hh.index, ac: 0 };
        refreshFormulaBar();
        scheduleDraw();
        scrollerRef.current?.focus();
        return;
      }
      if (e.button === 0 && filterHitsRef.current.length) {
        const srect = scrollerRef.current!.getBoundingClientRect();
        const px = e.clientX - srect.left, py = e.clientY - srect.top;
        const fh = filterHitsRef.current.find((h) => px >= h.x0 && px <= h.x1 && py >= h.y0 && py <= h.y1);
        if (fh) {
          e.preventDefault();
          const root = (e.currentTarget as HTMLElement).closest("[data-vp-root]")?.getBoundingClientRect();
          openFilterMenu(fh.col, e.clientX - (root?.left ?? 0), e.clientY - (root?.top ?? 0));
          return;
        }
      }
      const hit = cellAtPoint(e.clientX, e.clientY);
      if (!hit) return;
      if (e.metaKey || e.ctrlKey) {
        const s = selRef.current;
        extraRanges.current.push({ r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1) });
      } else if (!e.shiftKey) {
        extraRanges.current = [];
      }
      if (onFillHandle(e.clientX, e.clientY)) {
        const s = selRef.current;
        filling.current = { r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1), toR: Math.max(s.r0, s.r1), toC: Math.max(s.c0, s.c1) };
        e.preventDefault();
        return;
      }
      dragging.current = true;
      dragPos.current = { x: e.clientX, y: e.clientY };
      setSelection(hit.r, hit.c, e.shiftKey);
      scrollerRef.current?.focus();
      autoScrollTick();
    };

    /** While dragging past the body edges, scroll and extend the selection. */
    const autoScrollTick = () => {
      cancelAnimationFrame(autoScrollReq.current);
      autoScrollReq.current = 0;
      if (!dragging.current) return;
      const sc = scrollerRef.current;
      const pos = dragPos.current;
      if (sc && pos) {
        const rect = sc.getBoundingClientRect();
        const step = 24;
        let dx = 0, dy = 0;
        if (pos.x > rect.right - 16) dx = step;
        else if (pos.x < rect.left + HEADER_W) dx = -step;
        if (pos.y > rect.bottom - 16) dy = step;
        else if (pos.y < rect.top + HEADER_H) dy = -step;
        if (dx || dy) {
          sc.scrollLeft += dx;
          sc.scrollTop += dy;
          const hit = cellAtPoint(Math.min(Math.max(pos.x, rect.left + HEADER_W + 1), rect.right - 1), Math.min(Math.max(pos.y, rect.top + HEADER_H + 1), rect.bottom - 1));
          if (hit) setSelection(hit.r + (dy > 0 ? 1 : 0), hit.c + (dx > 0 ? 1 : 0), true);
        }
      }
      autoScrollReq.current = requestAnimationFrame(autoScrollTick);
    };
    const onMouseMove = (e: React.MouseEvent) => {
      if (filling.current) {
        const hit = cellAtPoint(e.clientX, e.clientY);
        if (hit) {
          const f = filling.current;
          // Extend along the dominant axis only (down/up or right/left).
          const dR = hit.r > f.r1 ? hit.r - f.r1 : hit.r < f.r0 ? f.r0 - hit.r : 0;
          const dC = hit.c > f.c1 ? hit.c - f.c1 : hit.c < f.c0 ? f.c0 - hit.c : 0;
          if (dR >= dC) { f.toR = hit.r; f.toC = f.c1; } else { f.toC = hit.c; f.toR = f.r1; }
          scheduleDraw();
        }
        return;
      }
      if (refDrag.current) {
        const hit = cellAtPoint(e.clientX, e.clientY);
        if (hit && refPick.current) putRef(hit, refPick.current.anchor);
        return;
      }
      const rs = resizing.current;
      if (rs) {
        const delta = (rs.kind === "col" ? e.clientX : e.clientY) - rs.start;
        setSize(activeRef.current, rs.kind, rs.index, Math.max(0, rs.size + delta), false);
        return;
      }
      if (!dragging.current) {
        const hh = headerHit(e.clientX, e.clientY);
        const sc = scrollerRef.current;
        if (sc) sc.style.cursor = hh?.edge !== null && hh?.edge !== undefined ? (hh.kind === "col" ? "col-resize" : "row-resize") : "";
        if (sc && !hh && onFillHandle(e.clientX, e.clientY)) sc.style.cursor = "crosshair";
        const hit = cellAtPoint(e.clientX, e.clientY);
        const note = hit ? notesRef.current.get(sheetNameOf(activeRef.current))?.get(`${hit.r},${hit.c}`) : undefined;
        if (note) {
          const root = sc?.parentElement?.getBoundingClientRect();
          const nx = e.clientX - (root?.left ?? 0) + 12, ny = e.clientY - (root?.top ?? 0) + 12;
          if (!hoverNote || hoverNote.text !== note) setHoverNote({ x: nx, y: ny, text: note });
        } else if (hoverNote) setHoverNote(null);
        return;
      }
      dragPos.current = { x: e.clientX, y: e.clientY };
      const hit = cellAtPoint(e.clientX, e.clientY);
      if (hit) setSelection(hit.r, hit.c, true);
    };
    const onMouseUp = (e: React.MouseEvent) => {
      if (filling.current) {
        const f = filling.current;
        filling.current = null;
        const sc = scrollerRef.current;
        if (sc) sc.style.cursor = "";
        performFill(f);
        return;
      }
      refDrag.current = false;
      dragging.current = false;
      dragPos.current = null;
      cancelAnimationFrame(autoScrollReq.current);
      autoScrollReq.current = 0;
      const rs = resizing.current;
      if (rs) {
        resizing.current = null;
        const delta = (rs.kind === "col" ? e.clientX : e.clientY) - rs.start;
        setSize(activeRef.current, rs.kind, rs.index, Math.max(0, rs.size + delta), true);
        onUserEdit?.();
      }
    };
    const onDoubleClick = (e: React.MouseEvent) => {
      const hh = headerHit(e.clientX, e.clientY);
      if (hh && hh.edge !== null) {
        void autofit(hh.kind, hh.edge);
        return;
      }
      if (hh) return;
      startEdit();
    };

    /** Column: widest rendered text in the column (+ padding). Row: tallest
     * wrapped text or font. Reads the whole column/row from the engine. */
    const autofit = async (kind: "col" | "row", index: number) => {
      const sheet = activeRef.current;
      const l = layoutRef.current?.sheets[sheet];
      const canvas = canvasRef.current;
      const ctx = canvas?.getContext("2d");
      if (!l || !ctx) return;
      await engine.settled(20_000); // pending writes first
      const win = kind === "col"
        ? await engine.fetchWindow(sheet, 1, Math.min(Math.max(l.maxRow, 1), 65_536), index + 1, index + 1)
        : await engine.fetchWindow(sheet, index + 1, index + 1, 1, Math.min(Math.max(l.maxCol, 1), 65_536));
      if (!win || !Array.isArray(win.cells)) return;
      stylesRef.current.merge(win.styles ?? {});
      const g = geom(sheet);
      let best = 0;
      for (const c of win.cells as EngineWindowCell[]) {
        const gc: GridCell = { v: c.v ?? null, ...(c.d !== undefined ? { d: c.d } : {}), ...(c.f !== undefined ? { f: c.f } : {}), ...(c.s !== undefined ? { s: c.s } : {}) };
        const text = cellText(gc);
        if (text === "") continue;
        const st = styleWithOverlay(stylesRef.current.get(c.s), overlayOf(sheet, c.r - 1, c.c - 1));
        ctx.font = fontString(st);
        if (kind === "col") {
          if (g?.mergeOf(c.r - 1, c.c - 1)) continue;
          best = Math.max(best, ctx.measureText(text).width + 6);
        } else {
          const m = ctx.measureText("Mg");
          const lineH = (m.actualBoundingBoxAscent || st.sizePt) + (m.actualBoundingBoxDescent || st.sizePt * 0.25) + 2;
          const w = g ? g.colWidth(c.c - 1) : DEFAULT_COL_PX;
          const lines = st.wrap ? wrapText(ctx, text, w - 4).length : 1;
          best = Math.max(best, lines * lineH + 4);
        }
      }
      if (best <= 0) return;
      setSize(sheet, kind, index, Math.max(kind === "col" ? 20 : DEFAULT_ROW_PX, Math.ceil(best)), true);
      onUserEdit?.();
    };

    const onKeyDown = (e: React.KeyboardEvent) => {
      if (menu) { if (e.key === "Escape") { e.preventDefault(); setMenu(null); } return; }
      if (editing) return;
      const s = selRef.current;
      const g = geom(activeRef.current);
      if (!g) return;
      const page = Math.max(1, Math.floor(((scrollerRef.current?.clientHeight ?? 400) - HEADER_H) / DEFAULT_ROW_PX) - 1);
      const pageCols = Math.max(1, Math.floor(((scrollerRef.current?.clientWidth ?? 800) - HEADER_W) / DEFAULT_COL_PX) - 1);
      const move = (dr: number, dc: number) => {
        e.preventDefault();
        if (!e.shiftKey) extraRanges.current = [];
        const r = s.ar + dr, c = s.ac + dc;
        setSelection(r, c, e.shiftKey);
        scrollIntoView(Math.max(0, Math.min(g.rowCount - 1, r)), Math.max(0, Math.min(g.colCount - 1, c)));
      };
      const ctrl = e.ctrlKey || e.metaKey;
      if (ctrl && e.altKey && (e.key === "=" || e.key === "+" || e.key === "-" || e.key === "≠" || e.key === "–" || e.key === "0" || e.key === "º")) {
        // Zoom: Cmd+Option and =, - or 0 (Excel keeps Cmd+- and Cmd+= for cells).
        e.preventDefault();
        setZoomLevel(e.key === "0" || e.key === "º" ? 1 : zoomRef.current + (e.key === "-" || e.key === "–" ? -0.1 : 0.1));
        return;
      }
      if (ctrl && !e.altKey && excelShortcut(e)) return;
      if (ctrl && (e.key === "a" || e.key === "A")) {
        e.preventDefault();
        const l = layoutRef.current?.sheets[activeRef.current];
        extraRanges.current = [];
        selRef.current = { r0: 0, c0: 0, r1: Math.max(0, (l?.usedRow ?? l?.maxRow ?? 1) - 1), c1: Math.max(0, (l?.usedCol ?? l?.maxCol ?? 1) - 1), ar: s.ar, ac: s.ac };
        refreshFormulaBar(); setStats(selectionStats()); scheduleDraw();
        return;
      }
      if (e.key === " " && (ctrl || e.shiftKey)) {
        e.preventDefault();
        const r0 = Math.min(s.r0, s.r1), r1 = Math.max(s.r0, s.r1), c0 = Math.min(s.c0, s.c1), c1 = Math.max(s.c0, s.c1);
        selRef.current = ctrl ? { ...s, r0: 0, r1: g.rowCount - 1, c0, c1 } : { ...s, c0: 0, c1: g.colCount - 1, r0, r1 };
        refreshFormulaBar(); setStats(selectionStats()); scheduleDraw();
        return;
      }
      if (ctrl && (e.key === "f" || e.key === "F")) {
        e.preventDefault();
        setFind((f) => f ?? { query: "", hits: [], at: -1, scope: "sheet", replace: "" });
        requestAnimationFrame(() => findInputRef.current?.focus());
        return;
      }
      if (ctrl && (e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "ArrowRight" || e.key === "ArrowLeft")) {
        e.preventDefault();
        const dr = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0;
        const dc = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0;
        const extend = e.shiftKey;
        void jumpBlock(dr, dc).then(({ r, c }) => {
          setSelection(r, c, extend);
          scrollIntoView(r, c);
        });
        return;
      }
      switch (e.key) {
        case "ArrowDown": return move(1, 0);
        case "ArrowUp": return move(-1, 0);
        case "ArrowRight": return move(0, 1);
        case "ArrowLeft": return move(0, -1);
        case "End": e.preventDefault(); if (ctrl) { const l = layoutRef.current?.sheets[activeRef.current]; const r = Math.max(0, (l?.usedRow ?? l?.maxRow ?? 1) - 1), c = Math.max(0, (l?.usedCol ?? l?.maxCol ?? 1) - 1); setSelection(r, c, e.shiftKey); scrollIntoView(r, c); } return;
        case "PageDown": return e.altKey ? move(0, pageCols) : move(page, 0);
        case "PageUp": return e.altKey ? move(0, -pageCols) : move(-page, 0);
        case "Home": e.preventDefault(); setSelection(e.ctrlKey || e.metaKey ? 0 : s.ar, 0); scrollIntoView(e.ctrlKey || e.metaKey ? 0 : s.ar, 0); return;
        case "Enter": e.preventDefault(); if (e.shiftKey) return move(-1, 0); return startEdit();
        case "F2": e.preventDefault(); return startEdit();
        case "Tab": return move(0, e.shiftKey ? -1 : 1);
        case "Delete":
        case "Backspace": {
          e.preventDefault();
          const name = layoutRef.current?.sheets[activeRef.current]?.name;
          if (!name) return;
          const cells: Array<{ row: number; col: number; value: null }> = [];
          const all = [...extraRanges.current, { r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1) }];
          for (const x of all) for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) cells.push({ row: r, col: c, value: null });
          if (cells.length > 50_000) return;
          writeCells(name, cells, "clear");
          onUserEdit?.();
          return;
        }
        default:
          break;
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === "z" || e.key === "Z")) {
        e.preventDefault();
        if (e.shiftKey) redo();
        else undo();
        return;
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === "y" || e.key === "Y")) {
        e.preventDefault();
        redo();
        return;
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === "v" || e.key === "V")) {
        if (e.shiftKey) {
          // Paste values only: what the cells show, no formulas, no formats.
          e.preventDefault();
          void navigator.clipboard?.readText().then((t) => { if (t) pasteText(t, undefined, "values"); }).catch(() => {});
        }
        // Otherwise onPaste handles it (the clipboard is only readable from the event).
        return;
      }
      if ((e.metaKey || e.ctrlKey) && (e.key === "c" || e.key === "C" || e.key === "x" || e.key === "X")) {
        e.preventDefault();
        copySelection(e.key === "x" || e.key === "X");
        return;
      }
      if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
        e.preventDefault();
        startEdit(e.key);
      }
    };

    /** Excel's Ctrl+arrow: to the edge of the current data block, or to the
     * next block when starting from an empty cell or a block edge; past the
     * last block it lands on the grid edge (Excel's XFD / 1048576). A merged
     * range counts as one cell (its anchor's content), as in Excel. Tiles
     * along the way are fetched on demand. */
    const jumpBlock = async (dr: number, dc: number): Promise<{ r: number; c: number }> => {
      const sheet = activeRef.current;
      const g = geom(sheet);
      const s = selRef.current;
      if (!g) return { r: s.ar, c: s.ac };
      const l = layoutRef.current?.sheets[sheet];
      // Content can only exist inside the layout's extent; the grid itself is larger.
      const maxR = Math.max(0, (l?.maxRow ?? 1) - 1), maxC = Math.max(0, (l?.maxCol ?? 1) - 1);
      const gridEdge = (r: number, c: number) => ({ r: dr > 0 ? g.rowCount - 1 : dr < 0 ? 0 : r, c: dc > 0 ? g.colCount - 1 : dc < 0 ? 0 : c });
      const has = async (r: number, c: number) => {
        if (r < 0 || c < 0 || r > maxR || c > maxC) return false;
        const m = g.mergeOf(r, c);
        if (m) { r = m.r0; c = m.c0; }
        const key = tileKey(sheet, Math.floor(r / TILE_ROWS), Math.floor(c / TILE_COLS));
        if (!tilesRef.current.has(key)) await fetchTile(sheet, Math.floor(r / TILE_ROWS), Math.floor(c / TILE_COLS));
        const cell = tilesRef.current.cell(sheet, r, c);
        return !!cell && cellText(cell) !== "";
      };
      // One step in the travel direction; a merged range is crossed in one go.
      const step = (r: number, c: number) => {
        const m = g.mergeOf(r, c);
        if (m) return { r: dr > 0 ? m.r1 + 1 : dr < 0 ? m.r0 - 1 : r, c: dc > 0 ? m.c1 + 1 : dc < 0 ? m.c0 - 1 : c };
        return { r: r + dr, c: c + dc };
      };
      const anchorOf = (r: number, c: number) => { const m = g.mergeOf(r, c); return m ? { r: m.r0, c: m.c0 } : { r, c }; };
      const inExtent = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC;
      let { r, c } = anchorOf(s.ar, s.ac);
      let n = step(r, c);
      if (!inExtent(n.r, n.c)) return gridEdge(r, c);
      const here = await has(r, c);
      const next = await has(n.r, n.c);
      if (here && next) {
        // Inside a block: go to its last filled cell.
        while (inExtent(n.r, n.c) && (await has(n.r, n.c))) { ({ r, c } = anchorOf(n.r, n.c)); n = step(r, c); }
        return { r, c };
      }
      // On an edge or in a gap: skip empties to the next filled cell.
      let steps = 0;
      while (inExtent(n.r, n.c) && !(await has(n.r, n.c)) && steps < 200_000) { n = step(n.r, n.c); steps++; }
      if (!inExtent(n.r, n.c)) return gridEdge(r, c);
      return anchorOf(n.r, n.c);
    };

    /** Copy the selection as tab-separated text; `cut` also clears it. */
    const copySelection = (cut: boolean) => {
      const s = selRef.current;
      const sheet = activeRef.current;
      const r0 = Math.min(s.r0, s.r1), r1 = Math.max(s.r0, s.r1), c0 = Math.min(s.c0, s.c1), c1 = Math.max(s.c0, s.c1);
      if ((r1 - r0 + 1) * (c1 - c0 + 1) > 200_000) return;
      const rows: string[] = [];
      const kept: Array<{ dr: number; dc: number; input: string; format: CellFormatShape | null }> = [];
      const name = layoutRef.current?.sheets[sheet]?.name ?? "";
      for (let r = r0; r <= r1; r++) {
        const cols: string[] = [];
        for (let c = c0; c <= c1; c++) {
          const cell = tilesRef.current.cell(sheet, r, c);
          cols.push(cellText(cell));
          const input = cell ? cell.f ?? rawInput(cell) : "";
          const format = cell ? formatOf(cell.s, overlayOf(sheet, r, c)) : null;
          if (input !== "" || (format && Object.keys(format).length > 0)) kept.push({ dr: r - r0, dc: c - c0, input, format: format && Object.keys(format).length ? format : null });
        }
        rows.push(cols.join("\t"));
      }
      const text = rows.join("\n");
      clipRef.current = { text, r0, c0, cells: kept };
      void name;
      // Plain text for everything, plus an HTML table with inline styles so
      // Excel, Numbers and Sheets keep bold, colours, fills and number
      // formats when they receive the paste.
      const html = htmlTableForCopy(r0, c0, r1, c1, (r, c) => {
        const cell = tilesRef.current.cell(sheet, r, c);
        return { text: cellText(cell), format: cell ? formatOf(cell.s, overlayOf(sheet, r, c)) : null, align: cell ? cellAlign(cell, styleWithOverlay(stylesRef.current.get(cell.s), overlayOf(sheet, r, c))) : "left" };
      });
      const clip = navigator.clipboard as Clipboard | undefined;
      if (clip && typeof ClipboardItem !== "undefined" && clip.write) {
        void clip.write([new ClipboardItem({ "text/plain": new Blob([text], { type: "text/plain" }), "text/html": new Blob([html], { type: "text/html" }) })]).catch(() => clip.writeText(text));
      } else {
        void clip?.writeText(text);
      }
      if (cut) {
        const name = layoutRef.current?.sheets[sheet]?.name;
        if (!name) return;
        const cells: Array<{ row: number; col: number; value: null }> = [];
        for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) cells.push({ row: r, col: c, value: null });
        if (cells.length > 50_000) return;
        writeCells(name, cells, "cut");
        onUserEdit?.();
      }
    };

    /** Count / sum / average of the selection, for the status line. */
    const selectionStats = (): string => {
      const s = selRef.current;
      const ranges = [...extraRanges.current, { r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1) }];
      const area = ranges.reduce((n, x) => n + (x.r1 - x.r0 + 1) * (x.c1 - x.c0 + 1), 0);
      if (area <= 1 || area > 100_000) return "";
      let count = 0, nums = 0, sum = 0;
      for (const x of ranges) {
        for (let r = x.r0; r <= x.r1; r++) {
          for (let c = x.c0; c <= x.c1; c++) {
            const cell = tilesRef.current.cell(activeRef.current, r, c);
            if (!cell || cellText(cell) === "") continue;
            count++;
            if (typeof cell.v === "number") { nums++; sum += cell.v; }
          }
        }
      }
      if (count === 0) return "";
      const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 2 });
      return nums > 0 ? `Count ${count} · Sum ${fmt(sum)} · Avg ${fmt(sum / nums)}` : `Count ${count}`;
    };

    const runFind = async (query: string, scope: "sheet" | "workbook") => {
      const hits = query ? await engine.find(query, scope === "sheet" ? activeRef.current : undefined, 500) : [];
      setFind((f) => ({ query, hits, at: hits.length ? 0 : -1, scope, replace: f?.replace ?? "" }));
      if (hits.length) gotoHit(hits[0]);
    };

    /** Replace the query in the current hit, or in every hit, case-insensitively. */
    const replaceHits = async (all: boolean) => {
      if (!find || find.hits.length === 0 || !find.query) return;
      const targets = all ? find.hits : [find.hits[Math.max(0, find.at)]];
      const re = new RegExp(find.query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
      const bySheet = new Map<number, Array<{ row: number; col: number; value: string | null }>>();
      for (const h of targets) {
        await prefetchRef.current(h.sheet, h.row - 1, h.row - 1, h.column - 1, h.column - 1);
        const cell = tilesRef.current.cell(h.sheet, h.row - 1, h.column - 1);
        if (!cell) continue;
        const input = cell.f ?? rawInput(cell);
        const next = input.replace(re, find.replace);
        if (next === input) continue;
        let list = bySheet.get(h.sheet);
        if (!list) bySheet.set(h.sheet, (list = []));
        list.push({ row: h.row - 1, col: h.column - 1, value: next === "" ? null : next });
      }
      withUndoGroup("replace", () => { for (const [sheet, cells] of bySheet) writeCells(sheetNameOf(sheet), cells, "replace"); });
      if (bySheet.size) onUserEdit?.();
      lastFindQuery.current = find.query;
      await runFind(find.query, find.scope);
    };
    const gotoHit = (h: { sheet: number; row: number; column: number }) => {
      if (h.sheet !== activeRef.current) activateSheet(h.sheet);
      setSelection(h.row - 1, h.column - 1);
      scrollIntoView(h.row - 1, h.column - 1);
    };
    const stepFind = (dir: 1 | -1) => {
      if (!find || find.hits.length === 0) return;
      const at = (find.at + dir + find.hits.length) % find.hits.length;
      setFind({ ...find, at });
      gotoHit(find.hits[at]);
    };

    /** Is the pointer on the selection's fill handle (bottom-right square)? */
    const onFillHandle = (px: number, py: number): boolean => {
      const sc = scrollerRef.current;
      const g = geom(activeRef.current);
      if (!sc || !g) return false;
      const s = selRef.current;
      const r1 = Math.max(s.r0, s.r1), c1 = Math.max(s.c0, s.c1);
      const rect = sc.getBoundingClientRect();
      const fx = g.colLeft[g.frozenCols], fy = g.rowTop[g.frozenRows];
      const cornerX = g.colLeft[c1 + 1], cornerY = g.rowTop[r1 + 1];
      const z = viewScale();
      const screenX = rect.left + (HEADER_W + (sc.scrollLeft === 0 || cornerX <= fx ? cornerX : cornerX - sc.scrollLeft / z)) * z;
      const screenY = rect.top + (HEADER_H + (sc.scrollTop === 0 || cornerY <= fy ? cornerY : cornerY - sc.scrollTop / z)) * z;
      return Math.abs(px - screenX) <= 5 && Math.abs(py - screenY) <= 5;
    };

    /** Excel-style fill: formulas shift relatively; a numeric column or row
     * of ≥2 values continues its linear series; everything else repeats. */
    const performFill = (f: { r0: number; c0: number; r1: number; c1: number; toR: number; toC: number }) => {
      const sheet = activeRef.current;
      const name = sheetNameOf(sheet);
      if (!name) return;
      const down = f.toR !== f.r1, right = f.toC !== f.c1;
      if (!down && !right) return;
      const writes: Array<{ row: number; col: number; value: string | null }> = [];
      const src = (r: number, c: number) => tilesRef.current.cell(sheet, r, c);
      const input = (cell: GridCell | undefined) => (cell ? cell.f ?? rawInput(cell) : "");
      if (down) {
        const up = f.toR < f.r0;
        const n = f.r1 - f.r0 + 1;
        const targets = up ? range(f.toR, f.r0 - 1) : range(f.r1 + 1, f.toR);
        for (let c = f.c0; c <= f.c1; c++) {
          const vals = range(f.r0, f.r1).map((r) => src(r, c));
          const nums = vals.map((v) => (v && v.f === undefined && typeof v.v === "number" ? v.v : null));
          const series = n >= 2 && nums.every((x) => x !== null) ? { a: nums[0] as number, d: ((nums[n - 1] as number) - (nums[0] as number)) / (n - 1) } : null;
          targets.forEach((r, k) => {
            const srcIdx = up ? ((((f.r0 - 1 - r) % n) + n) % n) : k % n;
            const from = vals[up ? n - 1 - srcIdx : srcIdx];
            const offset = up ? -(Math.floor((f.r0 - 1 - r) / n) + 1) * n : (Math.floor(k / n) + 1) * n;
            let value: string | null;
            if (series) value = String(series.a + series.d * (r - f.r0));
            else if (from?.f) value = shiftFormulaA1(from.f.replace(/^=/, ""), offset, 0) ?? from.f;
            else value = input(from) || null;
            if (value && from?.f && !value.startsWith("=")) value = `=${value}`;
            writes.push({ row: r, col: c, value });
          });
        }
      } else {
        const left = f.toC < f.c0;
        const n = f.c1 - f.c0 + 1;
        const targets = left ? range(f.toC, f.c0 - 1) : range(f.c1 + 1, f.toC);
        for (let r = f.r0; r <= f.r1; r++) {
          const vals = range(f.c0, f.c1).map((c) => src(r, c));
          const nums = vals.map((v) => (v && v.f === undefined && typeof v.v === "number" ? v.v : null));
          const series = n >= 2 && nums.every((x) => x !== null) ? { a: nums[0] as number, d: ((nums[n - 1] as number) - (nums[0] as number)) / (n - 1) } : null;
          targets.forEach((c, k) => {
            const srcIdx = left ? ((((f.c0 - 1 - c) % n) + n) % n) : k % n;
            const from = vals[left ? n - 1 - srcIdx : srcIdx];
            const offset = left ? -(Math.floor((f.c0 - 1 - c) / n) + 1) * n : (Math.floor(k / n) + 1) * n;
            let value: string | null;
            if (series) value = String(series.a + series.d * (c - f.c0));
            else if (from?.f) value = shiftFormulaA1(from.f.replace(/^=/, ""), 0, offset) ?? from.f;
            else value = input(from) || null;
            if (value && from?.f && !value.startsWith("=")) value = `=${value}`;
            writes.push({ row: r, col: c, value });
          });
        }
      }
      if (writes.length === 0 || writes.length > 50_000) return;
      writeCells(name, writes, "fill");
      onUserEdit?.();
      selRef.current = { r0: Math.min(f.r0, f.toR), c0: Math.min(f.c0, f.toC), r1: Math.max(f.r1, f.toR), c1: Math.max(f.c1, f.toC), ar: f.r0, ac: f.c0 };
      refreshFormulaBar();
      setStats(selectionStats());
      scheduleDraw();
    };

    /** Commit the entry being typed into every selected cell (Cmd+Enter),
     * shifting relative references per cell as Excel does. */
    const commitToSelection = (text: string) => {
      const e = editingRef.current;
      const name = sheetNameOf(activeRef.current);
      if (!e || !name) return;
      editingRef.current = null;
      refPick.current = null;
      refDrag.current = false;
      setEditing(null);
      const s = selRef.current;
      const cells: Array<{ row: number; col: number; value: string | null }> = [];
      for (let r = Math.min(s.r0, s.r1); r <= Math.max(s.r0, s.r1); r++) {
        for (let c = Math.min(s.c0, s.c1); c <= Math.max(s.c0, s.c1); c++) {
          const value = text.startsWith("=") ? `=${shiftFormulaA1(text.slice(1), r - e.row, c - e.col) ?? text.slice(1)}` : text;
          cells.push({ row: r, col: c, value: value === "" ? null : value });
        }
      }
      if (cells.length <= 50_000) { writeCells(name, cells, "fill selection"); onUserEdit?.(); }
      scrollerRef.current?.focus();
    };

    /** Excel's Ctrl+D / Ctrl+R: fill the selection from its first row / column
     * (a single cell takes the cell above / to the left). */
    const fillFromEdge = (dir: "down" | "right") => {
      const s = selRef.current;
      const sheet = activeRef.current;
      const name = sheetNameOf(sheet);
      if (!name) return;
      let r0 = Math.min(s.r0, s.r1), r1 = Math.max(s.r0, s.r1), c0 = Math.min(s.c0, s.c1), c1 = Math.max(s.c0, s.c1);
      if (r0 === r1 && c0 === c1) { if (dir === "down") { if (r0 === 0) return; r0 -= 1; } else { if (c0 === 0) return; c0 -= 1; } }
      performFill(dir === "down" ? { r0, c0, r1: r0, c1, toR: r1, toC: c1 } : { r0, c0, r1, c1: c0, toR: r1, toC: c1 });
      selRef.current = { ...s, r0, c0, r1, c1 };
      refreshFormulaBar();
      scheduleDraw();
    };

    const formatSelection = (format: CellFormatShape) => {
      const s = selRef.current;
      const name = sheetNameOf(activeRef.current);
      if (!name) return;
      const ranges = [...extraRanges.current, { r0: Math.min(s.r0, s.r1), c0: Math.min(s.c0, s.c1), r1: Math.max(s.r0, s.r1), c1: Math.max(s.c0, s.c1) }];
      withUndoGroup("format", () => { for (const x of ranges) applyFormatRef.current(name, x.r0, x.c0, x.r1, x.c1, format); });
      onUserEdit?.();
    };

    const excelSerialNow = (withTime: boolean) => {
      const now = new Date();
      const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), withTime ? now.getHours() : 0, withTime ? now.getMinutes() : 0, withTime ? now.getSeconds() : 0);
      const serial = (utc - Date.UTC(1899, 11, 30)) / 86_400_000;
      return withTime ? serial - Math.floor(serial) : Math.floor(serial);
    };

    /** The Excel keyboard set on Cmd (Ctrl). Returns true when handled. */
    const excelShortcut = (e: React.KeyboardEvent): boolean => {
      const s = selRef.current;
      const sheet = activeRef.current;
      const name = sheetNameOf(sheet);
      const g = geom(sheet);
      if (!name || !g) return false;
      const k = e.key.toLowerCase();
      const cell = tilesRef.current.cell(sheet, s.ar, s.ac);
      const fmt = formatOf(cell?.s, overlayOf(sheet, s.ar, s.ac));
      const wholeRows = Math.min(s.c0, s.c1) === 0 && Math.max(s.c0, s.c1) === g.colCount - 1;
      const wholeCols = Math.min(s.r0, s.r1) === 0 && Math.max(s.r0, s.r1) === g.rowCount - 1;
      const r0 = Math.min(s.r0, s.r1), r1 = Math.max(s.r0, s.r1), c0 = Math.min(s.c0, s.c1), c1 = Math.max(s.c0, s.c1);
      const done = () => { e.preventDefault(); return true; };
      if (!e.shiftKey) {
        switch (k) {
          case "d": fillFromEdge("down"); return done();
          case "r": fillFromEdge("right"); return done();
          case "b": formatSelection({ bold: !fmt.bold }); return done();
          case "i": formatSelection({ italic: !fmt.italic }); return done();
          case "u": formatSelection({ underline: !fmt.underline }); return done();
          case ";": writeCells(name, [{ row: s.ar, col: s.ac, value: String(excelSerialNow(false)) }], "date"); formatSelection({ number_format: "m/d/yyyy" }); return done();
          case "'": { // copy the formula or value from the cell above
            if (s.ar === 0) return done();
            const above = tilesRef.current.cell(sheet, s.ar - 1, s.ac);
            const input = above ? above.f ?? rawInput(above) : "";
            if (input) { writeCells(name, [{ row: s.ar, col: s.ac, value: input }], "copy above"); onUserEdit?.(); }
            return done();
          }
          case "9": for (let r = r0; r <= r1; r++) setSize(sheet, "row", r, 0, true); onUserEdit?.(); return done();
          case "0": for (let c = c0; c <= c1; c++) setSize(sheet, "col", c, 0, true); onUserEdit?.(); return done();
          case "-": { // delete rows (or columns when whole columns are selected)
            if (wholeCols && !wholeRows) structuralRef.current("delete_columns", name, c0 + 1, c1 - c0 + 1);
            else structuralRef.current("delete_rows", name, r0 + 1, r1 - r0 + 1);
            onUserEdit?.();
            return done();
          }
          case "h": setFind((f) => f ?? { query: "", hits: [], at: -1, scope: "sheet", replace: "" }); requestAnimationFrame(() => findInputRef.current?.focus()); return done();
          default: return false;
        }
      }
      switch (e.key) {
        case ":": case ";": writeCells(name, [{ row: s.ar, col: s.ac, value: String(excelSerialNow(true)) }], "time"); formatSelection({ number_format: "h:mm AM/PM" }); return done();
        case "~": case "`": formatSelection({ number_format: "General" }); return done();
        case "!": case "1": formatSelection({ number_format: "#,##0.00" }); return done();
        case "@": case "2": formatSelection({ number_format: "h:mm AM/PM" }); return done();
        case "#": case "3": formatSelection({ number_format: "d-mmm-yy" }); return done();
        case "$": case "4": formatSelection({ number_format: "$#,##0.00_);($#,##0.00)" }); return done();
        case "%": case "5": formatSelection({ number_format: "0%" }); return done();
        case "^": case "6": formatSelection({ number_format: "0.00E+00" }); return done();
        case "l": case "L": toggleFilterAtSelection(); return done();
        case "(": case "9": { const g0 = geom(sheet); for (let r = Math.max(0, r0 - 1); r <= Math.min(g.rowCount - 1, r1 + 1); r++) if (g0 && g0.rowHeight(r) === 0) setSize(sheet, "row", r, DEFAULT_ROW_PX, true); onUserEdit?.(); return done(); }
        case ")": case "0": { const g0 = geom(sheet); for (let c = Math.max(0, c0 - 1); c <= Math.min(g.colCount - 1, c1 + 1); c++) if (g0 && g0.colWidth(c) === 0) setSize(sheet, "col", c, DEFAULT_COL_PX, true); onUserEdit?.(); return done(); }
        case "+": case "=": { // insert rows above (or columns to the left when whole columns are selected)
          if (wholeCols && !wholeRows) structuralRef.current("insert_columns", name, c0 + 1, c1 - c0 + 1);
          else structuralRef.current("insert_rows", name, r0 + 1, r1 - r0 + 1);
          onUserEdit?.();
          return done();
        }
        default: return false;
      }
    };

    /** Selection per sheet, kept across tab switches like Excel. */
    const selPosRef = useRef<Map<number, Selection>>(new Map());

    /** First visit to a sheet: scroll and select where the file was saved
     * (the pane's topLeftCell and the active cell), as Excel opens it. */
    const applySavedView = (index: number) => {
      const sc = scrollerRef.current;
      const g = geom(index);
      const l = layoutRef.current?.sheets[index];
      if (!g || !l) return;
      const z = viewScale();
      const topRow = Math.min(g.rowCount - 1, Math.max(0, (l.topRow ?? 1) - 1));
      const leftCol = Math.min(g.colCount - 1, Math.max(0, (l.leftColumn ?? 1) - 1));
      const fy = g.rowTop[g.frozenRows], fx = g.colLeft[g.frozenCols];
      if (sc) {
        sc.scrollLeft = topRow >= 0 && leftCol >= g.frozenCols ? Math.max(0, g.colLeft[leftCol] - fx) * z : 0;
        sc.scrollTop = topRow >= g.frozenRows ? Math.max(0, g.rowTop[topRow] - fy) * z : 0;
      }
      const ar = Math.min(g.rowCount - 1, Math.max(0, (l.activeRow ?? 1) - 1));
      const ac = Math.min(g.colCount - 1, Math.max(0, (l.activeCol ?? 1) - 1));
      selRef.current = { r0: ar, c0: ac, r1: ar, c1: ac, ar, ac };
    };

    const activateSheet = (index: number) => {
      const sc = scrollerRef.current;
      if (sc) scrollPosRef.current.set(activeRef.current, { x: sc.scrollLeft, y: sc.scrollTop });
      selPosRef.current.set(activeRef.current, selRef.current);
      activeRef.current = index;
      setActive(index);
      const z = zoomForSheet(index);
      if (z !== zoomRef.current) { zoomRef.current = z; setZoom(z); }
      updateSpacer();
      const pos = scrollPosRef.current.get(index);
      if (pos) {
        if (sc) {
          sc.scrollLeft = pos.x;
          sc.scrollTop = pos.y;
        }
        selRef.current = selPosRef.current.get(index) ?? { r0: 0, c0: 0, r1: 0, c1: 0, ar: 0, ac: 0 };
      } else {
        applySavedView(index);
      }
      refreshFormulaBar();
      scheduleDraw();
    };

    // ----- the handle ------------------------------------------------------

    const loadLayout = (layout: EngineLayout) => {
      layoutRef.current = layout;
      geomsRef.current.clear();
      initFilters(layout);
      stylesRef.current = new StyleBook(layout.styles ?? {});
      tilesRef.current.clear();
      setSheetNames(layout.sheets.map((s) => ({ name: s.name, hidden: s.hidden, color: s.color })));
      // Open on the sheet Excel had selected, if it is visible; else the first visible.
      const selected = layout.sheets.findIndex((s) => s.selected && !s.hidden);
      const firstVisible = selected >= 0 ? selected : Math.max(0, layout.sheets.findIndex((s) => !s.hidden));
      activeRef.current = firstVisible;
      setActive(firstVisible);
      sheetZoomRef.current.clear();
      scrollPosRef.current.clear();
      selPosRef.current.clear();
      zoomRef.current = zoomForSheet(firstVisible);
      setZoom(zoomRef.current);
      const total = layout.sheets.reduce((n, s) => n + s.cellCount, 0);
      setStatus(`${layout.sheets.length} sheets · ${total.toLocaleString()} cells · viewport grid`);
      updateSpacer();
      applySavedView(firstVisible);
      refreshFormulaBar();
      scheduleDraw();
    };

    const reloadLayout = async () => {
      // After structural / sheet ops: fetch the layout again, keep scroll.
      const layout = (await engine.fetchLayout()) as EngineLayout | null;
      if (!layout) return;
      const keep = activeRef.current;
      layoutRef.current = layout;
      for (const sh of layout.sheets) pendingSheetsRef.current.delete(sh.name);
      geomsRef.current.clear();
      stylesRef.current.merge(layout.styles ?? {});
      tilesRef.current.clear();
      setSheetNames(layout.sheets.map((s) => ({ name: s.name, hidden: s.hidden, color: s.color })));
      activeRef.current = Math.min(keep, layout.sheets.length - 1);
      setActive(activeRef.current);
      updateSpacer();
      bumpVersion();
      scheduleDraw();
    };

    const handle = useMemo<UniverGridHandle>(() => {
      const real: Partial<UniverGridHandle> & Record<string, unknown> = {
        loadBytes: async (bytes, onProgress, source) => {
          sourceBytesRef.current = bytes;
          if (!source) return { engineLoaded: false };
          onProgress?.("Loading workbook…");
          const layout = (await engine.openLayout(source.path, bytes)) as EngineLayout | null;
          if (!layout) {
            readyRef.current?.resolve(false);
            return { engineLoaded: false };
          }
          onProgress?.("Building grid…");
          loadLayout(layout);
          readyRef.current?.resolve(true);
          // The first full evaluate holds the engine for seconds; land the
          // tiles on screen before it starts so the grid is never blank —
          // and the first screen of every other sheet, so switching sheets
          // during the evaluate shows cells too (one small window each).
          await firstTilesLanded(3000);
          const sheets = layout.sheets.map((sh, i) => ({ sh, i })).filter(({ sh, i }) => i !== activeRef.current && sh.maxRow > 0 && sh.maxCol > 0);
          await Promise.all(sheets.map(({ i }) => fetchBlock({ sheet: i, tr0: 0, tr1: 0, tc0: 0, tc1: Math.min(1, Math.floor((layout.sheets[i].maxCol - 1) / TILE_COLS)) })));
          engine.startDeferredSettle();
          setCalculating(true);
          void engine.settled(600_000).finally(() => setCalculating(false));
          return { engineLoaded: true };
        },
        whenReady: () => readyRef.current!.promise,
        getSheetNames: () => [...(layoutRef.current?.sheets.map((s) => s.name) ?? []), ...pendingSheetsRef.current],
        // Every sheet is "lazy": readers must prefetch the ranges they read.
        isSheetLazy: () => true,
        ensureSheetLoaded: async (sheetName) => {
          const sheet = sheetIndexOf(sheetName);
          const l = layoutRef.current?.sheets[sheet];
          if (!l) return false;
          // Whole-sheet prefetch is bounded: agent-built sheets are small;
          // the trackers are read through prefetchRange / engine tools.
          if (l.cellCount > 300_000) return true;
          await prefetch(sheet, 0, Math.max(0, l.maxRow - 1), 0, Math.max(0, l.maxCol - 1));
          return true;
        },
        prefetchRange: async (sheetName: string, sr: number, sc: number, er: number, ec: number) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return;
          await prefetch(sheet, sr, er, sc, ec);
        },
        getContentVersion: () => `v${contentVersionRef.current}`,
        setCell: (sheet, row, col, value) => {
          writeCells(sheet, [{ row, col, value }], "set");
        },
        setCells: (sheet, cells) => writeCells(sheet, cells, "batch"),
        whenCalculated: async (timeoutMs = 15000) => {
          await engine.settled(timeoutMs);
        },
        getCell: (sheetName, row, col) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return null;
          const cell = tilesRef.current.cell(sheet, row, col);
          if (!cell) return null;
          // Agent-facing formula text drops the implicit-intersection marker,
          // as the Univer snapshot mapper did (the grid UI keeps it).
          const out: { value: any; display?: string; formula: string | null; fileSaved?: boolean } = { value: cell.v ?? null, formula: cell.f ? stripImplicitIntersection(cell.f) : null };
          // `display` only when the number format changes the rendering, as
          // the Univer grid reports it (the preview prints both).
          if (cell.d !== undefined && cell.d !== String(cell.v)) out.display = cell.d;
          if (cell.fileSaved) out.fileSaved = true;
          return out;
        },
        // A window cell flagged fileSaved holds the file's last-saved value
        // because the engine could not produce a usable one — the same
        // fact the Univer path derives from its load-time baseline.
        getFileSavedCell: (sheetName, row, col) => {
          const sheet = sheetIndexOf(sheetName);
          const cell = sheet >= 0 ? tilesRef.current.cell(sheet, row, col) : undefined;
          if (!cell?.fileSaved) return null;
          return { f: cell.f ? cell.f.replace(/^=/, "") : undefined, v: cell.v ?? null };
        },
        getExternalPin: () => null,
        getCalcHealth: async () => {
          if (calcHealthRef.current !== undefined) return calcHealthRef.current;
          await engine.settled(20_000);
          const report = await engine.calcHealth();
          if (report === null) return null; // first evaluate still pending — retry next call
          calcHealthRef.current = report;
          return report;
        },
        // Delete-sheet undo: the sheet's cells are fetched in the background
        // before the delete runs (deleteSheet waits for captures in flight).
        getSheetSnapshot: (name) => {
          const sheet = sheetIndexOf(name);
          const l = layoutRef.current?.sheets[sheet];
          if (!l) return null;
          const cols = Math.min(l.maxCol, 256);
          const rowsPerWin = cols > 0 ? Math.floor(65_536 / cols) : 0;
          const cells: Promise<EngineWindowCell[]> = (async () => {
            const out: EngineWindowCell[] = [];
            if (cols === 0) return out;
            for (let r0 = 1; r0 <= l.maxRow && out.length < 200_000; r0 += rowsPerWin) {
              const win = await engine.fetchWindow(sheet, r0, Math.min(l.maxRow, r0 + rowsPerWin - 1), 1, cols);
              if (win && Array.isArray(win.cells)) out.push(...win.cells);
            }
            return out;
          })();
          capturesRef.current.push(cells);
          void cells.finally(() => { capturesRef.current = capturesRef.current.filter((p) => p !== cells); });
          return { viewport: true, name, cells };
        },
        restoreSheetSnapshot: (snap) => {
          if (!snap || !snap.viewport || typeof snap.name !== "string") return false;
          real.createSheet!(snap.name);
          void Promise.resolve(snap.cells as Promise<EngineWindowCell[]>).then((cells) => {
            const writes = cells
              .filter((c) => c.f !== undefined || (c.v !== null && c.v !== undefined))
              .map((c) => ({ row: c.r - 1, col: c.c - 1, value: (c.f ?? (typeof c.v === "boolean" ? (c.v ? "TRUE" : "FALSE") : String(c.v))) as string }));
            for (let i = 0; i < writes.length; i += 20_000) writeCells(snap.name, writes.slice(i, i + 20_000), "restore sheet", "skip");
          });
          return true;
        },
        getUsedExtent: (sheetName) => {
          const s = layoutRef.current?.sheets[sheetIndexOf(sheetName)];
          return s ? { lastRow: s.usedRow ?? s.maxRow, lastCol: s.usedCol ?? s.maxCol } : null;
        },
        getPreviewSnapshot: async () => {
          const layout = layoutRef.current;
          if (!layout) return null;
          const snapshot: any = { sheetOrder: [], sheets: {} };
          const data = new Map<string, Map<number, EngineWindowCell>>();
          for (let i = 0; i < layout.sheets.length; i++) {
            const s = layout.sheets[i];
            const cells = new Map<number, EngineWindowCell>();
            // The preview region: the top-left of the used range, one
            // window (≤ 65,536 cells). The capture's byte budget truncates
            // well inside this on dense sheets; the rest of the sheet shows
            // as the row-label index, as on the Univer path.
            const cols = Math.min(s.maxCol, PREVIEW_COLS);
            const rows = Math.min(s.maxRow, s.hidden ? PREVIEW_ROWS_HIDDEN : PREVIEW_ROWS, cols > 0 ? Math.floor(65_536 / cols) : 0);
            if (rows > 0 && cols > 0) {
              const win = await engine.fetchWindow(i, 1, rows, 1, cols);
              if (win && Array.isArray(win.cells)) {
                stylesRef.current.merge(win.styles ?? {});
                for (const c of win.cells) cells.set((c.r - 1) * 1_000_000 + (c.c - 1), c);
              }
            }
            if (s.maxRow > rows && cols > 0) {
              const labelCols = Math.min(cols, PREVIEW_LABEL_COLS);
              const lastRow = Math.min(s.maxRow, rows + Math.floor(65_536 / labelCols));
              const win = await engine.fetchWindow(i, rows + 1, lastRow, 1, labelCols);
              if (win && Array.isArray(win.cells)) {
                stylesRef.current.merge(win.styles ?? {});
                for (const c of win.cells) cells.set((c.r - 1) * 1_000_000 + (c.c - 1), c);
              }
            }
            // Cells changed this session outside the region: fetch their
            // tiles and keep just those cells, so the delta sees them.
            const dirty = dirtyRef.current.get(i);
            if (dirty) {
              const tiles = new Map<string, Set<number>>();
              for (const k of dirty) {
                const r = Math.floor(k / 1_000_000), c = k % 1_000_000;
                if (r < rows && c < cols) continue;
                const key = `${Math.floor(r / TILE_ROWS)}:${Math.floor(c / TILE_COLS)}`;
                let set = tiles.get(key);
                if (!set) tiles.set(key, (set = new Set()));
                set.add(k);
              }
              for (const [key, keys] of tiles) {
                const [tr, tc] = key.split(":").map(Number);
                const r0 = tr * TILE_ROWS + 1, c0 = tc * TILE_COLS + 1;
                const win = await engine.fetchWindow(i, r0, r0 + TILE_ROWS - 1, c0, c0 + TILE_COLS - 1);
                if (!win || !Array.isArray(win.cells)) continue;
                stylesRef.current.merge(win.styles ?? {});
                for (const c of win.cells) {
                  const k = (c.r - 1) * 1_000_000 + (c.c - 1);
                  if (keys.has(k)) cells.set(k, c);
                }
              }
            }
            const cellData: Record<string, Record<string, any>> = {};
            for (const [k, c] of cells) {
              const r = Math.floor(k / 1_000_000), col = k % 1_000_000;
              const entry: any = { v: c.v ?? null };
              if (c.f !== undefined) entry.f = stripImplicitIntersection(c.f);
              if (c.s !== undefined) entry.s = c.s;
              (cellData[r] ??= {})[col] = entry;
            }
            const g = geom(i);
            const id = `sheet_${i}`;
            snapshot.sheetOrder.push(id);
            // Capacity numbers as the Univer snapshot mapper reports them.
            snapshot.sheets[id] = {
              name: s.name,
              hidden: s.hidden ? 1 : 0,
              rowCount: Math.max(10_000, s.maxRow + 100),
              columnCount: Math.max(200, s.maxCol + 20),
              usedRange: (s.usedRow ?? s.maxRow) > 0 && (s.usedCol ?? s.maxCol) > 0 ? `A1:${colLetters((s.usedCol ?? s.maxCol) - 1)}${s.usedRow ?? s.maxRow}` : null,
              cellData,
            };
            data.set(s.name, cells);
          }
          const reader: CellReader = {
            getCell: (sheetName, row, col) => {
              const c = data.get(sheetName)?.get(row * 1_000_000 + col);
              if (!c) return null;
              const out: { value: any; display?: string; formula: string | null } = { value: c.v ?? null, formula: c.f ? stripImplicitIntersection(c.f) : null };
              if (c.d !== undefined && c.d !== String(c.v)) out.display = c.d;
              return out;
            },
            getCellFormat: (sheetName, row, col) => {
              const c = data.get(sheetName)?.get(row * 1_000_000 + col);
              return formatOf(c?.s, formatOverlayRef.current.get(sheetName)?.get(`${row},${col}`));
            },
            getExternalPin: () => null,
          };
          return { snapshot, reader };
        },
        getRangeData: (sheetName, sr, sc, er, ec) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return null;
          const values: any[][] = [];
          const formulas: (string | null)[][] = [];
          const displays: (string | null)[][] = [];
          for (let r = sr; r <= er; r++) {
            const vr: any[] = [], fr: (string | null)[] = [], dr: (string | null)[] = [];
            for (let c = sc; c <= ec; c++) {
              const cell = tilesRef.current.cell(sheet, r, c);
              vr.push(cell?.v ?? null);
              fr.push(cell?.f ? stripImplicitIntersection(cell.f) : null);
              dr.push(cell?.d ?? null);
            }
            values.push(vr); formulas.push(fr); displays.push(dr);
          }
          return { values, formulas, displays };
        },
        getWorkbookSnapshot: () => null,
        getActiveSelection: () => {
          const s = selRef.current;
          const name = layoutRef.current?.sheets[activeRef.current]?.name;
          if (!name) return null;
          return { sheet: name, startRow: Math.min(s.r0, s.r1), startCol: Math.min(s.c0, s.c1), endRow: Math.max(s.r0, s.r1), endCol: Math.max(s.c0, s.c1) };
        },
        getActiveSheetName: () => layoutRef.current?.sheets[activeRef.current]?.name ?? null,
        jumpToCell: (sheetName, row, col) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return false;
          if (sheet !== activeRef.current) activateSheet(sheet);
          setSelection(row, col);
          scrollIntoView(row, col);
          return true;
        },
        getColumnWidth: (sheetName, col) => geom(sheetIndexOf(sheetName))?.colWidth(col) ?? null,
        getRowHeight: (sheetName, row) => geom(sheetIndexOf(sheetName))?.rowHeight(row) ?? null,
        setColumnWidth: (sheetName, col, width) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) setSize(sheet, "col", col, width, true);
        },
        setRowHeight: (sheetName, row, height) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) setSize(sheet, "row", row, height, true);
        },
        hideRows: (sheetName: string, rows: number[]) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) for (const r of rows) setSize(sheet, "row", r, 0, false);
        },
        showRows: (sheetName: string, rows: number[]) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) for (const r of rows) setSize(sheet, "row", r, DEFAULT_ROW_PX, false);
        },
        hideColumns: (sheetName: string, cols: number[]) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) for (const c of cols) setSize(sheet, "col", c, 0, false);
        },
        showColumns: (sheetName: string, cols: number[]) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) for (const c of cols) setSize(sheet, "col", c, DEFAULT_COL_PX, false);
        },
        mergeCells: (sheetName, sr, sc, er, ec) => {
          const sheet = sheetIndexOf(sheetName);
          const l = layoutRef.current?.sheets[sheet];
          if (!l) return;
          const list = currentMerges(sheet).filter((m) => !rangesOverlap(m, sr, sc, er, ec));
          if (sr !== er || sc !== ec) list.push(`${colLetters(sc)}${sr + 1}:${colLetters(ec)}${er + 1}`);
          setMerges(sheet, list, "record");
        },
        unmergeCells: (sheetName, sr, sc, er, ec) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return;
          setMerges(sheet, currentMerges(sheet).filter((m) => !rangesOverlap(m, sr, sc, er, ec)), "record");
        },
        freezePanes: (sheetName, rows, cols) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) setFrozen(sheet, Math.max(0, rows), Math.max(0, cols), "record");
        },
        unfreezePanes: (sheetName) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet >= 0) setFrozen(sheet, 0, 0, "record");
        },
        setRangeBorders: (sheetName, sr, sc, er, ec, kind) => {
          if ((er - sr + 1) * (ec - sc + 1) > 200_000) return;
          let per = borderOverlayRef.current.get(sheetName);
          if (!per) borderOverlayRef.current.set(sheetName, (per = new Map()));
          for (let r = sr; r <= er; r++) {
            for (let c = sc; c <= ec; c++) {
              const b = kind === "none"
                ? { top: false, bottom: false, left: false, right: false }
                : kind === "all"
                  ? { top: true, bottom: true, left: true, right: true }
                  : { top: r === sr, bottom: r === er, left: c === sc, right: c === ec };
              per.set(`${r},${c}`, kind === "outer" ? { ...(per.get(`${r},${c}`) ?? {}), ...Object.fromEntries(Object.entries(b).filter(([, v]) => v)) } : b);
            }
          }
          bumpVersion();
          scheduleDraw();
        },
        setNote: (sheetName, row, col, text) => {
          let per = notesRef.current.get(sheetName);
          if (!per) notesRef.current.set(sheetName, (per = new Map()));
          per.set(`${row},${col}`, text);
          scheduleDraw();
        },
        deleteNote: (sheetName, row, col) => {
          notesRef.current.get(sheetName)?.delete(`${row},${col}`);
          scheduleDraw();
        },
        getNote: (sheetName, row, col) => notesRef.current.get(sheetName)?.get(`${row},${col}`) ?? null,
        toggleFilter: (sheetName, sr, sc, er, ec) => {
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return false;
          return toggleFilterRange(sheet, sr, sc, er, ec);
        },
        sortRange: (sheetName, sr, sc, er, ec, keyCol, direction) => {
          if ((er - sr + 1) * (ec - sc + 1) > 100_000) return;
          void (async () => {
            await prefetch(sheetIndexOf(sheetName), sr, er, sc, ec);
            const sheet = sheetIndexOf(sheetName);
            if (sheet < 0) return;
            const rows: Array<{ key: unknown; cells: Array<string | null> }> = [];
            for (let r = sr; r <= er; r++) {
              const cells: Array<string | null> = [];
              for (let c = sc; c <= ec; c++) {
                const cell = tilesRef.current.cell(sheet, r, c);
                cells.push(cell ? (cell.f ?? (rawInput(cell) || null)) : null);
              }
              const k = tilesRef.current.cell(sheet, r, keyCol);
              rows.push({ key: k?.v ?? null, cells });
            }
            const rank = (v: unknown) => (v === null || v === undefined || v === "" ? 2 : typeof v === "number" ? 0 : 1);
            rows.sort((a, b) => {
              const ra = rank(a.key), rb = rank(b.key);
              if (ra !== rb) return ra - rb; // numbers, then text, blanks last
              if (ra === 2) return 0;
              const cmp = ra === 0 ? (a.key as number) - (b.key as number) : String(a.key).localeCompare(String(b.key), undefined, { sensitivity: "base" });
              return direction === "desc" ? -cmp : cmp;
            });
            const writes: Array<{ row: number; col: number; value: string | null }> = [];
            rows.forEach((row, i) => row.cells.forEach((v, j) => writes.push({ row: sr + i, col: sc + j, value: v })));
            writeCells(sheetName, writes, "sort");
            onUserEdit?.();
          })();
        },
        defineName: (name, ref) => {
          namesRef.current.set(name, ref);
          // Every formula may have changed meaning: refetch what is loaded.
          void engine.defineName(name, ref).then((ok) => {
            if (!ok) return;
            tilesRef.current.clear();
            bumpVersion();
            scheduleDraw();
          });
        },
        deleteName: (name) => {
          namesRef.current.delete(name);
        },
        getDefinedNameRef: (name) => namesRef.current.get(name) ?? layoutRef.current?.definedNames.find((d) => d.name === name)?.formula ?? null,
        setCellFormat: (sheetName, row, col, format) => applyFormat(sheetName, row, col, row, col, format),
        setRangeFormat: (sheetName, sr, sc, er, ec, format) => applyFormat(sheetName, sr, sc, er, ec, format),
        getCellFormat: (sheetName, row, col) => {
          const sheet = sheetIndexOf(sheetName);
          const cell = sheet >= 0 ? tilesRef.current.cell(sheet, row, col) : undefined;
          return formatOf(cell?.s, sheet >= 0 ? overlayOf(sheet, row, col) : undefined);
        },
        clearRange: (sheetName, sr, sc, er, ec) => {
          const cells: Array<{ row: number; col: number; value: null }> = [];
          for (let r = sr; r <= er; r++) for (let c = sc; c <= ec; c++) cells.push({ row: r, col: c, value: null });
          if (cells.length === 0 || cells.length > 50_000) return;
          writeCells(sheetName, cells, "clear");
        },
        getFreezePanes: (sheetName: string) => {
          const g = geom(sheetIndexOf(sheetName));
          return g ? { freezeRows: g.frozenRows, freezeCols: g.frozenCols } : { freezeRows: 0, freezeCols: 0 };
        },
        createSheet: (name) => {
          // A sheet exists from the moment it is requested: the layout only
          // refreshes after the engine settles, and a second create for the
          // same name in that window would fail the save (the patcher
          // refuses duplicate sheets).
          const known = new Set([...(layoutRef.current?.sheets.map((s) => s.name) ?? []), ...pendingSheetsRef.current]);
          if (known.has(name)) return false;
          pendingSheetsRef.current.add(name);
          engine.sheetOp("add", name);
          // Not reported as a manual op: agent creates reach the save
          // through their batch, and a second create op makes the patcher
          // refuse the save. The tab menu reports the user's own inserts.
          void engine.settled().then(reloadLayout);
          return true;
        },
        deleteSheet: (name) => {
          // Let a snapshot capture of this sheet finish before the delete is queued.
          void Promise.all(capturesRef.current).then(() => engine.sheetOp("delete", name));
          void engine.settled().then(reloadLayout);
          return true;
        },
        renameSheet: (oldName: string, newName: string) => {
          engine.sheetOp("rename", oldName, newName);
          void engine.settled().then(reloadLayout);
          return true;
        },
        insertRows: (sheetName: string, at: number, count: number) => structural("insert_rows", sheetName, at + 1, count),
        deleteRows: (sheetName: string, at: number, count: number) => structural("delete_rows", sheetName, at + 1, count),
        insertColumns: (sheetName: string, at: number, count: number) => structural("insert_columns", sheetName, at + 1, count),
        deleteColumns: (sheetName: string, at: number, count: number) => structural("delete_columns", sheetName, at + 1, count),
        getSourceBytes: () => sourceBytesRef.current,
        syncEditLog: async () => {
          // The edit log keeps each formula as it was written, keyed by
          // where the cell is now; rows/columns inserted or deleted since
          // have rewritten its references in the engine. The patch must
          // carry the engine's text, or the file reopens with the old refs.
          await engine.settled(60_000);
          const live = (written: string, now: string) => {
            if (now === written || stripImplicitIntersection(now) === stripImplicitIntersection(written)) return written;
            return written.includes("@") ? now : stripImplicitIntersection(now);
          };
          for (const [name, cells] of editsRef.current) {
            const sheet = sheetIndexOf(name);
            if (sheet < 0) continue;
            const formulas = new Map([...cells].filter(([, e]) => e.f !== undefined));
            if (formulas.size === 0) continue;
            let r0 = Infinity, r1 = -1, c0 = Infinity, c1 = -1;
            for (const k of formulas.keys()) {
              const [r, c] = k.split(",").map(Number);
              r0 = Math.min(r0, r); r1 = Math.max(r1, r); c0 = Math.min(c0, c); c1 = Math.max(c1, c);
            }
            const take = (win: any) => {
              for (const w of win?.cells ?? []) {
                const e = formulas.get(`${w.r - 1},${w.c - 1}`);
                if (e?.f !== undefined && typeof w.f === "string" && w.f) e.f = live(e.f, w.f);
              }
            };
            const width = c1 - c0 + 1;
            if ((r1 - r0 + 1) * width <= 200_000) {
              const rowsPer = Math.max(1, Math.floor(65_536 / width));
              for (let a = r0; a <= r1; a += rowsPer) take(await engine.fetchWindow(sheet, a + 1, Math.min(r1, a + rowsPer - 1) + 1, c0 + 1, c1 + 1));
            } else {
              for (const k of formulas.keys()) {
                const [r, c] = k.split(",").map(Number);
                take(await engine.fetchWindow(sheet, r + 1, r + 1, c + 1, c + 1));
              }
            }
          }
        },
        setHighlights: (cells: Array<{ sheet: string; row: number; col: number }> | null) => {
          const next = new Map<string, Set<number>>();
          for (const c of cells ?? []) {
            let set = next.get(c.sheet);
            if (!set) next.set(c.sheet, (set = new Set()));
            set.add(c.row * 1_000_000 + c.col);
          }
          highlightsRef.current = next;
          scheduleDraw();
        },
        getFidelityRisks: () => null,
        commitSavedBaseline: async (bytes) => {
          sourceBytesRef.current = bytes;
          editsRef.current.clear();
          formatEditsRef.current.clear();
          sizeEditsRef.current = { cols: new Map(), rows: new Map() };
          manualRowColOpsRef.current = [];
        },
        exportPatch: (mirror) => {
          if (!sourceBytesRef.current) {
            return { ok: false, reason: "missing_baseline", detail: "no workbook loaded" };
          }
          // The shared patch builder diffs a Univer-shaped snapshot against a
          // baseline and folds in the mirror (formats, widths, sheet and
          // row/column ops, names). Here the "snapshot" is the session's
          // edit log and the baseline marks the cleared cells, so the diff
          // is exactly the cells this session wrote.
          const snapshot: any = { sheets: {} };
          const baseline = new Map<string, Map<string, BaselineCell>>();
          let i = 0;
          for (const [name, cells] of editsRef.current) {
            const cellData: Record<string, Record<string, any>> = {};
            const base = new Map<string, BaselineCell>();
            for (const [key, e] of cells) {
              const [r, c] = key.split(",");
              if (e.f === undefined && e.v === undefined) {
                base.set(key, { v: "\u0000cleared" });
                continue;
              }
              (cellData[r] ??= {})[c] = e.f !== undefined ? { f: e.f } : { v: e.v };
            }
            snapshot.sheets[`s${i++}`] = { name, cellData };
            baseline.set(name, base);
          }
          // Only edited sheets appear in both maps: a baseline sheet absent
          // from the snapshot would read as deleted to the builder.
          // Header drags and agent width/height ops ride the mirror's lists.
          const columnWidths = [...(mirror?.columnWidths ?? [])];
          for (const [sheet, per] of sizeEditsRef.current.cols) for (const [col, widthPx] of per) columnWidths.push({ sheet, col, widthPx });
          const rowHeights = [...(mirror?.rowHeights ?? [])];
          for (const [sheet, per] of sizeEditsRef.current.rows) for (const [row, heightPx] of per) rowHeights.push({ sheet, row, heightPx });
          const rowColOps = [...(mirror?.rowColOps ?? []), ...manualRowColOpsRef.current] as NonNullable<typeof mirror>["rowColOps"];
          const autoFilters = [...(mirror?.autoFilters ?? [])];
          for (const [sheet, e] of autoFilterEditsRef.current) autoFilters.push({ sheet, range: e.range, columns: e.columns });
          // The grid's own format log goes last: it is the net of every
          // change applied here (including the mirror's), so the live state wins.
          const cellFormats = [...(mirror?.cellFormats ?? [])];
          for (const [sheet, per] of formatEditsRef.current) {
            for (const [key, format] of per) {
              const [row, col] = key.split(",").map(Number);
              cellFormats.push({ sheet, row, col, format });
            }
          }
          const merged = { ...(mirror ?? {}), ...(cellFormats.length ? { cellFormats } : {}), ...(columnWidths.length ? { columnWidths } : {}), ...(rowHeights.length ? { rowHeights } : {}), ...(rowColOps && rowColOps.length ? { rowColOps } : {}), ...(autoFilters.length ? { autoFilters } : {}) };
          return buildWorkbookPatch(snapshot, baseline, merged);
        },
        exportBytes: async () => {
          throw new Error("Full export is not available in the viewport grid; the in-place save is used instead.");
        },
        containsFocus: () => !!scrollerRef.current && scrollerRef.current.contains(document.activeElement),
        setRenderActive: () => {},
        undo,
        redo,
        captureCellBand: (sheetName: string, sr: number, sc: number, er: number, ec: number) => {
          const sheet = sheetIndexOf(sheetName);
          const out: Array<{ row: number; col: number; value: any; formula: string | null; format: any }> = [];
          if (sheet < 0) return out;
          for (let r = sr; r <= er; r++) for (let c = sc; c <= ec; c++) {
            const cell = tilesRef.current.cell(sheet, r, c);
            out.push({ row: r, col: c, value: cell?.v ?? null, formula: cell?.f ? stripImplicitIntersection(cell.f) : null, format: {} });
          }
          return out;
        },
      };
      const FORMAT_DEFAULTS: Record<string, unknown> = { bold: false, italic: false, underline: false, strike: false, font_color: "#000000", background_color: null, font_size: 11, font_family: "Calibri", horizontal_align: "general", vertical_align: "bottom", number_format: "General", wrap_text: false, indent: 0 };
      const applyFormat = (sheetName: string, sr: number, sc: number, er: number, ec: number, format: CellFormatShape, history: "record" | "skip" = "record") => {
        if ((er - sr + 1) * (ec - sc + 1) > 200_000) return;
        if (history === "record") {
          // Before-formats of the loaded cells, keyed by their JSON so a
          // uniform range undoes in one call; mixed ranges per cell (capped).
          const sheet = sheetIndexOf(sheetName);
          const keys = Object.keys(format) as Array<keyof CellFormatShape>;
          const restoreOf = (before: CellFormatShape): CellFormatShape => {
            const out: Record<string, unknown> = {};
            for (const k of keys) out[k] = before[k] !== undefined ? before[k] : FORMAT_DEFAULTS[k as string];
            return out as CellFormatShape;
          };
          const groups = new Map<string, Array<[number, number]>>();
          let captured = 0;
          if (sheet >= 0) {
            outer: for (let r = sr; r <= er; r++) {
              for (let c = sc; c <= ec; c++) {
                const cell = tilesRef.current.cell(sheet, r, c);
                if (!cell && !tilesRef.current.has(tileKey(sheet, Math.floor(r / TILE_ROWS), Math.floor(c / TILE_COLS)))) continue;
                const restore = restoreOf(formatOf(cell?.s, overlayOf(sheet, r, c)));
                const key = JSON.stringify(restore);
                let list = groups.get(key);
                if (!list) groups.set(key, (list = []));
                list.push([r, c]);
                if (++captured >= 5000) break outer;
              }
            }
          }
          if (groups.size > 0) {
            const uniform = groups.size === 1 && captured === (er - sr + 1) * (ec - sc + 1);
            pushUndo({
              label: "format",
              undo: () => {
                if (uniform) applyFormat(sheetName, sr, sc, er, ec, JSON.parse([...groups.keys()][0]), "skip");
                else for (const [key, cells] of groups) { const f = JSON.parse(key); for (const [r, c] of cells) applyFormat(sheetName, r, c, r, c, f, "skip"); }
              },
              redo: () => applyFormat(sheetName, sr, sc, er, ec, format, "skip"),
            });
          }
        }
        // Immediate: the overlay draws the change now. Durable for display:
        // the engine takes the style too, so display strings (number
        // formats) and the style table reflect it; the touched tiles refetch.
        let per = formatOverlayRef.current.get(sheetName);
        if (!per) formatOverlayRef.current.set(sheetName, (per = new Map()));
        let saved = formatEditsRef.current.get(sheetName);
        if (!saved) formatEditsRef.current.set(sheetName, (saved = new Map()));
        for (let r = sr; r <= er; r++) {
          for (let c = sc; c <= ec; c++) {
            const key = `${r},${c}`;
            per.set(key, { ...(per.get(key) ?? {}), ...format });
            saved.set(key, { ...(saved.get(key) ?? {}), ...format });
          }
        }
        bumpVersion();
        scheduleDraw();
        void engine.setStyle(sheetName, sr, sc, er, ec, { ...format }).then((ok) => {
          if (!ok) return;
          const sheet = sheetIndexOf(sheetName);
          if (sheet < 0) return;
          const overlay = formatOverlayRef.current.get(sheetName);
          for (let r = sr; r <= er; r++) for (let c = sc; c <= ec; c++) overlay?.delete(`${r},${c}`);
          const seen = new Set<string>();
          for (let tr = Math.floor(sr / TILE_ROWS); tr <= Math.floor(er / TILE_ROWS); tr++) {
            for (let tc = Math.floor(sc / TILE_COLS); tc <= Math.floor(ec / TILE_COLS); tc++) {
              const key = tileKey(sheet, tr, tc);
              if (seen.has(key) || !tilesRef.current.has(key)) continue;
              seen.add(key);
              void fetchTile(sheet, tr, tc).then(scheduleDraw);
            }
          }
        });
      };
      const structural = (op: string, sheetName: string, index1: number, count: number, history: "record" | "skip" = "record") => {
        structuralRef.current = structural;
        if (history === "record") {
          const inverse: Record<string, string> = { insert_rows: "delete_rows", delete_rows: "insert_rows", insert_columns: "delete_columns", delete_columns: "insert_columns" };
          const isDelete = op.startsWith("delete");
          const rows = op.endsWith("rows");
          // A delete loses content: keep what the loaded tiles hold so undo can put it back.
          const band: Array<{ row: number; col: number; value: string | null }> = [];
          const sheet = sheetIndexOf(sheetName);
          const l = layoutRef.current?.sheets[sheet];
          if (isDelete && sheet >= 0 && l) {
            const r0 = rows ? index1 - 1 : 0, r1 = rows ? index1 - 2 + count : Math.max(0, l.maxRow - 1);
            const c0 = rows ? 0 : index1 - 1, c1 = rows ? Math.max(0, l.maxCol - 1) : index1 - 2 + count;
            for (let r = r0; r <= r1 && band.length < 50_000; r++) for (let c = c0; c <= c1; c++) {
              const cell = tilesRef.current.cell(sheet, r, c);
              if (cell && (cell.f !== undefined || (cell.v !== null && cell.v !== undefined))) band.push({ row: r, col: c, value: cell.f ?? rawInput(cell) });
            }
          }
          pushUndo({
            label: op,
            undo: () => {
              structural(inverse[op], sheetName, index1, count, "skip");
              if (band.length) writeCells(sheetName, band, "undo " + op, "skip");
            },
            redo: () => structural(op, sheetName, index1, count, "skip"),
          });
        }
        engine.structural(op, sheetName, index1, count);
        shiftSession(sheetName, op, index1 - 1, count);
        void engine.settled().then(reloadLayout);
        return true;
      };
      /** The edit log, format overlay, dirty set and size overrides are
       * keyed by coordinates; a row/column insert or delete shifts them the
       * way the engine shifted the cells (0-based `at`). */
      const shiftSession = (sheetName: string, op: string, at: number, count: number) => {
        const rows = op === "insert_rows" || op === "delete_rows";
        const insert = op === "insert_rows" || op === "insert_columns";
        const remap = (r: number, c: number): [number, number] | null => {
          const i = rows ? r : c;
          if (i < at) return [r, c];
          if (insert) return rows ? [r + count, c] : [r, c + count];
          if (i < at + count) return null;
          return rows ? [r - count, c] : [r, c - count];
        };
        const shiftKeyed = <T,>(m: Map<string, T> | undefined) => {
          if (!m) return;
          const next = new Map<string, T>();
          for (const [k, v] of m) {
            const [r, c] = k.split(",").map(Number);
            const to = remap(r, c);
            if (to) next.set(`${to[0]},${to[1]}`, v);
          }
          m.clear();
          for (const [k, v] of next) m.set(k, v);
        };
        shiftKeyed(editsRef.current.get(sheetName));
        shiftKeyed(formatOverlayRef.current.get(sheetName));
        shiftKeyed(formatEditsRef.current.get(sheetName));
        const sheet = sheetIndexOf(sheetName);
        if (sheet >= 0) {
          const dirty = dirtyRef.current.get(sheet);
          if (dirty) {
            const next = new Set<number>();
            for (const k of dirty) {
              const to = remap(Math.floor(k / 1_000_000), k % 1_000_000);
              if (to) next.add(to[0] * 1_000_000 + to[1]);
            }
            dirtyRef.current.set(sheet, next);
          }
          const o = sizeOverridesRef.current.get(sheet);
          if (o) {
            const src = rows ? o.rows : o.cols;
            const next = new Map<number, number>();
            for (const [i, px] of src) {
              const to = rows ? remap(i, 0) : remap(0, i);
              if (to) next.set(rows ? to[0] : to[1], px);
            }
            if (rows) o.rows = next; else o.cols = next;
          }
          tilesRef.current.dropSheet(sheet);
          geomsRef.current.delete(sheet);
        }
      };
      const prefetch = async (sheet: number, r0: number, r1: number, c0: number, c1: number) => {
        const g = geom(sheet);
        if (!g) return;
        const rr0 = Math.max(0, r0), rr1 = Math.min(g.rowCount - 1, r1), cc0 = Math.max(0, c0), cc1 = Math.min(g.colCount - 1, c1);
        const missing = tilesRef.current.missing(sheet, rr0, rr1, cc0, cc1);
        await Promise.all(missing.map((t) => fetchTile(sheet, t.tr, t.tc)));
        // Tiles another fetch already has in flight: wait for them too, so a
        // read right after a write sees the refreshed tile.
        const t0 = performance.now();
        while (tilesRef.current.pendingIn(sheet, rr0, rr1, cc0, cc1).length > 0 && performance.now() - t0 < 10_000) {
          await new Promise((r) => setTimeout(r, 10));
        }
        scheduleDraw();
      };
      // Anything not implemented above is a no-op that returns a safe empty
      // value, so the rest of the screen runs unchanged in big-model mode.
      structuralRef.current = structural;
      applyFormatRef.current = applyFormat;
      prefetchRef.current = prefetch;
      return new Proxy(real as UniverGridHandle, {
        get(target, prop: string) {
          if (prop in target) return (target as any)[prop];
          if (prop === "then") return undefined;
          return (..._args: unknown[]) => {
            if (/^(get|is|contains)/.test(prop)) return null;
            return false;
          };
        },
      });
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [tabId]);

    useImperativeHandle(ref, () => handle, [handle]);

    // ----- editor overlay position --------------------------------------

    const editorStyle = useMemo<React.CSSProperties | null>(() => {
      if (!editing) return null;
      const g = geom(activeRef.current);
      const sc = scrollerRef.current;
      if (!g || !sc) return null;
      const z = viewScale();
      // Content coordinates inside the scroller: sheet px × zoom; a frozen
      // row/column stays put, so it also moves with the scroll offset.
      const left = (HEADER_W + g.colLeft[editing.col]) * z + (editing.col < g.frozenCols ? sc.scrollLeft : 0);
      const top = (HEADER_H + g.rowTop[editing.row]) * z + (editing.row < g.frozenRows ? sc.scrollTop : 0);
      return { left, top, width: Math.max(g.colWidth(editing.col), DEFAULT_COL_PX) * z + 2, height: g.rowHeight(editing.row) * z + 2, fontSize: 13 * z };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [editing, zoom]);

    const lastFindQuery = useRef("");

    return (
      <Root data-vp-root="1">
        <FormulaBar>
          <NameBox>{nameBox}</NameBox>
          <FormulaInput
            value={editing ? editing.text : formulaText}
            ref={barRef}
            onChange={(e) => { refPick.current = null; editing ? setEditing({ ...editing, text: e.target.value }) : startEditFromBar(e.target.value); }}
            onKeyDown={(e) => {
              if (!editing) return;
              if (e.key === "Enter") { e.preventDefault(); commitEdit(editing.text, "down"); }
              if (e.key === "Escape") { e.preventDefault(); editingRef.current = null; setEditing(null); refreshFormulaBar(); }
            }}
            spellCheck={false}
          />
          <Status>
            <span>{calculating ? "calculating… " : ""}{stats || status}</span>
            <ZoomCtl>
              <button onClick={() => setZoomLevel(zoomRef.current - 0.1)} title="Zoom out (Cmd Option -)">−</button>
              <button onClick={() => setZoomLevel(1)} title="Reset zoom (Cmd Option 0)">{Math.round(zoom * 100)}%</button>
              <button onClick={() => setZoomLevel(zoomRef.current + 0.1)} title="Zoom in (Cmd Option +)">+</button>
            </ZoomCtl>
          </Status>
        </FormulaBar>
        <GridArea ref={gridAreaRef}>
        <Canvas ref={canvasRef} style={{ left: gutter.w, top: gutter.h }} />
        <Scroller
          ref={scrollerRef}
          style={{ left: gutter.w, top: gutter.h }}
          tabIndex={0}
          onMouseDown={onMouseDown}
          onMouseMove={onMouseMove}
          onMouseUp={onMouseUp}
          onDoubleClick={onDoubleClick}
          onKeyDown={onKeyDown}
          onContextMenu={(e) => {
            e.preventDefault();
            if (editing) return;
            const root = (e.currentTarget as HTMLElement).parentElement!.getBoundingClientRect();
            const hh = headerHit(e.clientX, e.clientY);
            const g = geom(activeRef.current);
            const s = selRef.current;
            if (hh && g) {
              // Right-click on a header: select that row/column unless it is inside the selection.
              const inside = hh.kind === "col" ? hh.index >= Math.min(s.c0, s.c1) && hh.index <= Math.max(s.c0, s.c1) && Math.min(s.r0, s.r1) === 0 && Math.max(s.r0, s.r1) === g.rowCount - 1
                : hh.index >= Math.min(s.r0, s.r1) && hh.index <= Math.max(s.r0, s.r1) && Math.min(s.c0, s.c1) === 0 && Math.max(s.c0, s.c1) === g.colCount - 1;
              if (!inside) {
                selRef.current = hh.kind === "col" ? { r0: 0, c0: hh.index, r1: g.rowCount - 1, c1: hh.index, ar: 0, ac: hh.index } : { r0: hh.index, c0: 0, r1: hh.index, c1: g.colCount - 1, ar: hh.index, ac: 0 };
                refreshFormulaBar();
                scheduleDraw();
              }
              setMenu({ x: e.clientX - root.left, y: e.clientY - root.top, header: hh.kind });
              return;
            }
            const hit = cellAtPoint(e.clientX, e.clientY);
            if (hit) {
              const inside = hit.r >= Math.min(s.r0, s.r1) && hit.r <= Math.max(s.r0, s.r1) && hit.c >= Math.min(s.c0, s.c1) && hit.c <= Math.max(s.c0, s.c1);
              if (!inside) setSelection(hit.r, hit.c);
            }
            setMenu({ x: e.clientX - root.left, y: e.clientY - root.top });
          }}
          onPaste={(e) => {
            if (editing) return;
            const text = e.clipboardData.getData("text/plain");
            if (!text) return;
            e.preventDefault();
            pasteText(text, e.clipboardData.getData("text/html") || undefined);
          }}
        >
          <Spacer ref={spacerRef} />
          {editing && editorStyle && (
            <CellEditor
              ref={editorRef}
              style={editorStyle}
              value={editing.text}
              onChange={(e) => { refPick.current = null; setEditing({ ...editing, text: e.target.value }); }}
              onKeyDown={(e) => {
                if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) { e.preventDefault(); commitToSelection(editing.text); }
                else if (e.key === "Enter") { e.preventDefault(); commitEdit(editing.text, e.shiftKey ? "up" : "down"); }
                else if (e.key === "Tab") { e.preventDefault(); commitEdit(editing.text, e.shiftKey ? "left" : "right"); }
                else if (e.key === "Escape") { e.preventDefault(); editingRef.current = null; setEditing(null); refreshFormulaBar(); scrollerRef.current?.focus(); }
                e.stopPropagation();
              }}
              onBlur={(ev) => { if (editingRef.current && !refDrag.current && !refPick.current) commitEdit((ev.target as HTMLInputElement).value, "none"); }}
            />
          )}
        </Scroller>
        <OutlineCanvas ref={outlineCanvasRef} />
        {gutter.w > 0 && <OutlineStrip data-vp-outline="rows" style={{ left: 0, top: 0, bottom: 0, width: gutter.w }} onMouseDown={onOutlineMouseDown} />}
        {gutter.h > 0 && <OutlineStrip data-vp-outline="cols" style={{ left: gutter.w, top: 0, right: 0, height: gutter.h }} onMouseDown={onOutlineMouseDown} />}
        </GridArea>
        {find && (
          <FindBar onMouseDown={(e) => e.stopPropagation()}>
            <input
              ref={findInputRef}
              value={find.query}
              placeholder="Find…"
              onChange={(e) => setFind({ ...find, query: e.target.value })}
              onKeyDown={(e) => {
                if (e.key === "Enter") { e.preventDefault(); if (find.hits.length && find.query === lastFindQuery.current) stepFind(e.shiftKey ? -1 : 1); else { lastFindQuery.current = find.query; void runFind(find.query, find.scope); } }
                else if (e.key === "Escape") { e.preventDefault(); setFind(null); scrollerRef.current?.focus(); }
                e.stopPropagation();
              }}
            />
            <span>{find.hits.length ? `${find.at + 1} of ${find.hits.length}${find.hits.length >= 500 ? "+" : ""}` : find.query && find.query === lastFindQuery.current ? "no matches" : ""}</span>
            <button onClick={() => stepFind(-1)} title="Previous (Shift+Enter)">↑</button>
            <button onClick={() => stepFind(1)} title="Next (Enter)">↓</button>
            <button onClick={() => { const scope = find.scope === "sheet" ? "workbook" : "sheet"; lastFindQuery.current = find.query; void runFind(find.query, scope); }} title="Search this sheet or the whole workbook">{find.scope === "sheet" ? "sheet" : "workbook"}</button>
            <input
              value={find.replace}
              placeholder="Replace with…"
              onChange={(e) => setFind({ ...find, replace: e.target.value })}
              onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); void replaceHits(false); } else if (e.key === "Escape") { e.preventDefault(); setFind(null); scrollerRef.current?.focus(); } e.stopPropagation(); }}
            />
            <button disabled={!find.hits.length} onClick={() => void replaceHits(false)}>Replace</button>
            <button disabled={!find.hits.length} onClick={() => void replaceHits(true)}>All</button>
            <button onClick={() => { setFind(null); scrollerRef.current?.focus(); }}>✕</button>
          </FindBar>
        )}
        {tabMenu && (
          <Menu data-vp-menu="1" style={{ left: tabMenu.x, top: tabMenu.y }} onMouseDown={(e) => e.stopPropagation()}>
            <MenuItem onClick={() => { setTabMenu(null); const base = "Sheet"; let n = sheetNames.length + 1; while (sheetNames.some((x) => x.name === `${base}${n}`)) n++; if (handle.createSheet(`${base}${n}`)) onManualSheetOp?.({ kind: "create", name: `${base}${n}` }); }}>Insert sheet</MenuItem>
            <MenuItem onClick={() => { setTabMenu(null); setRenaming({ index: tabMenu.index, text: sheetNames[tabMenu.index]?.name ?? "" }); }}>Rename</MenuItem>
            <MenuItem onClick={() => { setColorPick(tabMenu.index); setTabMenu(null); }}>Tab color…</MenuItem>
            {sheetNames[tabMenu.index]?.hidden ? (
              <MenuItem onClick={() => { setTabMenu(null); setSheetVisibility(tabMenu.index, false); }}>Unhide</MenuItem>
            ) : (
              <MenuItem disabled={sheetNames.filter((x) => !x.hidden).length <= 1} onClick={() => { setTabMenu(null); setSheetVisibility(tabMenu.index, true); }}>Hide</MenuItem>
            )}
            <MenuItem disabled={sheetNames.length <= 1} onClick={() => { setTabMenu(null); const name = sheetNames[tabMenu.index]?.name; if (name && sheetNames.length > 1) { handle.deleteSheet(name); onManualSheetOp?.({ kind: "delete", name }); } }}>Delete</MenuItem>
            {sheetNames.some((x) => x.hidden) && <MenuSep />}
            {sheetNames.some((x) => x.hidden) && (
              <MenuItem onClick={() => { setTabMenu(null); setShowHidden((v) => !v); }}>
                {showHidden ? "Hide hidden sheets" : `Show hidden sheets (${sheetNames.filter((x) => x.hidden).length})`}
              </MenuItem>
            )}
          </Menu>
        )}
        {filterMenu && (() => {
          const fm = filterMenu;
          const fsx = filtersRef.current.get(fm.sheet);
          const name = sheetNameOf(fm.sheet) ?? "";
          const q = fm.search.toLowerCase();
          const shown = q ? fm.items.filter((it) => it.text.toLowerCase().includes(q)) : fm.items;
          const allOn = shown.length > 0 && shown.every((it) => it.checked);
          const noneOn = shown.every((it) => !it.checked);
          const setItems = (f: (it: { text: string; checked: boolean }) => boolean) => setFilterMenu((m) => (m ? { ...m, items: m.items.map((it) => ({ ...it, checked: f(it) })) } : m));
          const ok = () => {
            setFilterMenu(null);
            if (!fsx) return;
            const checked = fm.items.filter((it) => it.checked);
            if (checked.length === fm.items.length) { void applyFilter(fm.sheet, fm.col, null); return; }
            const values = new Set(checked.map((it) => it.text).filter((t) => t !== ""));
            void applyFilter(fm.sheet, fm.col, { values, blank: checked.some((it) => it.text === "") });
          };
          const sortBy = (dir: "asc" | "desc") => { setFilterMenu(null); if (fsx && name) handle.sortRange(name, fsx.r0 + 1, fsx.c0, fsx.r1, fsx.c1, fm.col, dir); };
          return (
            <Menu data-vp-menu="1" style={{ left: fm.x, top: fm.y, width: 250 }} onMouseDown={(e) => e.stopPropagation()}>
              <MenuItem onClick={() => sortBy("asc")}>Sort A → Z</MenuItem>
              <MenuItem onClick={() => sortBy("desc")}>Sort Z → A</MenuItem>
              <MenuSep />
              <MenuItem disabled={!fsx?.columns.has(fm.col)} onClick={() => { setFilterMenu(null); void applyFilter(fm.sheet, fm.col, null); }}>Clear filter from {colLetters(fm.col)}</MenuItem>
              <MenuSep />
              <div style={{ padding: "4px 10px" }}>
                <input data-vp-filter-search="1" placeholder="Search" value={fm.search} autoFocus onChange={(e) => setFilterMenu((m) => (m ? { ...m, search: e.target.value } : m))} onKeyDown={(e) => { if (e.key === "Enter") ok(); e.stopPropagation(); }} style={{ width: "100%", boxSizing: "border-box", font: "inherit", padding: "3px 6px", border: "1px solid #ccc", borderRadius: 4 }} />
              </div>
              <div data-vp-filter-list="1" style={{ maxHeight: 220, overflowY: "auto", padding: "2px 10px 4px", fontSize: 12, lineHeight: "20px" }}>
                {fm.loading ? <div style={{ color: "#888" }}>Loading…</div> : (
                  <>
                    <label style={{ display: "block", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      <input type="checkbox" checked={allOn} ref={(el) => { if (el) el.indeterminate = !allOn && !noneOn; }} onChange={(e) => { const on = e.target.checked; const vis = new Set(shown.map((it) => it.text)); setItems((it) => (vis.has(it.text) ? on : it.checked)); }} /> (Select all{q ? " search results" : ""})
                    </label>
                    {shown.slice(0, 1000).map((it) => (
                      <label key={it.text} style={{ display: "block", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                        <input type="checkbox" checked={it.checked} onChange={(e) => { const on = e.target.checked; setItems((x) => (x.text === it.text ? on : x.checked)); }} /> {it.text === "" ? "(Blanks)" : it.text}
                      </label>
                    ))}
                    {shown.length > 1000 && <div style={{ color: "#888" }}>… {shown.length - 1000} more</div>}
                  </>
                )}
              </div>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", padding: "6px 10px 4px" }}>
                <button onClick={() => setFilterMenu(null)} style={{ font: "inherit", padding: "3px 10px" }}>Cancel</button>
                <button data-vp-filter-ok="1" onClick={ok} style={{ font: "inherit", padding: "3px 10px" }}>OK</button>
              </div>
            </Menu>
          );
        })()}
        {colorPick !== null && (
          <Menu data-vp-menu="1" style={{ left: 120, bottom: 34, padding: 8 }} onMouseDown={(e) => e.stopPropagation()}>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 22px)", gap: 6 }}>
              {TAB_COLORS.map((c) => (
                <button key={c} title={c} onClick={() => { setTabColor(colorPick, c); setColorPick(null); }} style={{ width: 22, height: 22, border: "1px solid #ccc", borderRadius: 4, background: c, cursor: "pointer" }} />
              ))}
            </div>
            <MenuItem style={{ marginTop: 6 }} onClick={() => { setTabColor(colorPick, null); setColorPick(null); }}>No color</MenuItem>
          </Menu>
        )}
        {hoverNote && (
          <Menu style={{ left: hoverNote.x, top: hoverNote.y, padding: "6px 10px", maxWidth: 320, whiteSpace: "pre-wrap", pointerEvents: "none", background: "#fffbe6" }}>
            {hoverNote.text}
          </Menu>
        )}
        {menu && (() => {
          const s = selRef.current;
          const r0 = Math.min(s.r0, s.r1), r1 = Math.max(s.r0, s.r1), c0 = Math.min(s.c0, s.c1), c1 = Math.max(s.c0, s.c1);
          const nR = r1 - r0 + 1, nC = c1 - c0 + 1;
          const name = layoutRef.current?.sheets[activeRef.current]?.name ?? "";
          const close = () => { setMenu(null); scrollerRef.current?.focus(); };
          const rowOp = (kind: "insertRows" | "deleteRows", at: number, count: number) => {
            close();
            manualRowColOpsRef.current.push(kind === "insertRows" ? { kind, sheet: name, before: at, count } : { kind, sheet: name, start: at, count });
            structuralRef.current(kind === "insertRows" ? "insert_rows" : "delete_rows", name, at + 1, count);
            onUserEdit?.();
          };
          const colOp = (kind: "insertColumns" | "deleteColumns", at: number, count: number) => {
            close();
            manualRowColOpsRef.current.push(kind === "insertColumns" ? { kind, sheet: name, before: at, count } : { kind, sheet: name, start: at, count });
            structuralRef.current(kind === "insertColumns" ? "insert_columns" : "delete_columns", name, at + 1, count);
            onUserEdit?.();
          };
          const plural = (n: number, w: string) => `${n} ${w}${n === 1 ? "" : "s"}`;
          return (
            <Menu data-vp-menu="1" style={{ left: menu.x, top: menu.y }} onMouseDown={(e) => e.stopPropagation()}>
              <MenuItem onClick={() => { close(); copySelection(true); }}>Cut</MenuItem>
              <MenuItem onClick={() => { close(); copySelection(false); }}>Copy</MenuItem>
              <MenuItem onClick={async () => { close(); try { const t = await navigator.clipboard.readText(); if (t) pasteText(t); } catch { /* clipboard unavailable */ } }}>Paste</MenuItem>
              <MenuItem onClick={() => { close(); const cells: Array<{ row: number; col: number; value: null }> = []; for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) cells.push({ row: r, col: c, value: null }); if (cells.length <= 50_000 && name) { writeCells(name, cells, "clear"); onUserEdit?.(); } }}>Clear contents</MenuItem>
              <MenuSep />
              {(() => {
                const fsx = filtersRef.current.get(activeRef.current);
                const inside = !!fsx && s.ar >= fsx.r0 && s.ar <= fsx.r1 && s.ac >= fsx.c0 && s.ac <= fsx.c1;
                if (!fsx || !inside) return <MenuItem onClick={() => { close(); toggleFilterAtSelection(); }}>Filter</MenuItem>;
                return (
                  <>
                    <MenuItem disabled={fsx.columns.size === 0} onClick={() => { close(); void (async () => { for (const c of [...fsx.columns.keys()]) await applyFilter(activeRef.current, c, null); })(); }}>Clear filter</MenuItem>
                    <MenuItem onClick={() => { close(); toggleFilterAtSelection(); }}>Remove filter</MenuItem>
                  </>
                );
              })()}
              <MenuSep />
              <MenuItem onClick={() => rowOp("insertRows", r0, nR)}>Insert {plural(nR, "row")} above</MenuItem>
              <MenuItem onClick={() => rowOp("deleteRows", r0, nR)}>Delete {plural(nR, "row")}</MenuItem>
              <MenuItem onClick={() => colOp("insertColumns", c0, nC)}>Insert {plural(nC, "column")} left</MenuItem>
              <MenuItem onClick={() => colOp("deleteColumns", c0, nC)}>Delete {plural(nC, "column")}</MenuItem>
              {menu.header && <MenuSep />}
              {menu.header === "row" && <MenuItem onClick={() => { close(); for (let r = r0; r <= r1; r++) setSize(activeRef.current, "row", r, 0, true); onUserEdit?.(); }}>Hide {plural(nR, "row")}</MenuItem>}
              {menu.header === "row" && <MenuItem onClick={() => { close(); const g = geom(activeRef.current); for (let r = Math.max(0, r0 - 1); r <= Math.min((g?.rowCount ?? 0) - 1, r1 + 1); r++) if (g && g.rowHeight(r) === 0) setSize(activeRef.current, "row", r, DEFAULT_ROW_PX, true); onUserEdit?.(); }}>Unhide rows</MenuItem>}
              {menu.header === "col" && <MenuItem onClick={() => { close(); for (let c = c0; c <= c1; c++) setSize(activeRef.current, "col", c, 0, true); onUserEdit?.(); }}>Hide {plural(nC, "column")}</MenuItem>}
              {menu.header === "col" && <MenuItem onClick={() => { close(); const g = geom(activeRef.current); for (let c = Math.max(0, c0 - 1); c <= Math.min((g?.colCount ?? 0) - 1, c1 + 1); c++) if (g && g.colWidth(c) === 0) setSize(activeRef.current, "col", c, DEFAULT_COL_PX, true); onUserEdit?.(); }}>Unhide columns</MenuItem>}
            </Menu>
          );
        })()}
        <TabBar>
          <TabArrows>
            <button disabled={tabScroll.left <= 0} onClick={() => { const el = tabsRef.current; if (el) { el.scrollLeft = Math.max(0, el.scrollLeft - el.clientWidth * 0.8); measureTabs(); } }} title="Scroll tabs left">◀</button>
            <button disabled={tabScroll.left >= tabScroll.max} onClick={() => { const el = tabsRef.current; if (el) { el.scrollLeft = Math.min(tabScroll.max, el.scrollLeft + el.clientWidth * 0.8); measureTabs(); } }} title="Scroll tabs right">▶</button>
          </TabArrows>
        <Tabs ref={tabsRef} onScroll={measureTabs}>
          {sheetNames.map((s, i) =>
            s.hidden && !showHidden ? null : renaming && renaming.index === i ? (
              <TabInput
                key={`rename-${i}`}
                autoFocus
                value={renaming.text}
                onChange={(e) => setRenaming({ index: i, text: e.target.value })}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    const next = renaming.text.trim();
                    if (next && next !== s.name && !sheetNames.some((x) => x.name === next)) { handle.renameSheet(s.name, next); onManualSheetOp?.({ kind: "rename", oldName: s.name, newName: next }); }
                    setRenaming(null);
                  } else if (e.key === "Escape") setRenaming(null);
                  e.stopPropagation();
                }}
                onBlur={() => setRenaming(null)}
              />
            ) : (
              <Tab
                key={`${s.name}-${i}`}
                $active={i === active}
                $hidden={s.hidden}
                $color={s.color}
                onClick={() => activateSheet(i)}
                onDoubleClick={() => setRenaming({ index: i, text: s.name })}
                onContextMenu={(e) => { e.preventDefault(); const root = (e.currentTarget as HTMLElement).closest("[data-vp-root]")?.getBoundingClientRect(); setTabMenu({ x: e.clientX - (root?.left ?? 0), y: e.clientY - (root?.top ?? 0) - 110, index: i }); }}
                title={s.hidden ? "hidden sheet" : undefined}
              >
                {s.name}
              </Tab>
            ),
          )}
        </Tabs>
        </TabBar>
      </Root>
    );

    function startEditFromBar(text: string) {
      const s = selRef.current;
      setEditing({ row: s.ar, col: s.ac, text, origin: "bar" });
    }
  }),
);

/** An HTML table of the copied block with inline styles, for other apps. */
function htmlTableForCopy(
  r0: number, c0: number, r1: number, c1: number,
  at: (r: number, c: number) => { text: string; format: CellFormatShape | null; align: "left" | "center" | "right" },
): string {
  const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const rows: string[] = [];
  for (let r = r0; r <= r1; r++) {
    const tds: string[] = [];
    for (let c = c0; c <= c1; c++) {
      const { text, format: f, align } = at(r, c);
      const st: string[] = [];
      if (f?.bold) st.push("font-weight:700");
      if (f?.italic) st.push("font-style:italic");
      if (f?.underline) st.push("text-decoration:underline");
      if (f?.font_color) st.push(`color:${f.font_color}`);
      if (f?.background_color) st.push(`background:${f.background_color}`);
      if (f?.font_family) st.push(`font-family:${f.font_family}`);
      if (f?.font_size) st.push(`font-size:${f.font_size}pt`);
      if (f?.number_format && f.number_format !== "General") st.push(`mso-number-format:'${f.number_format.replace(/'/g, "\\'")}'`);
      st.push(`text-align:${align}`);
      tds.push(`<td style="${st.join(";")}">${esc(text)}</td>`);
    }
    rows.push(`<tr>${tds.join("")}</tr>`);
  }
  return `<table>${rows.join("")}</table>`;
}

/** Formats from an HTML table on the clipboard (Excel, Sheets, Numbers): one
 * entry per cell with any of bold/italic/colour/fill/alignment/number format. */
function formatsFromHtmlTable(html: string): Array<{ dr: number; dc: number; format: CellFormatShape }> {
  const out: Array<{ dr: number; dc: number; format: CellFormatShape }> = [];
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(html, "text/html");
  } catch {
    return out;
  }
  const rows = doc.querySelectorAll("table tr");
  const toHex = (css: string | null): string | null => {
    if (!css) return null;
    const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(css);
    if (m) return `#${[m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, "0")).join("").toUpperCase()}`;
    if (/^#[0-9a-f]{6}$/i.test(css)) return css.toUpperCase();
    return null;
  };
  rows.forEach((tr, dr) => {
    let dc = 0;
    tr.querySelectorAll("td, th").forEach((td) => {
      const el = td as HTMLElement;
      const st = el.getAttribute("style") ?? "";
      const get = (prop: string) => { const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*([^;]+)`, "i").exec(st); return m ? m[1].trim() : null; };
      const f: CellFormatShape = {};
      const fw = get("font-weight");
      if (fw === "bold" || fw === "700" || el.querySelector("b, strong")) f.bold = true;
      if (get("font-style") === "italic" || el.querySelector("i, em")) f.italic = true;
      const color = toHex(get("color"));
      if (color && color !== "#000000") f.font_color = color;
      const bg = toHex(get("background") ?? get("background-color"));
      if (bg && bg !== "#FFFFFF") f.background_color = bg;
      const ta = get("text-align");
      if (ta === "left" || ta === "center" || ta === "right") f.horizontal_align = ta;
      const nf = get("mso-number-format");
      if (nf) { const clean = nf.replace(/^["']|["']$/g, "").replace(/\\/g, ""); if (clean && clean.toLowerCase() !== "general") f.number_format = clean; }
      if (Object.keys(f).length) out.push({ dr, dc, format: f });
      dc += Math.max(1, Number(el.getAttribute("colspan") ?? 1));
    });
  });
  return out;
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

/** Does an A1 merge range intersect the 0-based rectangle? */
function rangesOverlap(a1: string, sr: number, sc: number, er: number, ec: number): boolean {
  const [a, b] = a1.split(":");
  const p = parseA1(a);
  const q = b ? parseA1(b) : p;
  if (!p || !q) return false;
  const r0 = Math.min(p.row0, q.row0), r1 = Math.max(p.row0, q.row0), c0 = Math.min(p.col0, q.col0), c1 = Math.max(p.col0, q.col0);
  return !(r1 < sr || r0 > er || c1 < sc || c0 > ec);
}

/** Greedy word wrap for canvas text; long words break by character. */
function wrapText(ctx: CanvasRenderingContext2D, text: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split(/\r?\n/)) {
    let line = "";
    for (const word of para.split(" ")) {
      const probe = line ? `${line} ${word}` : word;
      if (ctx.measureText(probe).width <= maxW || !line) {
        if (ctx.measureText(word).width > maxW && !line) {
          // Break a single long word.
          let chunk = "";
          for (const ch of word) {
            if (ctx.measureText(chunk + ch).width > maxW && chunk) { out.push(chunk); chunk = ""; }
            chunk += ch;
          }
          line = chunk;
        } else line = probe;
      } else {
        out.push(line);
        line = word;
      }
    }
    out.push(line);
    if (out.length > 200) break;
  }
  return out;
}

/** The text a cell shows in the formula bar when it has no formula. */
function rawInput(cell: GridCell): string {
  const v = cell.v;
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return String(v);
}

function coerceForPatch(input: string): unknown {
  if (input === "TRUE") return true;
  if (input === "FALSE") return false;
  if (/^-?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(input)) return Number(input);
  return input;
}
