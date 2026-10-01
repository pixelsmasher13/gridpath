/**
 * gridpath review page: the workbook (Univer core preset, read-only, cached
 * values from the Node engine) with pending changes highlighted, a panel of
 * change batches with before/after, and Accept / Reject / Save.
 *
 * Server is the source of truth: every action refetches and reloads.
 */
import { createUniver, LocaleType, merge } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import coreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import "@univerjs/preset-sheets-core/lib/index.css";
import { groupBatchForReview } from "../src/core/reviewGroups";
import type { UniverMutation } from "../src/core/types";

type DiffCell = { sheet: string; row: number; col: number; before: unknown; after: unknown; before_formula: string | null; after_formula: string | null };
type Batch = { id: string; tool: string; justification: string; status: "pending" | "accepted"; created_at: string; cells: DiffCell[]; formats: number; structural: string[] };
type Workbook = { path: string; sheetNames: string[]; snapshot: any; display: Record<string, string>; batches: Batch[] };

const params = new URLSearchParams(location.search);
const PATH = params.get("path") ?? "";
const TOKEN = params.get("token") ?? "";
const PENDING_BG = "#E6FFEC";
const ACCEPTED_BG = "#E8F0FE";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const status = (msg: string) => ($("status").textContent = msg);

async function api(path: string, body?: unknown): Promise<any> {
  const r = await fetch(`/api/${path}${path.includes("?") ? "&" : "?"}token=${TOKEN}`, body ? { method: "POST", body: JSON.stringify(body) } : undefined);
  return r.json();
}

// ---- grid ----

let univerAPI: any = null;
let activeSheet = "";
/** Last cell the user jumped to; re-applied after every reload so Accept/Reject don't lose the scroll position. */
let lastJump: { sheet: string; row: number; col: number } | null = null;

function ensureUniver() {
  if (univerAPI) return univerAPI;
  const worker = new Worker(new URL("./calc.worker.ts", import.meta.url), {
    type: "module",
    name: "univer-calc-noformula",
  });
  const u = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: merge({}, coreEnUS) },
    presets: [UniverSheetsCorePreset({ container: "grid", toolbar: false, footer: false, workerURL: worker })],
  });
  univerAPI = u.univerAPI;
  (window as any).__gp = { univerAPI };
  return univerAPI;
}

/** Paint changed cells into the snapshot before mounting (no per-cell commands needed). */
function highlight(snapshot: any, batches: Batch[]): any {
  const byId: Record<string, string> = {};
  for (const id of snapshot.sheetOrder) byId[snapshot.sheets[id].name] = id;
  for (const b of batches) {
    const bg = b.status === "accepted" ? ACCEPTED_BG : PENDING_BG;
    for (const c of b.cells) {
      const sid = byId[c.sheet];
      if (!sid) continue;
      const cd = snapshot.sheets[sid].cellData;
      const cell = ((cd[c.row] ??= {})[c.col] ??= {});
      cell.s = { ...(cell.s ?? {}), bg: { rgb: bg } };
    }
  }
  return snapshot;
}

function loadGrid(wb: Workbook) {
  const api = ensureUniver();
  const snap = highlight(structuredClone(wb.snapshot), wb.batches);
  snap.id = `review_${Date.now()}`;
  const prev = api.getActiveWorkbook?.();
  if (prev) api.disposeUnit?.(prev.getId());
  const fwb = api.createWorkbook(snap);
  try {
    fwb.setEditable?.(false);
  } catch {
    /* older facade */
  }
  if (activeSheet) {
    const ws = fwb.getSheetByName?.(activeSheet);
    if (ws) fwb.setActiveSheet?.(ws);
  }
  if (lastJump && lastJump.sheet === activeSheet) {
    const ws = fwb.getSheetByName?.(lastJump.sheet);
    if (ws) void scrollIntoView(ws, Math.max(0, lastJump.row - 2), lastJump.col);
  }
}

function jumpTo(sheet: string, row: number, col: number) {
  const fwb = univerAPI?.getActiveWorkbook?.();
  const ws = fwb?.getSheetByName?.(sheet);
  if (!ws) return;
  fwb.setActiveSheet?.(ws);
  activeSheet = sheet;
  lastJump = { sheet, row, col };
  const range = ws.getRange?.(row, col);
  range?.activate?.();
  void scrollIntoView(ws, Math.max(0, row - 2), col);
  renderTabs(currentWb!);
}

/**
 * Univer's scroll-to-cell lands the viewport a few columns past the target
 * on sheets with frozen panes (measured: a constant +4 on a Canalyst model).
 * Scroll, wait for the render, read back what's visible, and correct by the
 * observed error. Reading before the render settles gives stale ranges, so
 * every step waits a frame.
 */
async function scrollIntoView(ws: any, row: number, col: number) {
  const settle = () => new Promise((r) => setTimeout(r, 90));
  let targetRow = row;
  let targetCol = col;
  for (let i = 0; i < 4; i++) {
    try {
      ws.scrollToCell?.(Math.max(0, targetRow), Math.max(0, targetCol));
    } catch {
      return;
    }
    await settle();
    const vr = ws.getVisibleRange?.();
    if (!vr) return;
    const colOk = col >= vr.startColumn && col <= vr.endColumn - 1;
    const rowOk = row >= vr.startRow && row <= vr.endRow - 1;
    if (colOk && rowOk) return;
    if (!colOk) targetCol = Math.max(0, targetCol - (vr.startColumn - col));
    if (!rowOk) targetRow = Math.max(0, targetRow - (vr.startRow - row));
  }
}

// ---- panel ----

let currentWb: Workbook | null = null;

const fmt = (v: unknown, f: string | null) => (f ? f : v === null || v === undefined || v === "" ? "∅" : String(v));
const a1 = (row: number, col: number) => {
  let s = "";
  let n = col + 1;
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${row + 1}`;
};

function renderTabs(wb: Workbook) {
  const counts: Record<string, number> = {};
  for (const b of wb.batches) for (const c of b.cells) counts[c.sheet] = (counts[c.sheet] ?? 0) + 1;
  const el = $("tabs");
  el.innerHTML = "";
  for (const name of wb.sheetNames) {
    const btn = document.createElement("button");
    btn.className = name === activeSheet ? "active" : "";
    btn.textContent = name;
    if (counts[name]) {
      const n = document.createElement("span");
      n.className = "n";
      n.textContent = `●${counts[name]}`;
      btn.appendChild(n);
    }
    btn.onclick = () => {
      activeSheet = name;
      const fwb = univerAPI?.getActiveWorkbook?.();
      const ws = fwb?.getSheetByName?.(name);
      if (ws) fwb.setActiveSheet?.(ws);
      renderTabs(wb);
    };
    el.appendChild(btn);
  }
}

function renderPanel(wb: Workbook) {
  const root = $("batches");
  root.innerHTML = "";
  $("empty").hidden = wb.batches.length > 0;
  const pending = wb.batches.filter((b) => b.status === "pending").length;
  ($("accept-all") as HTMLButtonElement).disabled = pending === 0;
  ($("save") as HTMLButtonElement).disabled = wb.batches.length === 0;
  ($("save-as") as HTMLButtonElement).disabled = wb.batches.length === 0;

  for (const b of wb.batches) {
    const card = document.createElement("div");
    card.className = "card";
    const sheets = [...new Set(b.cells.map((c) => c.sheet))];
    const summary = [
      b.cells.length ? `${b.cells.length} cell${b.cells.length === 1 ? "" : "s"}${sheets.length ? ` on ${sheets.join(", ")}` : ""}` : "",
      b.formats ? `${b.formats} format${b.formats === 1 ? "" : "s"}` : "",
      b.structural.length ? `<span class="structural">${b.structural.join(", ")}</span>` : "",
    ]
      .filter(Boolean)
      .join(" · ");
    const shown = b.cells.slice(0, 12);
  const sheetById = Object.fromEntries(wb.snapshot.sheetOrder.map((id: string) => [wb.snapshot.sheets[id].name, wb.snapshot.sheets[id]]));
  const hiddenCol = (c: DiffCell) => sheetById[c.sheet]?.columnData?.[c.col]?.hd === 1;
    card.innerHTML = `
      <div class="head"><span class="tool">${b.tool}</span><span class="badge ${b.status}">${b.status}</span><span class="spacer"></span><span class="muted">${b.id}</span></div>
      <div class="summary">${summary || "no cell changes"}</div>
      <div class="diff">${shown
        .map(
          (c) =>
            `<span class="addr" data-sheet="${c.sheet}" data-row="${c.row}" data-col="${c.col}" title="${hiddenCol(c) ? "This column is hidden in the sheet" : "Jump to cell"}">${sheets.length > 1 ? c.sheet + "!" : ""}${a1(c.row, c.col)}${hiddenCol(c) ? " <span class=\"muted\">(hidden col)</span>" : ""}</span><span class="row"><span class="old">${escapeHtml(fmt(c.before, c.before_formula))}</span> → <span class="new">${escapeHtml(fmt(c.after, c.after_formula))}</span>${b.status === "pending" ? `<button class="x" title="Reject this cell only" data-x-sheet="${c.sheet}" data-x-row="${c.row}" data-x-col="${c.col}">✕</button>` : ""}</span>`,
        )
        .join("")}${b.cells.length > shown.length ? `<span class="more">… ${b.cells.length - shown.length} more</span>` : ""}</div>
      <div class="actions">
        ${b.status === "pending" ? `<button class="btn" data-act="accept">Accept</button>` : ""}
        <button class="btn danger" data-act="reject">Reject</button>
      </div>`;
    card.querySelectorAll<HTMLElement>(".addr").forEach((el) => (el.onclick = () => jumpTo(el.dataset.sheet!, Number(el.dataset.row), Number(el.dataset.col))));
    card.querySelectorAll<HTMLButtonElement>("button.x").forEach((el) =>
      el.addEventListener("click", () => act("batch/reject-cells", { path: PATH, batch_id: b.id, cells: [{ sheet: el.dataset.xSheet, row: Number(el.dataset.xRow), col: Number(el.dataset.xCol) }] })),
    );
    card.querySelector<HTMLButtonElement>('[data-act="accept"]')?.addEventListener("click", () => act("batch/accept", { path: PATH, batch_id: b.id }));
    card.querySelector<HTMLButtonElement>('[data-act="reject"]')?.addEventListener("click", () => act("batch/reject", { path: PATH, batch_id: b.id }));
    root.appendChild(card);
  }
}

function escapeHtml(s: string) {
  return s.replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[ch]!);
}

// ---- flow ----

async function refresh() {
  status("loading…");
  if (!univerAPI) $("grid").innerHTML = '<div class="loading">Loading workbook…</div>';
  const wb = (await api(`workbook?path=${encodeURIComponent(PATH)}`)) as Workbook & { error?: string };
  if (wb.error) return status(`error: ${wb.error}`);
  currentWb = wb;
  if (!activeSheet || !wb.sheetNames.includes(activeSheet)) {
    activeSheet = wb.batches.find((b) => b.cells.length)?.cells[0].sheet ?? wb.sheetNames[0];
  }
  $("file").textContent = wb.path.split("/").slice(-2).join("/");
  $("file").title = wb.path;
  loadGrid(wb);
  renderTabs(wb);
  renderPanel(wb);
  // First load: land on the first changed cell so the edit is in view even
  // on a 1,400-row sheet. Later refreshes keep the user's own position.
  if (!lastJump) {
    const first = wb.batches.find((b) => b.cells.length)?.cells[0];
    if (first) setTimeout(() => jumpTo(first.sheet, first.row, first.col), 400);
  }
  const pending = wb.batches.filter((b) => b.status === "pending").length;
  status(wb.batches.length ? `${wb.batches.length} batch(es), ${pending} pending — nothing is written to the file until you save.` : "No pending changes. The file on disk is untouched.");
}

async function act(path: string, body: unknown) {
  status("…");
  const r = await api(path, body);
  if (r?.ok === false) status(`error: ${r.error ?? JSON.stringify(r)}`);
  await refresh();
  if (path === "save" && r?.saved) status(r.copy ? r.note : `Saved ${r.bytes} bytes to ${r.path}. Only the parts you changed were rewritten.`);
}

$("accept-all").onclick = () => act("batch/accept", { path: PATH });
$("save").onclick = () => {
  if (confirm("Write all accepted and pending changes into the file on disk?")) void act("save", { path: PATH });
};
$("save-as").onclick = () => {
  const suggested = PATH.replace(/(\.xlsm|\.xlsx)$/i, " (edited)$1");
  const target = prompt("Save a copy with all accepted and pending changes to:", suggested);
  if (target && target.trim()) void act("save", { path: PATH, as: target.trim() });
};

// keep the panel live while an agent is still working
refresh().catch((e) => status(`error: ${e}`));
setInterval(() => {
  if (document.visibilityState === "visible") {
    api(`workbook?path=${encodeURIComponent(PATH)}`).then((wb: Workbook) => {
      const sig = JSON.stringify(wb.batches.map((b) => [b.id, b.status, b.cells.length]));
      const cur = JSON.stringify((currentWb?.batches ?? []).map((b) => [b.id, b.status, b.cells.length]));
      if (sig !== cur) void refresh();
    });
  }
}, 3000);

// (imported so the pure grouping helper is available for a richer panel later)
void groupBatchForReview;
export type { UniverMutation };
