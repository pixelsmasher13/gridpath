/**
 * The site demo's file report: what the open workbook carries and how long
 * it took to calculate, and, on request, what a save changed, part by part
 * against the file that was opened. Plus the "Open your .xlsx" button and
 * the drop zone.
 */
import React, { useEffect, useRef, useState } from "react";
import type { Spreadsheet as SheetApi } from "gridpath";
import { compareSaves, countFormulasAndNames, fileContents, partNote, type FileContents, type SaveComparison } from "./fileReport";

export const BLUE = "#1F4E79";

/** A workbook opened in the demo, with the bytes it was opened from. */
export interface Opened {
  name: string;
  bytes: Uint8Array;
  /** Until the grid drew, and until every formula was calculated (null while calculating). */
  openMs: number;
  calcMs: number | null;
  contents: FileContents;
  formulas: number | null;
  names: number | null;
  /** Formulas showing Excel's saved value. */
  issues: number | null;
  /** Charts the grid draws (null: the package can't say, as before 0.5). */
  drawnCharts: number | null;
}

/** An .xlsx/.xlsm (a zip), as opposed to CSV text opened as a new workbook. */
export const isWorkbookFile = (b: Uint8Array) => b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b;

/** Fills in what takes a moment: the calculation, the counts, the issues. */
export async function analyse(sheet: SheetApi, name: string, bytes: Uint8Array, t0: number, update: (o: Opened) => void): Promise<void> {
  const base: Opened = { name, bytes, openMs: performance.now() - t0, calcMs: null, contents: fileContents(bytes), formulas: null, names: null, issues: null, drawnCharts: null };
  update(base);
  const counted = countFormulasAndNames(bytes).catch(() => null);
  let calcMs: number | null = null;
  try {
    await sheet.whenCalculated(600_000);
    calcMs = performance.now() - t0;
  } catch {
    // Still calculating after ten minutes: leave the time out.
  }
  const counts = await counted;
  const issues = await sheet.getCalcIssues(1000).then((list) => list.length, () => null);
  const charts = (sheet as { getCharts?: () => Promise<unknown[]> }).getCharts;
  const drawnCharts = charts ? await charts.call(sheet).then((list) => list.length, () => null) : null;
  update({ ...base, calcMs, formulas: counts?.formulas ?? null, names: counts?.names ?? null, issues, drawnCharts });
}

const secs = (ms: number) => (ms < 1000 ? `${Math.max(0.1, ms / 1000).toFixed(1)} s` : ms < 10_000 ? `${(ms / 1000).toFixed(1)} s` : `${Math.round(ms / 1000)} s`);
const n = (v: number) => v.toLocaleString("en-US");
const plural = (v: number, word: string, many = `${word}s`) => `${n(v)} ${v === 1 ? word : many}`;

/** The parts worth naming: what a typical library loses on save. */
function carries(c: FileContents, names: number | null, drawn: number | null): string[] {
  const out = [plural(c.sheets, "sheet")];
  // Every chart is kept in the file on save; the grid draws line charts.
  if (c.charts) {
    const note = drawn === null ? null : drawn >= c.charts ? "" : drawn > 0 ? ` (${n(drawn)} drawn, all kept)` : " (kept, not drawn)";
    out.push(`${plural(c.charts, "chart")}${note ?? " (kept, not drawn)"}`);
  }
  if (c.pivotTables) out.push(plural(c.pivotTables, "pivot table"));
  if (c.tables) out.push(plural(c.tables, "table"));
  if (c.macros) out.push("VBA macros");
  if (names) out.push(plural(names, "defined name"));
  if (c.comments) out.push("comments");
  if (c.images) out.push(plural(c.images, "image"));
  if (c.customXml) out.push(plural(c.customXml, "custom XML part"));
  if (c.externalLinks) out.push(plural(c.externalLinks, "link to another workbook", "links to other workbooks"));
  return out;
}

type Check = { state: "running" } | { state: "done"; result: SaveComparison; saved: Uint8Array } | { state: "error"; message: string };

export function ReportBar({ sheet, opened, compact, stale }: { sheet: SheetApi; opened: Opened; compact?: boolean; stale: number }) {
  const [check, setCheck] = useState<Check | null>(null);
  // An edit makes the last check out of date.
  useEffect(() => setCheck(null), [stale, opened]);

  const run = async () => {
    setCheck({ state: "running" });
    try {
      const saved = await sheet.save();
      setCheck({ state: "done", result: compareSaves(opened.bytes, saved), saved });
    } catch (e: any) {
      setCheck({
        state: "error",
        message: e?.code === "pending-changes" ? "Accept or reject the AI's changes first (the AI changes panel), then check again." : `${e?.message ?? e}. ${e?.hint ?? ""}`,
      });
    }
  };
  const download = (bytes: Uint8Array) => {
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = Object.assign(document.createElement("a"), { href: url, download: opened.name });
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const o = opened;
  const formulas = o.formulas === null ? "formulas" : plural(o.formulas, "formula");
  return (
    <div data-report="" style={{ borderBottom: "1px solid #e4e4e7", background: "#f7f8fa", padding: compact ? "6px 12px" : "8px 14px", fontSize: 13, color: "#333", lineHeight: 1.5 }}>
      <div style={{ display: "flex", flexWrap: "wrap", alignItems: "baseline", columnGap: 10, rowGap: 2 }}>
        <span>
          <strong>{o.name}</strong> opened in {secs(o.openMs)}
          {o.calcMs === null ? <span style={{ color: "#777" }}> · calculating {formulas}…</span> : <> · {formulas} calculated in {secs(o.calcMs)}, in this browser</>}
        </span>
        {!compact && <span style={{ color: "#666" }}>{carries(o.contents, o.names, o.drawnCharts).join(" · ")}</span>}
        {!!o.issues && (
          <span style={{ color: "#8a5a00" }} title="GridPath couldn't calculate these formulas, so they show the value Excel saved; the ⚠ button in the toolbar lists them">
            {plural(o.issues, "formula shows", "formulas show")} Excel's saved value
          </span>
        )}
        <span style={{ flex: 1 }} />
        <button
          onClick={() => void run()}
          disabled={check?.state === "running" || o.calcMs === null}
          title="Saves a copy (not downloaded) and compares it with the file you opened, part by part"
          style={{ padding: "3px 10px", borderRadius: 6, border: `1px solid ${BLUE}`, background: "#fff", color: BLUE, cursor: "pointer", fontSize: 13 }}
        >
          {check?.state === "running" ? "Saving a copy…" : "Check what a save keeps"}
        </button>
      </div>
      {check?.state === "error" && <div style={{ color: "#b00020", marginTop: 4 }}>{check.message}</div>}
      {check?.state === "done" && <SaveResult result={check.result} onDownload={() => download(check.saved)} />}
    </div>
  );
}

function SaveResult({ result: r, onDownload }: { result: SaveComparison; onDownload: () => void }) {
  const changed = [...r.rewritten.map((p) => ({ p, how: "rewritten" })), ...r.added.map((p) => ({ p, how: "added" })), ...r.removed.map((p) => ({ p, how: "removed" }))];
  return (
    <div style={{ marginTop: 6, padding: "6px 10px", background: "#fff", border: "1px solid #e4e4e7", borderRadius: 6 }}>
      <div>
        <strong style={{ color: "#166534" }}>
          {n(r.unchanged)} of {n(r.total)} parts came back byte-identical.
        </strong>{" "}
        {changed.length === 0 ? (
          <span style={{ color: "#555" }}>Nothing has changed yet: edit a cell or a format and check again to see what a change rewrites.</span>
        ) : (
          <span style={{ color: "#555" }}>
            Only what the session changed was rewritten:{" "}
            {changed.slice(0, 4).map(({ p, how }, i) => (
              <span key={p}>
                {i > 0 ? ", " : ""}
                <code style={{ fontSize: 12 }}>{p}</code>
                {how !== "rewritten" ? ` (${how})` : partNote(p) ? ` (${partNote(p)})` : ""}
              </span>
            ))}
            {changed.length > 4 ? `, and ${changed.length - 4} more` : ""}.
          </span>
        )}{" "}
        <button onClick={onDownload} style={{ border: 0, background: "none", color: BLUE, textDecoration: "underline", cursor: "pointer", fontSize: 13, padding: 0 }}>
          Download this copy
        </button>
      </div>
    </div>
  );
}

/** The header's file button. */
export function OpenButton({ onFile, disabled, small }: { onFile: (f: File) => void; disabled?: boolean; small?: boolean }) {
  const input = useRef<HTMLInputElement | null>(null);
  return (
    <>
      <button
        disabled={disabled}
        onClick={() => input.current?.click()}
        title="Opens in this browser; nothing is uploaded"
        style={{ padding: small ? "4px 10px" : "6px 14px", borderRadius: 6, border: `1px solid ${BLUE}`, background: BLUE, color: "#fff", cursor: "pointer", fontWeight: 600, fontSize: small ? 13 : 14 }}
      >
        Open your .xlsx
      </button>
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xlsm,.csv,.tsv,.txt"
        hidden
        data-open-file=""
        onChange={(e) => {
          const f = e.target.files?.[0];
          e.target.value = "";
          if (f) onFile(f);
        }}
      />
    </>
  );
}

/** A file dragged over the page shows where to drop it. */
export function useFileDrop(onFile: (f: File) => void): boolean {
  const [over, setOver] = useState(false);
  const depth = useRef(0);
  const latest = useRef(onFile);
  latest.current = onFile;
  useEffect(() => {
    const hasFiles = (e: DragEvent) => !!e.dataTransfer && [...e.dataTransfer.types].includes("Files");
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth.current++;
      setOver(true);
    };
    const leave = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      depth.current = Math.max(0, depth.current - 1);
      if (depth.current === 0) setOver(false);
    };
    const overFn = (e: DragEvent) => hasFiles(e) && e.preventDefault();
    const drop = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth.current = 0;
      setOver(false);
      const f = e.dataTransfer?.files[0];
      if (f) latest.current(f);
    };
    window.addEventListener("dragenter", enter);
    window.addEventListener("dragleave", leave);
    window.addEventListener("dragover", overFn);
    window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragenter", enter);
      window.removeEventListener("dragleave", leave);
      window.removeEventListener("dragover", overFn);
      window.removeEventListener("drop", drop);
    };
  }, []);
  return over;
}

export function DropOverlay() {
  return (
    <div
      style={{ position: "fixed", inset: 0, zIndex: 100, background: "rgba(31,78,121,0.12)", display: "flex", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}
    >
      <div style={{ padding: "24px 32px", border: `2px dashed ${BLUE}`, borderRadius: 12, background: "#fff", fontSize: 16, color: BLUE, textAlign: "center" }}>
        <strong>Drop an .xlsx to open it</strong>
        <div style={{ fontSize: 13, color: "#555", marginTop: 4 }}>It opens, calculates and saves in this browser. Nothing is uploaded.</div>
      </div>
    </div>
  );
}
