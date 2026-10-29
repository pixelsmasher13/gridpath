import React, { useCallback, useEffect, useRef, useState } from "react";
import { Spreadsheet, type Spreadsheet as SheetApi } from "gridpath";
import { compareSaves, countFormulasAndNames, fileContents, type SaveComparison } from "./fileReport";
import "./recording.css";

const BASE = (import.meta as any).env.BASE_URL as string;
const FILE = "larkspur-income-statement.xlsx";
const SHEET = "Income statement";
type Expected = { checks: Array<{sheet: string; ref: string; value: number}>; charts: Array<{title: string; series: number[][]}> };
type Fixture = { bytes: Uint8Array; formulas: number; years: number; periods: number; sheets: number; chartPoints: number; cells: number; sha256: string; base: Expected; edited: Expected };
type Timings = { openMs?: number; calculateMs?: number; editMs?: number; saveMs?: number };
type Phase = "opening" | "calculating" | "editing" | "saving";
const seconds = (ms?: number) => ms === undefined ? "—" : `${(ms / 1000).toFixed(2)} s`;
const paint = () => new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
function download(bytes: BlobPart, name: string, type: string) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function RecordingDemo() {
  const [fixture, setFixture] = useState<Fixture | null>(null);
  const [source, setSource] = useState<Uint8Array | null>(null);
  const [api, setApi] = useState<SheetApi | null>(null);
  const [phase, setPhase] = useState<Phase | null>(null);
  const [elapsed, setElapsed] = useState(0);
  const [times, setTimes] = useState<Timings>({});
  const [editBreakdown, setEditBreakdown] = useState<ReturnType<SheetApi["getLastEditTimings"]>>(null);
  const [progress, setProgress] = useState("Preparing the synthetic workbook…");
  const [error, setError] = useState<string | null>(null);
  const [edited, setEdited] = useState(false);
  const [verified, setVerified] = useState(false);
  const [issues, setIssues] = useState<number | null>(null);
  const [comparison, setComparison] = useState<SaveComparison | null>(null);
  const [saved, setSaved] = useState<Uint8Array | null>(null);
  const [run, setRun] = useState(0);
  const [showTimings, setShowTimings] = useState(false);
  const [chartView, setChartView] = useState(false);
  const started = useRef(0);
  const openedAt = useRef("");
  const openTrigger = useRef<"automatic" | "reset">("automatic");
  const generation = useRef(0);

  useEffect(() => {
    const previous = document.title;
    document.title = "Larkspur Orbital · Mission Control";
    return () => { document.title = previous; };
  }, []);

  const openFixture = useCallback((next: Fixture, trigger: "automatic" | "reset") => {
    generation.current++;
    openTrigger.current = trigger;
    openedAt.current = new Date().toISOString();
    setFixture(next); setEditBreakdown(null); setChartView(false);
    setTimes({}); setEdited(false); setVerified(false); setIssues(null); setComparison(null); setSaved(null); setError(null); setApi(null);
    setRun(n => n + 1);
    setProgress("Opening workbook…");
    started.current = performance.now();
    setElapsed(0); setPhase("opening"); setSource(next.bytes);
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      const [file, metadata] = await Promise.all([fetch(`${BASE}${FILE}`), fetch(`${BASE}larkspur-income-statement.json`)]);
      if (!file.ok || !metadata.ok) throw new Error("Synthetic workbook is missing. See RECORDING.md for setup.");
      const bytes = new Uint8Array(await file.arrayBuffer());
      const meta = await metadata.json();
      const [counts, digest] = await Promise.all([countFormulasAndNames(bytes), crypto.subtle.digest("SHA-256", bytes)]);
      const sha256 = Array.from(new Uint8Array(digest), n => n.toString(16).padStart(2, "0")).join("");
      if (!counts || sha256 !== meta.sha256 || counts.formulas !== meta.formulas) throw new Error("Workbook does not match its verified metadata.");
      if (!cancelled) {
        openFixture({ ...meta, bytes, sha256, formulas: counts.formulas }, "automatic");
      }
    })().catch(e => !cancelled && setError(String(e.message ?? e)));
    return () => { cancelled = true; };
  }, [openFixture]);

  useEffect(() => {
    if (!phase) return;
    const id = setInterval(() => setElapsed(performance.now() - started.current), 50);
    return () => clearInterval(id);
  }, [phase]);

  const fail = (e: unknown) => { setError(e instanceof Error ? e.message : String(e)); setPhase(null); };
  const verify = async (sheet: SheetApi, expected: Expected) => {
    const cells = await Promise.all(expected.checks.map(c => sheet.getCell(c.sheet, c.ref)));
    if (!cells.every((cell, i) => typeof cell?.value === "number" && !cell.fromFile && Math.abs(cell.value - expected.checks[i].value) < 0.000001)) {
      throw new Error("Calculated summary differs from the independent expected totals.");
    }
    const charts = (await sheet.getCharts()).filter(c => c.sheet === SHEET);
    if (charts.length !== expected.charts.length || !expected.charts.every(want => {
      const chart = charts.find(c => c.title === want.title);
      return chart && !chart.unsupported && chart.series.length === want.series.length && chart.series.every((series, i) =>
        series.values.length === want.series[i].length && series.values.every((value, j) =>
          typeof value === "number" && Math.abs(value - want.series[i][j]) < 0.000001));
    })) {
      throw new Error("The live charts do not match the independently calculated income statement.");
    }
  };
  const ready = async (sheet: SheetApi) => {
    const thisRun = generation.current;
    const openMs = performance.now() - started.current;
    setApi(sheet); setTimes({ openMs }); setPhase("calculating"); setProgress("Calculating formulas…");
    try {
      await sheet.whenCalculated(600_000);
      const calculateMs = performance.now() - started.current;
      if (thisRun !== generation.current) return;
      setTimes({ openMs, calculateMs }); setElapsed(calculateMs);
      await verify(sheet, fixture!.base);
      const list = await sheet.getCalcIssues(1_000_000);
      if (thisRun !== generation.current) return;
      setIssues(list.length); setVerified(true); setElapsed(calculateMs); setPhase(null); setProgress("Income statement verified. Edit the launch assumption or explore the operating and capital schedules.");
    } catch (e) { if (thisRun === generation.current) fail(e); }
  };
  const edit = async () => {
    if (!api || !fixture) return;
    setError(null); setComparison(null); setSaved(null); setVerified(false); setPhase("editing"); setProgress("Changing launches from 120 to 180 and checking the forecast…");
    started.current = performance.now(); setElapsed(0);
    try {
      api.setCell(SHEET, "F5", 180);
      await api.whenCalculated(600_000);
      await verify(api, fixture.edited);
      await paint();
      const editMs = performance.now() - started.current;
      setEditBreakdown(api.getLastEditTimings());
      setTimes(t => ({...t, editMs})); setElapsed(editMs); setEdited(true); setVerified(true); setPhase(null);
      setProgress("120 → 180 paid standard launches. Revenue, EBITDA, net income and all charts verified.");
    } catch (e) { fail(e); }
  };
  const save = async (downloadFile = false) => {
    if (!api || !fixture) return;
    setError(null); setPhase("saving"); setProgress("Saving the edited workbook…"); started.current = performance.now(); setElapsed(0);
    try {
      const bytes = await api.save();
      const saveMs = performance.now() - started.current;
      setTimes(t => ({...t, saveMs})); setElapsed(saveMs); setSaved(bytes);
      setComparison(compareSaves(fixture.bytes, bytes)); setPhase(null); setProgress("Saved. The comparison below is against the original workbook.");
      if (downloadFile) download(bytes as BlobPart, "larkspur-income-statement-edited.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } catch (e) { fail(e); }
  };
  const undo = async () => {
    if (!api || !fixture) return;
    setPhase("editing"); started.current = performance.now(); setVerified(false); setError(null);
    setProgress("Undoing the assumption and verifying the restored chart…");
    try {
      api.undo(); await api.whenCalculated(600_000); await verify(api, fixture.base);
      setEdited(false); setVerified(true); setSaved(null); setComparison(null); setPhase(null);
      setElapsed(performance.now() - started.current);
      setTimes(t => ({openMs:t.openMs, calculateMs:t.calculateMs})); setEditBreakdown(null);
      setProgress("Undo verified: 120 launches and the original forecast restored.");
    } catch (e) { fail(e); }
  };
  const report = () => download(JSON.stringify({
    workbook: FILE, sha256: fixture!.sha256, bytes: fixture!.bytes.length, formulas: fixture!.formulas,
    populatedCells: fixture!.cells, forecastYears: fixture!.years, annualPeriods: fixture!.periods, sheets: fixture!.sheets, startedAt: openedAt.current, openTrigger: openTrigger.current,
    scenario: { launchesBefore: 120, launchesAfter: 180, baseline: fixture!.base, edited: fixture!.edited },
    userAgent: navigator.userAgent, logicalProcessors: navigator.hardwareConcurrency,
    gridpathVersion: "0.5.2", ...times, editBreakdown, summaryVerified: verified, chartPointsVerified: verified ? fixture!.chartPoints : 0, cachedFallbackFormulas: issues, saveComparison: comparison,
    methodology: {
      openMs: "Automatic workbook mount (or Reset workbook) to component onReady; fresh component and engine worker, including worker/WASM startup and workbook import. File fetch, formula counting and hash verification finish before timing begins. This is not page-navigation time. HTTP cache not cleared.",
      calculateMs: "Same start as openMs through whenCalculated; total, not additional time.",
      editMs: `Set Income statement!F5 from 120 to 180 paid standard launches in 2027, wait for calculation, independently verify twelve statement results (including unchanged modeled base years) and all ${fixture!.chartPoints} values across the native chart series, then two animation frames. No viewport scroll.`,
      saveMs: "sheet.save(), including recalculation wait and XLSX serialization; excludes zip comparison and download.",
      fidelity: "Original and saved ZIP parts compared by uncompressed length and CRC-32. Native chart source ranges and formatting retained; referenced data caches updated. This is not a claim of complete Excel compatibility."
    }
  }, null, 2), "gridpath-larkspur-income-timings.json", "application/json");

  const manualEdit = () => {
    setVerified(false); setEdited(false); setSaved(null); setComparison(null); setEditBreakdown(null);
    setTimes(t => ({openMs:t.openMs, calculateMs:t.calculateMs}));
    setProgress("Custom edits. Reset the workbook to run the verified 120 → 180 launch scenario.");
  };

  return <div className="recording">
    <header className="recording-appbar">
      <div className="recording-brand">
        <svg width="23" height="23" viewBox="0 0 24 24" fill="none" aria-hidden="true"><circle cx="12" cy="12" r="5" stroke="currentColor" strokeWidth="1.5"/><ellipse cx="12" cy="12" rx="11" ry="4" transform="rotate(-35 12 12)" stroke="currentColor" strokeWidth="1.5"/></svg>
        <span>Mission Control</span>
      </div>
      <nav className="recording-breadcrumb" aria-label="Workspace location"><span>Finance</span><span aria-hidden="true">/</span><span>Models</span></nav>
      <span className="recording-workspace-label">Demo workspace</span>
    </header>
    <section className="recording-document" aria-label="Larkspur Orbital financial model">
    <header className="recording-header">
      <svg className="recording-file-icon" width="18" height="20" viewBox="0 0 18 20" fill="none" aria-hidden="true"><rect x="2" y="1" width="14" height="18" rx="2" stroke="currentColor" strokeWidth="1.3"/><path d="M5 7h8M5 11h8M5 15h8M9 7v8" stroke="currentColor" strokeWidth="1.3"/></svg>
      <h1>Larkspur Orbital income statement</h1>
      <span className="recording-busy" role="status">{phase ? {opening: "Opening…", calculating: "Calculating…", editing: "Calculating…", saving: "Saving…"}[phase] : ""}</span>
      <button className="quiet" disabled={!api || !!phase} onClick={() => { api?.select(SHEET, chartView ? "D10" : "S10"); setChartView(show => !show); }}>{chartView ? "View statement" : "View chart"}</button>
      <button onClick={() => void save(true)} disabled={!api || !!phase}>Download .xlsx</button>
      <button className="quiet" aria-expanded={showTimings} aria-controls="recording-timings" onClick={() => setShowTimings(show => !show)}>Timings</button>
    </header>
    {error && <div className="recording-error" role="alert">{error}</div>}
    {showTimings && <section className="recording-panel" id="recording-timings" aria-label="Timing details">
      <div className="recording-heading"><h2>Timing details</h2><div className={`recording-clock ${phase ? "active" : ""}`} aria-live="off">{phase ?? "Elapsed"}: {seconds(elapsed)}</div></div>
      <div className="recording-actions">
        <button onClick={() => fixture && openFixture(fixture, "reset")} disabled={!fixture || !!phase}>Reset workbook</button>
        <button onClick={() => void edit()} disabled={!api || !!phase || !verified || edited}>Model 180 paid launches</button>
        <button onClick={() => void undo()} disabled={!edited || !!phase}>Undo scenario</button>
        <button onClick={() => void save()} disabled={!api || !!phase}>Save & compare</button>
        <button className="quiet" onClick={report} disabled={!times.calculateMs || !!phase}>Export timings</button>
        <button className="quiet" onClick={() => api?.select(SHEET, "F5")} disabled={!api || !!phase}>Select launch input</button>
      </div>
      <div className="recording-metrics">
        <div><span>Grid ready</span><strong>{seconds(times.openMs)}</strong></div>
        <div><span>Fully calculated · total</span><strong>{seconds(times.calculateMs)}</strong></div>
        <div><span>Edit → updated result</span><strong>{seconds(times.editMs)}</strong></div>
        <div><span>Save .xlsx</span><strong>{seconds(times.saveMs)}</strong></div>
      </div>
      <div className="recording-status" role="status">{phase ? progress : verified ? `${progress} ${issues === 0 ? "No cached fallback formulas." : `${issues} formulas using saved values.`}` : progress}</div>
      {comparison && <div className="recording-result"><strong>{comparison.unchanged}/{comparison.total} original ZIP parts unchanged.</strong> {comparison.removed.length} removed. {saved ? fileContents(saved).charts : 0} native Excel charts saved with updated values.
        <details><summary>Part comparison</summary><div>Rewritten: {comparison.rewritten.join(", ") || "none"}</div><div>Added: {comparison.added.join(", ") || "none"}</div><div>Removed: {comparison.removed.join(", ") || "none"}</div></details>
        {saved && <button className="quiet" onClick={() => download(saved as BlobPart, "larkspur-income-statement-edited.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}>Download saved copy</button>}
      </div>}
      <div className="recording-method">Measured on this workbook, at normal speed. File fetch excluded. Reset reloads the original workbook in a fresh engine and discards edits.</div>
      <a className="recording-original" href={`${BASE}${FILE}`} download>Download original workbook</a>
    </section>}
    <main className="recording-grid">{source ? <Spreadsheet key={run} file={source} fileName={FILE} store={false} toolbar={false} fileMenu={false} documentId={`recording-larkspur-income-${run}`} onReady={sheet => void ready(sheet)} onEdit={manualEdit} onProgress={setProgress} onError={e => fail(e)} /> : <div className="recording-empty">{error ? "Workbook unavailable." : "Opening workbook…"}</div>}</main>
    </section>
  </div>;
}
