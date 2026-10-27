import React, { useCallback, useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { browserStore, httpStore, Spreadsheet, type Spreadsheet as SheetApi } from "gridpath";
import { ask, type AgentEvent, type Message } from "./agentLoop";
import { analyse, BLUE, DropOverlay, isWorkbookFile, OpenButton, ReportBar, useFileDrop, type Opened } from "./Report";
import { RecordingDemo } from "./RecordingDemo";

type LogLine = { who: "you" | "agent" | "tool" | "error"; text: string };

// ?worker=0 runs the engine on the page instead of a Web Worker;
// ?engine=broken points it at a worker script that doesn't exist.
const params = new URLSearchParams(location.search);
const engineOptions =
  params.get("engine") === "broken" ? { workerUrl: "/missing-engine.worker.js" } : params.get("worker") === "0" ? { worker: false } : undefined;
// ?versions=all keeps a version for every save (default: one per 5 minutes);
// ?store=none saves nowhere (Download only); ?store=http saves to /api/books.
const store =
  params.get("store") === "none"
    ? false
    : params.get("store") === "http"
      ? httpStore({ url: "/api/books" })
      : params.get("versions") === "all"
        ? browserStore({ checkpointMinutes: 0 })
        : undefined;
// Site mode (VITE_SITE=1, the build at gridpath.dev/demo): a static page with
// no server, so no Ask panel; a sample workbook, and a scripted AI edit that
// shows the review flow without calling a model.
const SITE = (import.meta as any).env?.VITE_SITE === "1";
// ?embed=1 (inside the /embed page): a slim bar instead of the site header.
const EMBED = SITE && params.get("embed") === "1";
const BASE: string = (import.meta as any).env?.BASE_URL ?? "/";
// The site's sample: a fictional three-sheet income statement with native charts.
const SAMPLE_FILE = "larkspur-income-statement.xlsx";
const SAMPLE_NAME = "Larkspur Orbital income statement.xlsx";

/** The scripted "AI edit": the tool calls an agent would make, run through
 * sheet.ai like a real agent's, so they land as changes to review. */
async function sampleEdit(sheet: SheetApi): Promise<string | null> {
  const IS = "Income statement";
  if (!sheet.getSheetNames().includes(IS)) return `The sample AI edit works on the sample workbook: reopen it from File → Recent → ${SAMPLE_NAME}.`;
  // An EPS growth row under diluted EPS, styled like the statement's other growth rows.
  const growth = ["E", "F", "G", "H", "I", "J"].map((c, i) => {
    const prev = "DEFGHI"[i];
    return `=IFERROR(IF(${prev}43<=0,"n.a.",${c}43/${prev}43-1),"n.a.")`;
  });
  await sheet.ai.run("set_range", { sheet: IS, top_left: "C44", values: [["Diluted EPS growth", "n.a.", ...growth]] });
  await sheet.ai.run("set_format", { sheet: IS, range: "C44:J44", format: { italic: true, font_color: "#61748A" } });
  await sheet.ai.run("set_format", { sheet: IS, range: "D44:J44", format: { number_format: '0.0%;(0.0%);"–"', horizontal_align: "right" } });
  // Bring the new row into view, highlighted for review.
  sheet.select(IS, "C44");
  return null;
}

function AgentPanel({ sheet }: { sheet: SheetApi | null }) {
  const [prompt, setPrompt] = useState("");
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<LogLine[]>([]);
  const history = useRef<Message[]>([]);

  useEffect(() => {
    history.current = [];
    setLog([]);
  }, [sheet]);

  const submit = useCallback(async () => {
    if (!sheet || !prompt.trim() || busy) return;
    const text = prompt.trim();
    setPrompt("");
    setBusy(true);
    setLog((l) => [...l, { who: "you", text }]);
    const onEvent = (e: AgentEvent) =>
      setLog((l) => [...l, e.kind === "text" ? { who: "agent", text: e.text } : e.kind === "tool" ? { who: "tool", text: `${e.ok ? "✓" : "✗"} ${e.name}` } : { who: "error", text: e.text }]);
    try {
      history.current = await ask(sheet, history.current, text, onEvent);
    } catch (e) {
      onEvent({ kind: "error", text: e instanceof Error ? e.message : String(e) });
    } finally {
      setBusy(false);
    }
  }, [sheet, prompt, busy]);

  return (
    <aside style={{ width: 340, borderLeft: "1px solid #ddd", display: "flex", flexDirection: "column", fontSize: 13 }}>
      <div style={{ padding: 10, borderBottom: "1px solid #eee" }}>
        <textarea
          value={prompt}
          onChange={(e) => setPrompt(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) void submit(); }}
          placeholder="Ask the sheet… (⌘↩ to send)"
          rows={3}
          style={{ width: "100%", boxSizing: "border-box", font: "inherit" }}
          disabled={!sheet}
        />
        <button onClick={() => void submit()} disabled={!sheet || busy || !prompt.trim()} style={{ marginTop: 6 }}>
          {busy ? "Working…" : "Ask"}
        </button>
      </div>
      <div style={{ flex: 1, overflow: "auto", padding: 10 }}>
        {log.map((l, i) => (
          <div key={i} style={{ marginBottom: 6, color: l.who === "error" ? "#b00020" : l.who === "tool" ? "#666" : "#111", whiteSpace: "pre-wrap", fontFamily: l.who === "tool" ? "ui-monospace, monospace" : undefined }}>
            {l.who === "you" ? <strong>You: </strong> : null}
            {l.text}
          </div>
        ))}
      </div>
      {/* Reviewing the agent's changes (accept, reject, jump to cells) is the
          spreadsheet's own "AI changes" panel. */}
    </aside>
  );
}

function App() {
  const [sheet, setSheet] = useState<SheetApi | null>(null);
  const sheetRef = useRef<SheetApi | null>(null);
  const [status, setStatus] = useState("Starting the engine…");
  const [sample, setSample] = useState<Uint8Array | null>(null);
  const sampleT0 = useRef(0);
  // The site's file report: the open workbook, and edits since the last check.
  const [opened, setOpened] = useState<Opened | null>(null);
  const [stale, setStale] = useState(0);
  const [message, setMessage] = useState<string | null>(null);
  const opening = useRef(false);
  const describe = (api: SheetApi) => `${api.getDocument().name}: ${api.getSheetNames().length} sheet(s)`;
  useEffect(() => {
    if (!SITE) return;
    void fetch(`${BASE}${SAMPLE_FILE}`)
      .then((r) => r.arrayBuffer())
      .then((b) => {
        sampleT0.current = performance.now();
        setSample(new Uint8Array(b));
      });
  }, []);

  /** A file from the button or a drop: opened here, then reported on. */
  const openFile = useCallback(async (f: File) => {
    const api = sheetRef.current;
    if (!api) return;
    const bytes = new Uint8Array(await f.arrayBuffer());
    setMessage(null);
    const t0 = performance.now();
    opening.current = true;
    try {
      await api.open(bytes, { name: f.name });
    } catch (e: any) {
      setMessage(e?.code === "pending-changes" ? "Accept or reject the AI's changes first (the AI changes panel), then open the file." : `${e?.message ?? e}. ${e?.hint ?? ""}`);
      return;
    } finally {
      opening.current = false;
    }
    // A CSV opens as a new workbook: there is no original .xlsx to report on.
    if (SITE) isWorkbookFile(bytes) ? void analyse(api, f.name, bytes, t0, setOpened) : setOpened(null);
  }, []);
  const dragging = useFileDrop((f) => void openFile(f));

  const aiEdit = () => void (sheet && sampleEdit(sheet).then((msg) => setMessage(msg)));
  const secondary = { padding: EMBED ? "4px 10px" : "6px 12px", borderRadius: 6, border: `1px solid ${BLUE}`, background: "#fff", color: BLUE, cursor: "pointer", fontSize: EMBED ? 13 : 14 } as const;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "system-ui, sans-serif" }}>
      {dragging && <DropOverlay />}
      {EMBED ? (
        // Inside the /embed page: the live component with a slim bar.
        <header style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center", padding: "6px 12px", borderBottom: "1px solid #ddd", fontSize: 13 }}>
          <OpenButton small onFile={(f) => void openFile(f)} disabled={!sheet} />
          <button disabled={!sheet} title="Runs the tool calls an agent would make (scripted; no AI model is called)" onClick={aiEdit} style={secondary}>
            Try a sample AI edit
          </button>
          <span style={{ flex: 1 }} />
          <span style={{ color: "#666" }}>The real component, running in this page</span>
          <a href="/demo" target="_top" style={{ color: BLUE }}>Full demo ↗</a>
        </header>
      ) : SITE ? (
        <header style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", padding: "10px 14px", borderBottom: "1px solid #ddd", fontSize: 14 }}>
          <a href="/" style={{ fontWeight: 700, color: "#111", textDecoration: "none" }}>GridPath</a>
          <span className="demo-tag" style={{ color: "#555" }}>Open a real .xlsx, edit it, save it back. Your file never leaves this browser.</span>
          <span style={{ flex: 1 }} />
          <OpenButton onFile={(f) => void openFile(f)} disabled={!sheet} />
          <button disabled={!sheet} title="Runs the tool calls an agent would make (scripted; no AI model is called)" onClick={aiEdit} style={secondary}>
            Try a sample AI edit
          </button>
          <a href="/embed" style={{ color: BLUE }}>Embed it in your app →</a>
        </header>
      ) : (
        <header style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", borderBottom: "1px solid #ddd", fontSize: 14 }}>
          <strong>GridPath embed demo</strong>
          <span style={{ color: "#666" }}>{status}</span>
        </header>
      )}
      {message && (
        <div role="status" style={{ padding: "6px 14px", background: "#fff8e6", borderBottom: "1px solid #f3e3b5", color: "#6a4a00", fontSize: 13, display: "flex", gap: 8 }}>
          <span style={{ flex: 1 }}>{message}</span>
          <button onClick={() => setMessage(null)} aria-label="Dismiss" style={{ border: 0, background: "none", cursor: "pointer", color: "#6a4a00" }}>✕</button>
        </div>
      )}
      {SITE && sheet && opened && <ReportBar sheet={sheet} opened={opened} compact={EMBED} stale={stale} />}
      <div style={{ flex: 1, minHeight: 0, display: "flex" }}>
        <main style={{ flex: 1, minWidth: 0 }}>
          {/* The toolbar, File menu (open, new, recent, version history),
              save status and Download are the component's own; it saves in
              this browser as you work and reopens where you left off. */}
          {SITE && !sample ? null : (
          <Spreadsheet
            engine={engineOptions}
            store={store}
            {...(SITE && sample ? { file: sample, fileName: SAMPLE_NAME, documentId: "gridpath-demo-larkspur" } : {})}
            onProgress={setStatus}
            onReady={(api) => {
              setSheet(api);
              sheetRef.current = api;
              setStatus(describe(api));
              (window as any).__demo = { sheet: api, ready: true };
              if (SITE && sample) void analyse(api, api.getDocument().name, sample, sampleT0.current, setOpened);
            }}
            onDocumentChange={() => {
              const api = (window as any).__demo?.sheet as SheetApi | undefined;
              if (api) setStatus(describe(api));
              // Opened from the File menu: there are no original bytes to compare against.
              if (!opening.current) setOpened(null);
            }}
            onStatusChange={(s) => {
              if (s.state === "unsaved" || s.state === "review") setStale((v) => v + 1);
            }}
            onError={(err) => {
              setStatus(`Error: ${err.message}`);
              (window as any).__demo = { ...(window as any).__demo, error: err.message, code: err.code, ready: true };
            }}
          />
          )}
        </main>
        {!SITE && <AgentPanel sheet={sheet} />}
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")!).render(params.get("record") === "1" ? <RecordingDemo /> : <App />);
