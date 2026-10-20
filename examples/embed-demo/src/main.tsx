import React, { useCallback, useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { browserStore, httpStore, Spreadsheet, type Spreadsheet as SheetApi } from "gridpath";
import { ask, type AgentEvent, type Message } from "./agentLoop";

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
const BASE: string = (import.meta as any).env?.BASE_URL ?? "/";

/** The scripted "AI edit": the tool calls an agent would make, run through
 * sheet.ai like a real agent's, so they land as changes to review. */
async function sampleEdit(sheet: SheetApi): Promise<string | null> {
  if (!sheet.getSheetNames().includes("Sales")) return "The sample edit works on the sample workbook (File → Recent → Sample sales.xlsx).";
  const header = { bold: true, font_color: "#FFFFFF", background_color: "#1F4E79", horizontal_align: "center" };
  await sheet.ai.run("set_range", { sheet: "Sales", top_left: "H1", values: [["Share"], ...[2, 3, 4, 5, 6, 7, 8].map((r) => [`=E${r}/$E$8`])] });
  await sheet.ai.run("set_format", { sheet: "Sales", range: "H2:H8", format: { number_format: "0.0%" } });
  await sheet.ai.run("set_range", { sheet: "Sales", top_left: "I1", values: [["Rank"], ...[2, 3, 4, 5, 6, 7].map((r) => [`=RANK(E${r},$E$2:$E$7)`])] });
  await sheet.ai.run("set_format", { sheet: "Sales", range: "H1:I1", format: header });
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
  const [status, setStatus] = useState("Starting the engine…");
  const [sample, setSample] = useState<ArrayBuffer | null>(null);
  const describe = (api: SheetApi) => `${api.getDocument().name}: ${api.getSheetNames().length} sheet(s)`;
  useEffect(() => {
    if (SITE) void fetch(`${BASE}sample.xlsx`).then((r) => r.arrayBuffer()).then(setSample);
  }, []);

  return (
    <div
      style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "system-ui, sans-serif" }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault();
        const f = e.dataTransfer.files[0];
        if (f && sheet) void f.arrayBuffer().then((b) => sheet.open(b, { name: f.name }));
      }}
    >
      {SITE ? (
        <header style={{ display: "flex", flexWrap: "wrap", gap: 12, alignItems: "center", padding: "10px 14px", borderBottom: "1px solid #ddd", fontSize: 14 }}>
          <a href="/" style={{ fontWeight: 700, color: "#111", textDecoration: "none" }}>GridPath</a>
          <span style={{ color: "#555" }}>The embeddable spreadsheet. Open, edit and save real .xlsx files; everything stays in this browser.</span>
          <span style={{ flex: 1 }} />
          <button
            disabled={!sheet}
            title="Runs the tool calls an agent would make (scripted; no AI model is called)"
            onClick={() => void (sheet && sampleEdit(sheet).then((msg) => msg && alert(msg)))}
            style={{ padding: "6px 12px", borderRadius: 6, border: "1px solid #1F4E79", background: "#1F4E79", color: "#fff", cursor: "pointer" }}
          >
            Try a sample AI edit
          </button>
          <a href="/embed" style={{ color: "#1F4E79" }}>Embed it in your app →</a>
        </header>
      ) : (
        <header style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", borderBottom: "1px solid #ddd", fontSize: 14 }}>
          <strong>GridPath embed demo</strong>
          <span style={{ color: "#666" }}>{status}</span>
        </header>
      )}
      <div style={{ flex: 1, minHeight: 0, display: "flex" }}>
        <main style={{ flex: 1, minWidth: 0 }}>
          {/* The toolbar, File menu (open, new, recent, version history),
              save status and Download are the component's own; it saves in
              this browser as you work and reopens where you left off. */}
          {SITE && !sample ? null : (
          <Spreadsheet
            engine={engineOptions}
            store={store}
            {...(SITE && sample ? { file: sample, fileName: "Sample sales.xlsx", documentId: "gridpath-demo-sample" } : {})}
            onProgress={setStatus}
            onReady={(api) => {
              setSheet(api);
              setStatus(describe(api));
              (window as any).__demo = { sheet: api, ready: true };
            }}
            onDocumentChange={() => {
              const api = (window as any).__demo?.sheet as SheetApi | undefined;
              if (api) setStatus(describe(api));
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

createRoot(document.getElementById("root")!).render(<App />);
