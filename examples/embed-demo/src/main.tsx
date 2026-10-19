import React, { useCallback, useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { Spreadsheet, type AgentBatch, type Spreadsheet as SheetApi } from "gridpath";
import { ask, type AgentEvent, type Message } from "./agentLoop";

type Source = { key: number; name: string; file?: ArrayBuffer };
type LogLine = { who: "you" | "agent" | "tool" | "error"; text: string };

// ?worker=0 runs the engine on the page instead of a Web Worker.
const engineOptions = new URLSearchParams(location.search).get("worker") === "0" ? { worker: false } : undefined;

function AgentPanel({ sheet }: { sheet: SheetApi | null }) {
  const [prompt, setPrompt] = useState("");
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<LogLine[]>([]);
  const [batches, setBatches] = useState<AgentBatch[]>([]);
  const history = useRef<Message[]>([]);

  useEffect(() => {
    history.current = [];
    setLog([]);
    setBatches(sheet?.ai.list() ?? []);
    return sheet?.ai.onChange(setBatches);
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

  const pending = batches.filter((b) => b.status === "pending");
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
      <div style={{ padding: 10, borderTop: "1px solid #eee" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
          <strong style={{ flex: 1 }}>Changes{pending.length ? ` (${pending.length} pending)` : ""}</strong>
          <button disabled={!pending.length} onClick={() => sheet?.ai.accept()}>Accept all</button>
          <button disabled={!pending.length} onClick={() => void sheet?.ai.reject()}>Reject all</button>
        </div>
        <div style={{ maxHeight: 160, overflow: "auto" }}>
          {batches.map((b) => (
            <div key={b.id} style={{ display: "flex", gap: 6, alignItems: "center", color: b.status === "pending" ? "#7a5c00" : "#666" }}>
              <span style={{ flex: 1 }}>
                {b.tool} · {b.cells.length} cell{b.cells.length === 1 ? "" : "s"}
              </span>
              {b.status === "pending" ? (
                <button title="Undo this change and every change after it" onClick={() => void sheet?.ai.reject(b.id)}>Undo from here</button>
              ) : (
                <span>accepted</span>
              )}
            </div>
          ))}
        </div>
      </div>
    </aside>
  );
}

function App() {
  const [source, setSource] = useState<Source>({ key: 0, name: "Untitled.xlsx" });
  const [sheet, setSheet] = useState<SheetApi | null>(null);
  const [status, setStatus] = useState("Starting the engine…");
  const [edits, setEdits] = useState(0);

  const open = useCallback(async (f: File) => {
    setSheet(null);
    setEdits(0);
    setSource((s) => ({ key: s.key + 1, name: f.name, file: undefined }));
    const buf = await f.arrayBuffer();
    setSource((s) => ({ ...s, file: buf }));
  }, []);

  const download = useCallback(async () => {
    if (!sheet) return;
    setStatus("Saving…");
    const blob = await sheet.saveBlob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = source.name;
    a.click();
    URL.revokeObjectURL(url);
    setStatus(`Saved ${source.name} (${(blob.size / 1024).toFixed(0)} KB)`);
  }, [sheet, source.name]);

  // A file being read: wait for its bytes before mounting the grid.
  const waiting = source.key > 0 && source.name !== "Untitled.xlsx" && !source.file;

  return (
    <div
      style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "system-ui, sans-serif" }}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault();
        const f = e.dataTransfer.files[0];
        if (f) void open(f);
      }}
    >
      <header style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 12px", borderBottom: "1px solid #ddd", fontSize: 14 }}>
        <strong>GridPath embed demo</strong>
        <label style={{ cursor: "pointer", padding: "4px 10px", border: "1px solid #bbb", borderRadius: 4 }}>
          Open .xlsx
          <input type="file" accept=".xlsx,.xlsm" hidden onChange={(e) => e.target.files?.[0] && void open(e.target.files[0])} />
        </label>
        <button
          onClick={() => {
            setSheet(null);
            setEdits(0);
            setSource((s) => ({ key: s.key + 1, name: "Untitled.xlsx" }));
          }}
        >
          New blank
        </button>
        <button disabled={!sheet} onClick={() => void download()}>
          Download .xlsx
        </button>
        <span style={{ color: "#666" }}>
          {status}
          {edits > 0 ? ` · ${edits} edit${edits === 1 ? "" : "s"}` : ""}
        </span>
      </header>
      <div style={{ flex: 1, minHeight: 0, display: "flex" }}>
        <main style={{ flex: 1, minWidth: 0 }}>
          {!waiting && (
            <Spreadsheet
              key={source.key}
              file={source.file}
              fileName={source.name}
              engine={engineOptions}
              onProgress={setStatus}
              onEdit={() => setEdits((n) => n + 1)}
              onReady={(api) => {
                setSheet(api);
                setStatus(`${source.name}: ${api.getSheetNames().length} sheet(s)`);
                (window as any).__demo = { sheet: api, ready: true };
              }}
              onError={(err) => {
                setStatus(`Error: ${err.message}`);
                (window as any).__demo = { error: err.message, ready: true };
              }}
            />
          )}
        </main>
        <AgentPanel sheet={sheet} />
      </div>
    </div>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
