"use client";
// The app: the signed-in user's workbooks beside a spreadsheet. Each
// workbook has its own URL (/books/<id>); going to another one remounts the
// spreadsheet, which saves the workbook you leave and opens the next one by
// its document ID from your API.
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useCallback, useEffect, useMemo, useState } from "react";
import { Spreadsheet, httpStore, type GridPathError, type StoredWorkbook, type WorkbookStore } from "gridpath";
import { getAccessToken } from "./session";
import { forgetUpload, keepUpload, uploadFor } from "./uploads";

export function Workbooks({ id }: { id?: string }) {
  const router = useRouter();
  // Saves go to /api/books as the signed-in user; the token is read before
  // every request, so a refreshed one is picked up.
  const store = useMemo(
    () => httpStore({ url: "/api/books", headers: async () => ({ Authorization: `Bearer ${await getAccessToken()}` }) }),
    [],
  );
  const [books, setBooks] = useState<StoredWorkbook[] | null>(null);
  const [listError, setListError] = useState<string | null>(null);
  const refresh = useCallback(
    () =>
      store.list().then(
        (list) => {
          setBooks(list);
          setListError(null);
        },
        (e: Error) => setListError(e.message),
      ),
    [store],
  );
  useEffect(() => void refresh(), [refresh]);

  const upload = async (file: File) => {
    const next = crypto.randomUUID();
    keepUpload(next, { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
    router.push(`/books/${next}`);
  };
  const remove = async (book: StoredWorkbook) => {
    if (!window.confirm(`Delete "${book.name}"?`)) return;
    await store.remove?.(book.id);
    if (book.id === id) router.push("/");
    void refresh();
  };

  return (
    <div style={{ display: "flex", height: "100vh" }}>
      <aside style={{ width: 240, flex: "none", borderRight: "1px solid #e3e5e8", display: "flex", flexDirection: "column", fontSize: 13 }}>
        <div style={{ padding: "12px 14px", borderBottom: "1px solid #eee" }}>
          <Link href="/" style={{ fontWeight: 600, color: "#111", textDecoration: "none" }}>
            Your workbooks
          </Link>
          <div style={{ display: "flex", gap: 6, marginTop: 10 }}>
            <button onClick={() => router.push(`/books/${crypto.randomUUID()}`)}>New</button>
            <label style={{ border: "1px solid #ccc", borderRadius: 4, padding: "1px 8px", cursor: "pointer", background: "#f6f6f6" }}>
              Upload .xlsx
              <input
                type="file"
                accept=".xlsx,.xlsm,.csv,.tsv,.txt"
                hidden
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  e.target.value = "";
                  if (f) void upload(f);
                }}
              />
            </label>
          </div>
        </div>
        <nav style={{ flex: 1, overflowY: "auto" }}>
          {listError && <p style={{ padding: "0 14px", color: "#b00020" }}>Couldn't load the list: {listError}</p>}
          {books === null && !listError && <p style={{ padding: "0 14px", color: "#888" }}>Loading…</p>}
          {books?.length === 0 && <p style={{ padding: "0 14px", color: "#888" }}>No workbooks yet: start a new one or upload an .xlsx.</p>}
          {books?.map((b) => (
            <div key={b.id} style={{ display: "flex", alignItems: "center", background: b.id === id ? "#eef3fb" : undefined }}>
              <Link href={`/books/${b.id}`} title={b.name} style={{ flex: 1, minWidth: 0, padding: "8px 14px", color: "#111", textDecoration: "none" }}>
                <div style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{b.name}</div>
                <div style={{ color: "#888", fontSize: 11 }}>{new Date(b.savedAt).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}</div>
              </Link>
              <button title={`Delete ${b.name}`} onClick={() => void remove(b)} style={{ marginRight: 8, border: 0, background: "none", color: "#999", cursor: "pointer" }}>
                ✕
              </button>
            </div>
          ))}
        </nav>
      </aside>
      <main style={{ flex: 1, minWidth: 0, position: "relative", display: "flex", flexDirection: "column" }}>
        {id ? (
          // A new key per workbook: switching unmounts this one (saving what's left) and mounts the next.
          <Editor key={id} id={id} store={store} onSaved={refresh} />
        ) : (
          <p style={{ margin: "auto", color: "#888" }}>Pick a workbook, start a new one, or upload an .xlsx.</p>
        )}
      </main>
    </div>
  );
}

function Editor({ id, store, onSaved }: { id: string; store: WorkbookStore; onSaved: () => void }) {
  // A file just uploaded for this id; otherwise the workbook comes from the
  // store (a new id opens a blank one).
  const [upload] = useState(() => uploadFor(id));
  const [stage, setStage] = useState<string | null>("Starting…");
  const [error, setError] = useState<GridPathError | null>(null);
  return (
    <>
      {error && <ErrorBar error={error} />}
      <div style={{ flex: 1, minHeight: 0, position: "relative" }}>
        <Spreadsheet
          store={store}
          documentId={id}
          file={upload?.bytes}
          fileName={upload?.name ?? "Untitled.xlsx"}
          // The app has its own workbook list; the toolbar keeps Version history and Download.
          fileMenu={false}
          onProgress={setStage}
          onReady={(sheet) => {
            setStage(null);
            forgetUpload(id);
            // Not in the store yet (an upload, a new workbook): save it there now.
            if (sheet.getStatus().state === "idle") void sheet.saveNow();
            (window as any).__sheet = sheet; // for poking at it from the devtools console
          }}
          onError={(e) => {
            setStage(null);
            setError(e);
          }}
          onStatusChange={(s) => s.state === "saved" && onSaved()}
        />
        {stage && (
          <div style={{ position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(255,255,255,0.7)", color: "#555" }}>
            {stage}
          </div>
        )}
      </div>
    </>
  );
}

/** The spreadsheet explains the problem itself; this adds the app's way out, by error code. */
function ErrorBar({ error }: { error: GridPathError }) {
  let action: React.ReactNode = null;
  switch (error.code) {
    case "unsupported-file":
    case "malformed-file":
      action = <Link href="/">Back to your workbooks</Link>;
      break;
    case "storage-failed":
    case "engine-failed":
      action = <button onClick={() => location.reload()}>Reload</button>;
      break;
  }
  return (
    <div role="status" style={{ display: "flex", gap: 10, alignItems: "center", padding: "6px 12px", background: "#fdecea", color: "#7a1c12", fontSize: 13 }}>
      <span>{error.message}</span>
      <span style={{ flex: 1 }} />
      {action}
    </div>
  );
}
