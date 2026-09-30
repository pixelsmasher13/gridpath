/**
 * Local HTTP server for the review tab. Loopback only, plain node:http.
 *
 *   GET  /                      review page (static; served from ./ui when built)
 *   GET  /api/workbooks         open workbooks with pending-batch counts
 *   GET  /api/workbook?path=    snapshot + batches + per-cell diff for the review UI
 *   POST /api/batch/accept      {path, batch_id?}   (no batch_id = accept all)
 *   POST /api/batch/reject      {path, batch_id}
 *   POST /api/save              {path}
 *
 * Requests must carry the per-process token (in the URL the agent is given)
 * and come from a loopback Origin or none, so a web page can't drive it.
 */
import crypto from "node:crypto";
import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { ToolHandlers } from "../tools/handlers";
import { cellKey } from "../workbook/snapshot";

export type ReviewServer = { url: string; token: string; port: number; close(): Promise<void> };

const json = (res: http.ServerResponse, status: number, body: unknown) => {
  res.writeHead(status, { "content-type": "application/json", "cache-control": "no-store" });
  res.end(JSON.stringify(body));
};

function readBody(req: http.IncomingMessage): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    let s = "";
    req.on("data", (c) => (s += c));
    req.on("end", () => {
      try {
        resolve(s ? JSON.parse(s) : {});
      } catch (e) {
        reject(e);
      }
    });
    req.on("error", reject);
  });
}

function uiDir(): string | null {
  const here = path.dirname(fileURLToPath(import.meta.url));
  for (const c of [path.resolve(here, "../../ui"), path.resolve(here, "../ui")]) {
    if (fs.existsSync(path.join(c, "index.html"))) return c;
  }
  return null;
}

/** Per-cell before/after for every pending batch of one workbook. */
export function diffForReview(handlers: ToolHandlers, workbookPath: string) {
  const s = handlers.session(workbookPath);
  const batches = s.batches.map((b) => {
    const cells: Array<{ sheet: string; row: number; col: number; before: unknown; after: unknown; before_formula: string | null; after_formula: string | null }> = [];
    for (const m of b.mutations) {
      if (m.type === "set_cell") {
        cells.push({ sheet: m.address.sheet, row: m.address.row, col: m.address.col, before: m.old_value ?? null, after: m.new_value ?? null, before_formula: m.old_formula ?? null, after_formula: m.new_formula ?? null });
      } else if (m.type === "set_range") {
        m.values.forEach((rowVals, dr) =>
          rowVals.forEach((v, dc) => {
            if (v === null || v === undefined || v === "") return;
            cells.push({ sheet: m.sheet, row: m.start_row + dr, col: m.start_col + dc, before: m.old_values?.[dr]?.[dc] ?? null, after: v, before_formula: null, after_formula: typeof v === "string" && v.startsWith("=") ? v : null });
          }),
        );
      } else if (m.type === "clear_range") {
        for (const c of m.cells) cells.push({ sheet: m.sheet, row: c.row, col: c.col, before: c.old_value ?? null, after: null, before_formula: c.old_formula ?? null, after_formula: null });
      }
    }
    const structural = b.mutations.filter((m) => !["set_cell", "set_range", "clear_range", "set_format"].includes(m.type)).map((m) => m.type);
    const formats = b.mutations.filter((m) => m.type === "set_format").length;
    return { id: b.id, tool: b.prompt, justification: b.justification, status: b.status, created_at: b.created_at, cells, formats, structural };
  });
  const snapshot = s.snapshot();
  const mapped = (s as any).mappedSnapshot?.() ?? null;
  const display: Record<string, string> = {};
  if (mapped?.display) for (const [k, v] of mapped.display as Map<string, string>) display[k] = v;
  return { path: workbookPath, sheetNames: s.sheetNames(), snapshot, display, batches, cellKey: "sheet\\u0000row\\u0000col" };
}

export async function startReviewServer(handlers: ToolHandlers, opts: { port?: number } = {}): Promise<ReviewServer> {
  const token = crypto.randomBytes(12).toString("hex");
  const ui = uiDir();

  const server = http.createServer(async (req, res) => {
    try {
      const url = new URL(req.url ?? "/", "http://127.0.0.1");
      const origin = req.headers.origin;
      if (origin && !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return json(res, 403, { error: "forbidden origin" });

      if (url.pathname.startsWith("/api/")) {
        const t = url.searchParams.get("token") ?? req.headers["x-gridpath-token"];
        if (t !== token) return json(res, 401, { error: "bad token" });
        if (req.method === "GET" && url.pathname === "/api/workbooks") {
          return json(res, 200, { workbooks: handlers.openWorkbooks() });
        }
        if (req.method === "GET" && url.pathname === "/api/workbook") {
          const p = url.searchParams.get("path") ?? "";
          return json(res, 200, diffForReview(handlers, p));
        }
        if (req.method === "POST" && url.pathname === "/api/batch/accept") {
          const b = await readBody(req);
          return json(res, 200, handlers.acceptBatch(String(b.path ?? ""), b.batch_id == null ? null : String(b.batch_id)));
        }
        if (req.method === "POST" && url.pathname === "/api/batch/reject") {
          const b = await readBody(req);
          return json(res, 200, await handlers.call("reject_batch", { path: b.path, batch_id: b.batch_id }));
        }
        if (req.method === "POST" && url.pathname === "/api/save") {
          const b = await readBody(req);
          return json(res, 200, await handlers.saveFromReview(String(b.path ?? "")));
        }
        return json(res, 404, { error: "not found" });
      }

      // Static review page.
      if (!ui) {
        res.writeHead(200, { "content-type": "text/html" });
        return res.end(`<!doctype html><title>gridpath review</title><p>Review UI not built. API is live at /api/*.</p>`);
      }
      const rel = url.pathname === "/" || url.pathname.startsWith("/review") ? "index.html" : url.pathname.slice(1);
      const file = path.join(ui, rel);
      if (!file.startsWith(ui) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) return json(res, 404, { error: "not found" });
      const ext = path.extname(file);
      const type = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".wasm": "application/wasm", ".svg": "image/svg+xml", ".woff2": "font/woff2" }[ext] ?? "application/octet-stream";
      res.writeHead(200, { "content-type": type });
      fs.createReadStream(file).pipe(res);
    } catch (e: unknown) {
      json(res, 500, { error: e instanceof Error ? e.message : String(e) });
    }
  });

  await new Promise<void>((resolve) => server.listen(opts.port ?? 0, "127.0.0.1", resolve));
  const port = (server.address() as { port: number }).port;
  return {
    url: `http://127.0.0.1:${port}`,
    token,
    port,
    close: () => new Promise<void>((r) => server.close(() => r())),
  };
}

export { cellKey };
