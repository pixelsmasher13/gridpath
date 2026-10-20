/**
 * A WorkbookStore over your own HTTP API: workbooks saved to the signed-in
 * user's account instead of this browser. The endpoints it calls (all
 * relative to `url`; ids are URL-encoded):
 *
 *   PUT    /{id}?name=…&source=…[&checkpoint=1]   body: the .xlsx bytes
 *   GET    /{id}                                  → the .xlsx bytes, with headers
 *                                                    X-Workbook-Name (percent-encoded, as
 *                                                    encodeURIComponent writes it),
 *                                                    X-Workbook-Saved-At (ISO time),
 *                                                    X-Workbook-Source; 404 if none
 *   GET    /                                      → [{ id, name, savedAt, size }], newest first
 *   GET    /{id}/versions                         → [{ id, savedAt, size }], newest first
 *                                                    (404: the API keeps no history)
 *   GET    /{id}/versions/{versionId}             → that version's .xlsx bytes
 *   DELETE /{id}
 *
 * Authenticate with `headers` (called before every request, so a fresh
 * token each time). A failed request throws with its `status`; the
 * spreadsheet then keeps the edits in this browser and retries.
 */
import type { StoredWorkbook, WorkbookStore, WorkbookVersion } from "./store";

export interface HttpStoreOptions {
  /** Your workbook API, e.g. "/api/books". */
  url: string;
  /** Headers for every request (an Authorization token…); may be async. */
  headers?: HeadersInit | (() => HeadersInit | Promise<HeadersInit>);
  /** The status line's "Saved …" wording; default "to your account". */
  label?: string;
  /** A fetch to use instead of the global one. */
  fetch?: typeof fetch;
}

/** An HTTP failure, with the response status for the app to act on. */
export class HttpStoreError extends Error {
  readonly status: number;
  constructor(status: number, message: string) {
    super(message);
    this.name = "HttpStoreError";
    this.status = status;
  }
}

/** Header values are ASCII: names travel percent-encoded. */
function decodeName(raw: string | null): string | null {
  if (!raw) return null;
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

export function httpStore(options: HttpStoreOptions): WorkbookStore {
  const base = options.url.replace(/\/+$/, "");
  const doFetch = options.fetch ?? ((...a: Parameters<typeof fetch>) => fetch(...a));
  const call = async (method: string, path: string, body?: Uint8Array): Promise<Response> => {
    const headers = new Headers(typeof options.headers === "function" ? await options.headers() : options.headers);
    if (body) headers.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    const res = await doFetch(`${base}${path}`, { method, headers, body: body as BodyInit | undefined });
    return res;
  };
  const fail = async (res: Response): Promise<never> => {
    const text = (await res.text().catch(() => "")).trim();
    throw new HttpStoreError(res.status, `HTTP ${res.status}${text ? `: ${text.slice(0, 200)}` : ""}`);
  };
  /** A page instead of data: the url misses the API (a dev server's fallback page, a login redirect). */
  const notHtml = (res: Response, path: string) => {
    if ((res.headers.get("Content-Type") ?? "").includes("text/html")) {
      throw new HttpStoreError(res.status, `${base}${path} returned a web page, not workbook data: check the store's url`);
    }
  };
  const enc = encodeURIComponent;
  return {
    label: options.label ?? "to your account",
    async save({ id, name, bytes, source, checkpoint }) {
      const q = new URLSearchParams({ name, ...(source ? { source } : {}), ...(checkpoint ? { checkpoint: "1" } : {}) });
      const res = await call("PUT", `/${enc(id)}?${q}`, bytes);
      if (!res.ok) await fail(res);
    },
    async load(id) {
      const res = await call("GET", `/${enc(id)}`);
      if (res.status === 404) return null;
      if (!res.ok) await fail(res);
      notHtml(res, `/${enc(id)}`);
      return {
        id,
        name: decodeName(res.headers.get("X-Workbook-Name")) ?? id,
        bytes: new Uint8Array(await res.arrayBuffer()),
        savedAt: res.headers.get("X-Workbook-Saved-At") ?? new Date(0).toISOString(),
        source: res.headers.get("X-Workbook-Source") ?? undefined,
      };
    },
    async list(): Promise<StoredWorkbook[]> {
      const res = await call("GET", "");
      if (!res.ok) await fail(res);
      notHtml(res, "");
      return res.json();
    },
    async versions(id): Promise<WorkbookVersion[]> {
      const res = await call("GET", `/${enc(id)}/versions`);
      if (res.status === 404) return [];
      if (!res.ok) await fail(res);
      notHtml(res, `/${enc(id)}/versions`);
      return res.json();
    },
    async loadVersion(id, versionId) {
      const res = await call("GET", `/${enc(id)}/versions/${enc(versionId)}`);
      if (res.status === 404) return null;
      if (!res.ok) await fail(res);
      notHtml(res, `/${enc(id)}/versions/${enc(versionId)}`);
      return new Uint8Array(await res.arrayBuffer());
    },
    async remove(id) {
      const res = await call("DELETE", `/${enc(id)}`);
      if (!res.ok && res.status !== 404) await fail(res);
    },
  };
}
