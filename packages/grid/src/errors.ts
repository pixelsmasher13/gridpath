/**
 * Errors the spreadsheet reports: each has a `code` an app can act on, a
 * message that says what happened, and a `hint` that says what to do.
 */

export type GridPathErrorCode =
  /** Not a spreadsheet: an old .xls, a password-protected workbook, a PDF… */
  | "unsupported-file"
  /** An .xlsx (zip) the engine could not read (damaged or truncated), or a
   * CSV that could not be read as a table. */
  | "malformed-file"
  /** The calculation engine (WebAssembly in a Web Worker) did not start or
   * stopped: the page can't load the worker or .wasm, or the engine failed. */
  | "engine-failed"
  /** The store (your backend, or this browser's storage) failed. */
  | "storage-failed"
  /** The workbook's changes could not be written into an .xlsx. */
  | "save-failed"
  /** Recalculation did not finish in time (a very large workbook). */
  | "calculation-timeout"
  /** Agent changes are waiting for review. */
  | "pending-changes"
  /** Opening another workbook would lose changes that aren't saved. */
  | "unsaved-changes"
  /** `openStored()` found no workbook with that id. */
  | "not-found"
  /** Called before the workbook has opened. */
  | "not-ready";

export class GridPathError extends Error {
  readonly code: GridPathErrorCode;
  /** What to do about it, in a sentence an app can show. */
  readonly hint: string;
  /** The underlying error, when there is one. */
  readonly cause?: unknown;
  constructor(code: GridPathErrorCode, message: string, hint: string, cause?: unknown) {
    super(message);
    this.name = "GridPathError";
    this.code = code;
    this.hint = hint;
    this.cause = cause;
  }
}

/** `save()` (or `open()`) was called while agent changes were pending review. */
export class PendingChangesError extends GridPathError {
  /** How many batches are pending. */
  readonly pending: number;
  constructor(pending: number) {
    super(
      "pending-changes",
      `${pending} agent change${pending === 1 ? " is" : "s are"} pending review: accept or reject ${pending === 1 ? "it" : "them"} before saving, or save with { includePending: true }`,
      "Accept or reject the AI's changes (sheet.ai.accept() / reject()), then try again.",
    );
    this.name = "PendingChangesError";
    this.pending = pending;
  }
}

/** `open()` would discard changes that could not be saved first. */
export class UnsavedChangesError extends GridPathError {
  constructor(name: string) {
    super(
      "unsaved-changes",
      `"${name}" has changes that aren't saved`,
      "Save or download it first, or open with { discardUnsaved: true } to drop the changes.",
    );
    this.name = "UnsavedChangesError";
  }
}

const isZip = (b: Uint8Array) => b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && (b[2] === 0x03 || b[2] === 0x05 || b[2] === 0x07);
const isOle = (b: Uint8Array) => b.length >= 8 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0;
const isUtf16 = (b: Uint8Array) => b.length >= 2 && ((b[0] === 0xff && b[1] === 0xfe) || (b[0] === 0xfe && b[1] === 0xff));
const isText = (b: Uint8Array) =>
  isUtf16(b) || b.subarray(0, 4096).every((c) => c === 9 || c === 10 || c === 13 || (c >= 32 && c < 127) || c >= 128);
/** Text that isn't a table: refused, not read as a column of lines. */
const NOT_A_TABLE = /\.(json|xml|html?|svg|md|markdown|ya?ml|rtf|pdf|js|mjs|ts|css)$/i;

/** Delimited text (CSV, TSV…): opened by converting it to a new .xlsx. */
export function isDelimitedText(bytes: Uint8Array, name: string): boolean {
  return bytes.length > 0 && !isZip(bytes) && !isOle(bytes) && isText(bytes) && !NOT_A_TABLE.test(name);
}

/** The name a workbook opened from delimited text saves under: sales.csv → sales.xlsx. */
export function workbookName(name: string, bytes: Uint8Array | null): string {
  return bytes && isDelimitedText(bytes, name) ? `${name.replace(/\.[^./\\]*$/, "")}.xlsx` : name;
}

/** A file that can't be opened (empty, an old .xls, a PDF…), told from its
 * first bytes before anything is opened; null when it looks like an .xlsx
 * or delimited text. */
export function fileProblem(bytes: Uint8Array, name: string): GridPathError | null {
  if (bytes.length === 0) {
    return new GridPathError("unsupported-file", `"${name}" is empty`, "Choose an .xlsx or CSV file.");
  }
  if (isOle(bytes)) {
    return new GridPathError(
      "unsupported-file",
      `"${name}" is an old-format .xls workbook or a password-protected .xlsx`,
      "Open it in Excel and save it as an .xlsx workbook without a password.",
    );
  }
  if (!isZip(bytes) && !isDelimitedText(bytes, name)) {
    return new GridPathError(
      "unsupported-file",
      isText(bytes) ? `"${name}" is a text file, not a spreadsheet` : `"${name}" is not a spreadsheet`,
      "Choose an .xlsx file, or a CSV or other delimited text file.",
    );
  }
  return null;
}

/** Delimited text that could not be read as a table. */
export function textTableError(name: string, cause: unknown): GridPathError {
  const reason = cleanReason(cause instanceof Error ? cause.message : String(cause));
  return new GridPathError(
    "malformed-file",
    `"${name}" could not be read as CSV (${reason})`,
    "Check that it is comma-, tab-, semicolon- or pipe-separated text.",
    cause,
  );
}

/** Why a file failed to open, from its first bytes and the engine's reason. */
export function openError(bytes: Uint8Array, name: string, reason: string, cause?: unknown): GridPathError {
  const plain = fileProblem(bytes, name);
  if (plain) return new GridPathError(plain.code, plain.message, plain.hint, cause);
  return new GridPathError(
    "malformed-file",
    `"${name}" could not be read as an .xlsx workbook (${cleanReason(reason)})`,
    "The file may be damaged or incomplete: open it in Excel, save it again, and retry.",
    cause,
  );
}

/** Recalculation was still running after `ms`. */
export function calcTimeout(ms: number): GridPathError {
  return new GridPathError(
    "calculation-timeout",
    `Recalculation is still running after ${ms < 1000 ? `${ms} ms` : `${Math.round(ms / 1000)} s`}`,
    "Large workbooks take a while after opening or a big change; try again once it finishes.",
  );
}

/** The session's changes could not be written into the file. */
export function saveFailed(reason: string, cause?: unknown): GridPathError {
  return new GridPathError(
    "save-failed",
    `The changes could not be written into the .xlsx (${reason})`,
    "Your edits are still on screen; undoing the last change usually lets the save go through.",
    cause,
  );
}

/** The engine's reason without its wrapping (`Error: Zip("…")` → `…`). */
export function cleanReason(reason: string): string {
  return reason
    .replace(/^(Error: )+/, "")
    .replace(/^open failed: (Error: )?/, "")
    .replace(/^[A-Za-z]+\("(.*)"\)$/s, "$1");
}

const SERVE_HINT =
  "Check that the app serves gridpath's engine.worker.js and .wasm files: with Vite, add optimizeDeps: { exclude: [\"gridpath\"] }; elsewhere pass engine: { workerUrl, wasmUrl }.";

/** The engine did not start: the worker or the .wasm did not load. */
export function engineStartError(cause: unknown): GridPathError {
  const reason = (cause instanceof Error ? cause.message : String(cause)).split(". ")[0];
  return new GridPathError("engine-failed", `The calculation engine could not start (${reason})`, SERVE_HINT, cause);
}

/** The engine stopped after it started. */
export function engineError(reason: string, cause?: unknown): GridPathError {
  return new GridPathError(
    "engine-failed",
    `The calculation engine stopped: ${cleanReason(reason)}`,
    "Values on screen may be out of date. Download a copy to keep your work, then reload the page.",
    cause,
  );
}

export function notReady(unmounted = false): GridPathError {
  return unmounted
    ? new GridPathError("not-ready", "This spreadsheet was unmounted", "Use the object a newly mounted spreadsheet passes to onReady.")
    : new GridPathError("not-ready", "The workbook is still opening", "Wait for onReady (or the promise from createSpreadsheet()) before calling this.");
}

/**
 * The store failed. `kept`: the edits are in this browser's drafts;
 * `retrying`: autosave tries again.
 */
export function storageError(action: string, cause: unknown, opts: { kept?: boolean; retrying?: boolean } = {}): GridPathError {
  const status = (cause as { status?: number } | null)?.status;
  const first =
    status === 401 || status === 403
      ? "The storage service refused the request: sign in again."
      : status && status < 300
        ? "Point the store's url at your workbook API."
        : status && status < 500
          ? "The storage service rejected the request."
          : status
            ? "The storage service had a problem."
            : "Check the connection.";
  const kept = opts.kept ? " Your changes are kept in this browser" : "";
  const retry = opts.retrying ? (kept ? " and saving will retry." : " Saving will retry.") : kept ? "." : "";
  const detail = cause instanceof Error ? cause.message : String(cause);
  return new GridPathError("storage-failed", `Could not ${action}: ${detail}`, `${first}${kept}${retry}`, cause);
}
