/**
 * A workspace: several workbooks kept live together, and the links between
 * them.
 *
 * Every workbook is its own engine model, opened on first use. When a
 * formula written into one reads another (`'[NOW.xlsx]Model'!B23`), the
 * reader gets a hidden copy of that sheet's values under the link's name,
 * and from then on:
 *
 *  - a cell edit in the source carries the cells it changed into the copy
 *    and recalculates the reader, and whatever reads the reader;
 *  - a row, column or sheet edit in the source is repeated on the copy, so
 *    the reader's links keep pointing at the cells they meant. The repeat is
 *    recorded in the reader as a batch of its own (`link_origin`), which is
 *    how a reopened session gets back to the same state, and it is dropped
 *    when the source batch is rejected.
 *
 * Pending batches are stored next to each workbook as before, and a
 * workbook's file is written by its own session. What the workspace adds to
 * a save is the links:
 *
 *  - a reader's file names each workbook it reads by a number (`[1]Model!B23`)
 *    and keeps that file's path and the last values read from it, as Excel
 *    does. `savePlan` says which workbook gets which number;
 *  - a file opened here has those numbered links attached again, to the
 *    workbooks they name, when those are found (`attachSaved`). One that is
 *    not found keeps the values the file was saved with;
 *  - a reader is not saved while a workbook it reads has unsaved row, column
 *    or sheet edits on a sheet it reads: its links would point at rows the
 *    other file does not have yet. The other one is saved first.
 *
 * Limits: rows moved in a source by another program (Excel) are not
 * followed; a defined name cannot point into another workbook; a copy saved
 * into another folder keeps the paths of the links its file already had.
 */
import fs from "node:fs";
import path from "node:path";
import type { ChangeBatch, UniverMutation } from "../core/types";
import { findLinks, isLinkSheetName, linkSheetName, linkText, quotedPrefix, rewriteLinks, splitLinkSheetName, type FoundLink } from "./links";
import { WorkbookSession, type EngineDelta, type LinkSave } from "./session";
import { baseStamp, loadReaders, loadStored, saveReaders, saveStored } from "./store";

/** One hidden copy held by a workbook. */
type LinkEntry = {
  /** The copy's name, and the link as formulas spell it: `[NOW.xlsx]Model`. */
  name: string;
  /** Absolute path of the workbook it copies. */
  source: string;
  /** The sheet there, as the name spells it. */
  sheet: string;
  /** `created_at` of the batch that first read it: later source edits are the ones to follow. */
  since: string;
};
/** A workbook's copies, by upper-cased name. */
type Entries = Map<string, LinkEntry>;

/** One numbered link of a workbook's file, and the workbook it was found to mean (null: not found, or not a workbook). */
type FileLink = { index: number; file: string | null };

/** What `savePlan` answers: how to write the links, or what has to be saved first. */
export type SavePlan =
  | { ok: true; links: LinkSave | null; reads: string[]; notes: string[] }
  | { ok: false; save_first: string[]; reason: string };

export type LinkStatus =
  | "linked"
  | "same_workbook"
  | "workbook_not_found"
  | "several_workbooks"
  | "sheet_not_found"
  | "numbered"
  | "circular"
  | "not_allowed"
  | "cannot_open"
  | "unsupported";

/** What became of one link a write contained. */
export type LinkReport = {
  /** The link as it was written. */
  link: string;
  status: LinkStatus;
  /** The file it means, when one was found. */
  workbook?: string;
  sheet?: string;
  detail?: string;
  /** The source's sheets, when the sheet was not among them. */
  sheets?: string[];
  /** The files it could mean, when there are several. */
  candidates?: string[];
};

export type ApplyResult = {
  batch: ChangeBatch;
  links: LinkReport[];
  /** Workbooks that read this one and were recalculated. */
  updated: string[];
};

/** A write was refused because links in it could not be followed. Nothing was applied. */
export class LinkError extends Error {
  constructor(readonly links: LinkReport[]) {
    const failed = links.filter(linkFailed);
    super(
      `Nothing was written: ${failed.length === 1 ? "a link" : `${failed.length} links`} to another workbook could not be followed. ` +
        failed.map((r) => `${r.link}: ${r.detail ?? r.status}${r.sheets ? ` (its sheets: ${r.sheets.join(", ")})` : ""}`).join("; "),
    );
    this.name = "LinkError";
  }
}

/** True for a link that would only ever show #REF!. A numbered link is how saved files spell theirs, and is left alone. */
export const linkFailed = (r: LinkReport): boolean => r.status !== "linked" && r.status !== "same_workbook" && r.status !== "numbered";

type PrepareOptions = {
  /** `created_at` of the batch being applied. */
  at: string;
  /** Replay: the links the batch was recorded with, and the batches stored after it. */
  recorded?: ChangeBatch["links"];
  later?: ChangeBatch[];
  /** False: only work out what each link means; attach nothing. */
  attach?: boolean;
};

/** What a batch that repeats another workbook's edit is labelled with. */
export const MIRROR_PROMPT = "linked workbook";

/** True for a batch nobody wrote: the repeat of another workbook's row, column or sheet edit. */
export const isMirrorBatch = (b: ChangeBatch): boolean => !!b.link_origin;

const eq = (a: string, b: string) => a.toUpperCase() === b.toUpperCase();
const originKey = (id: string, createdAt: string) => `${id}\u0000${createdAt}`;
const messageOf = (e: unknown) => (e instanceof Error ? e.message : String(e));

/** Characters a link's path cannot hold as they are. */
const encodeSegment = (seg: string) => seg.replace(/[%\s#?"<>|\\^`{}]/g, (ch) => `%${ch.charCodeAt(0).toString(16).toUpperCase().padStart(2, "0")}`);

/**
 * How a file in `dir` names the workbook `source` it reads: relative when
 * `source` is in that folder or below it (`NOW.xlsx`, `models/NOW%202026.xlsx`),
 * else the whole path (`file:///Users/me/NOW.xlsx`).
 */
export function linkTarget(source: string, dir: string): string {
  const rel = path.relative(dir, source);
  if (rel && !path.isAbsolute(rel) && rel.split(path.sep)[0] !== "..") return rel.split(path.sep).map(encodeSegment).join("/");
  const parts = source.split(/[\\/]/).filter((seg) => seg !== "");
  return `file:///${parts.map((seg, i) => (i === 0 && /^[A-Za-z]:$/.test(seg) ? seg : encodeSegment(seg))).join("/")}`;
}

/** The link sheet a repeated edit is about. */
function mirrorTarget(m: UniverMutation): string | null {
  if (m.type === "rename_sheet") return m.old_name;
  if (m.type === "delete_sheet") return m.name;
  if (m.type === "insert_rows" || m.type === "delete_rows" || m.type === "insert_columns" || m.type === "delete_columns") return m.sheet;
  return null;
}

/** Every sheet name a mutation writes to or names. */
function sheetsNamedBy(m: UniverMutation): string[] {
  const any = m as Record<string, any>;
  const names = [any.address?.sheet, any.sheet, any.old_name, any.new_name];
  if (m.type === "create_sheet" || m.type === "delete_sheet") names.push(any.name);
  return names.filter((n): n is string => typeof n === "string");
}

export class Workspace {
  private readonly sessions = new Map<string, WorkbookSession>();
  /** Reader path → its copies. */
  private readonly links = new Map<string, Entries>();
  /** Workbooks being opened or rebuilt right now. */
  private readonly opening = new Set<string>();
  /** Source path → the readers listed next to it on disk, once looked at. */
  private readonly listed = new Map<string, Set<string>>();
  /** Reader path → the numbered links of its file. */
  private readonly fileLinks = new Map<string, FileLink[]>();
  private seq = 0;
  private lastTick = 0;

  /** `allowed` says whether a file may be opened (the server's allowed folders). */
  constructor(private readonly allowed: (abs: string) => boolean = () => true) {}

  // -- members --

  /** The live session of a workbook; opened, with its stored batches replayed, on first use. */
  session(abs: string): WorkbookSession {
    const have = this.sessions.get(abs);
    if (have) return have;
    if (this.opening.has(abs)) throw new Error(`${path.basename(abs)} is already being opened (workbooks that read each other in a circle)`);
    this.opening.add(abs);
    try {
      const s = WorkbookSession.open(abs, fs.readFileSync(abs));
      const entries: Entries = new Map();
      const attached = this.attachSaved(s, entries);
      const stored = loadStored(abs);
      if (stored || attached) {
        const batches = stored?.batches ?? [];
        this.noteIds(batches);
        this.replay(s, batches, entries);
        // Repeats added for source edits this workbook missed, or dropped for rejected ones.
        if (s.batches.length !== batches.length || s.batches.some((b, i) => b.id !== batches[i].id)) saveStored(abs, s.batches);
      }
      this.sessions.set(abs, s);
      this.links.set(abs, entries);
      return s;
    } finally {
      this.opening.delete(abs);
    }
  }

  isOpen(abs: string): boolean {
    return this.sessions.has(abs);
  }

  open(): WorkbookSession[] {
    return [...this.sessions.values()];
  }

  /** Forget a workbook's session (after a save: the next use reopens it from the file). */
  close(abs: string): void {
    this.sessions.delete(abs);
    this.links.delete(abs);
    this.fileLinks.delete(abs);
  }

  /**
   * Before `abs` is saved: the workbooks that read it but are not open take
   * up its pending row, column and sheet edits now. Once those are in the
   * file there is no record of them left to follow.
   */
  beforeSave(abs: string): void {
    const s = this.sessions.get(abs);
    if (!s || !s.batches.some((b) => !b.link_origin && b.mutations.some((m) => mirrorTarget(m) !== null))) return;
    const listed = loadReaders(abs);
    const still = listed.filter((reader) => {
      if (!fs.existsSync(reader)) return false;
      if (!this.allowed(reader)) return true;
      try {
        // Opening it replays its links and repeats the edits it has not seen.
        this.session(reader);
        return this.followersOf(abs).includes(reader);
      } catch {
        return true;
      }
    });
    if (still.length !== listed.length) {
      this.listed.set(abs, new Set(still));
      try {
        saveReaders(abs, still);
      } catch {
        // a folder that cannot be written: the list is only a help
      }
    }
  }

  /** Remember, next to `source`, that `reader` reads it. */
  private noteReader(source: string, reader: string): void {
    let known = this.listed.get(source);
    if (!known) this.listed.set(source, (known = new Set(loadReaders(source))));
    if (known.has(reader)) return;
    known.add(reader);
    try {
      saveReaders(source, [...known]);
    } catch {
      // a folder that cannot be written: the list is only a help
    }
  }

  /** The workbooks `abs` reads, and which of their sheets. */
  linksOf(abs: string): Array<{ workbook: string; sheets: string[] }> {
    const bySource = new Map<string, string[]>();
    for (const e of this.links.get(abs)?.values() ?? []) {
      const list = bySource.get(e.source) ?? [];
      list.push(e.sheet);
      bySource.set(e.source, list);
    }
    return [...bySource].map(([workbook, sheets]) => ({ workbook, sheets: sheets.sort() }));
  }

  /** The open workbooks that read `abs`. */
  readersOf(abs: string): string[] {
    const out: string[] = [];
    for (const [reader, entries] of this.links) {
      if (this.sessions.has(reader) && [...entries.values()].some((e) => e.source === abs)) out.push(reader);
    }
    return out;
  }

  /**
   * The open workbooks whose state depends on `abs`: the ones that read it,
   * and the ones with a pending link to it that shows #REF! for now (its
   * sheet was deleted there), which a reject in `abs` can bring back.
   */
  followersOf(abs: string): string[] {
    const out = new Set(this.readersOf(abs));
    for (const [p, s] of this.sessions) {
      if (p !== abs && s.batches.some((b) => b.links?.some((l) => l.source === abs))) out.add(p);
    }
    return [...out];
  }

  /** Where the hidden copy `name` of workbook `abs` comes from. */
  sourceOf(abs: string, name: string): { workbook: string; sheet: string } | null {
    const e = this.links.get(abs)?.get(name.toUpperCase());
    return e ? { workbook: e.source, sheet: e.sheet } : null;
  }

  // -- links in the file --

  /**
   * A file that was saved reading other workbooks spells them `[1]Model`.
   * Attach each one whose workbook is found: the copy is taken from that
   * workbook as it is now, and the formulas get its file name
   * (`[NOW.xlsx]Model`), the spelling used everywhere else here. A link whose
   * workbook is not found, or cannot be opened, keeps the values the file
   * was saved with. True when anything was attached.
   */
  private attachSaved(s: WorkbookSession, entries: Entries): boolean {
    const saved = s.savedLinks() ?? [];
    const taken = new Set<string>();
    const found: FileLink[] = saved.map((l) => {
      const file = l.target === null ? null : this.findSaved(s.path, l.target);
      // Two numbers for one file: the first is the one used.
      if (file === null || taken.has(file.toUpperCase())) return { index: l.index, file: null };
      taken.add(file.toUpperCase());
      return { index: l.index, file };
    });
    this.fileLinks.set(s.path, found);
    if (!found.some((f) => f.file) || !s.supportsLinks()) return false;
    let attached = false;
    for (const name of s.unresolvedLinks()) {
      const parts = splitLinkSheetName(name);
      if (!parts || !/^\d+$/.test(parts.book)) continue;
      const file = found.find((f) => f.index === Number(parts.book))?.file;
      if (!file) continue;
      const book = path.basename(file);
      if (/[[\]]/.test(book) || this.opening.has(file)) continue;
      let src: WorkbookSession;
      try {
        src = this.session(file);
      } catch {
        continue;
      }
      if (this.reads(file, s.path)) continue;
      const from = src.sheetNames().find((n) => eq(n, parts.sheet));
      if (!from) continue;
      let linked = linkSheetName(book, from);
      if (entries.has(linked.toUpperCase())) continue;
      try {
        s.syncLink(name, src, from);
      } catch {
        continue;
      }
      try {
        s.renameLink(name, linked);
      } catch {
        // It stays under its number, and is followed all the same.
        linked = name;
      }
      // "": every edit still pending in the source came after the file was saved.
      entries.set(linked.toUpperCase(), { name: linked, source: file, sheet: from, since: "" });
      this.noteReader(file, s.path);
      attached = true;
    }
    return attached;
  }

  /**
   * The file a saved link's path means: where the path says (relative to
   * the reader's folder), else a file of that name next to the reader, as
   * for a model that came from another machine. Only a workbook the server
   * may open: the path is whatever the file says.
   */
  private findSaved(reader: string, target: string): string | null {
    let t = target;
    try {
      t = decodeURIComponent(t);
    } catch {
      // as written
    }
    t = t.replace(/^file:\/\/(localhost)?/i, "");
    if (/^\/[A-Za-z]:[\\/]/.test(t)) t = t.slice(1);
    if (!/\.xls[xm]$/i.test(t)) return null;
    const here = path.dirname(reader);
    const base = t.split(/[\\/]/).pop() ?? "";
    for (const candidate of [path.resolve(here, t), path.join(here, base)]) {
      let names: string[];
      try {
        names = fs.readdirSync(path.dirname(candidate));
      } catch {
        continue;
      }
      // The file's own spelling of its name, and one session per file.
      const wanted = path.basename(candidate);
      const real = names.find((n) => n === wanted) ?? names.find((n) => eq(n, wanted));
      if (!real) continue;
      const onDisk = path.join(path.dirname(candidate), real);
      const abs = [...this.sessions.keys()].find((p) => eq(p, onDisk)) ?? onDisk;
      if (eq(abs, reader) || !this.allowed(abs)) continue;
      try {
        if (!fs.statSync(abs).isFile()) continue;
      } catch {
        continue;
      }
      return abs;
    }
    return null;
  }

  /**
   * What a save of `abs` into `target` (its own file, or a copy somewhere)
   * has to know about the workbooks it reads; or the workbooks that have to
   * be saved before it.
   */
  savePlan(abs: string, target: string = abs): SavePlan {
    // A copy nothing reads any more (its formula was replaced) is not part of the file.
    const read = this.sessions.get(abs)?.copiesRead();
    const entries = [...(this.links.get(abs)?.values() ?? [])].filter((e) => !read || read.has(e.name.toUpperCase()));
    if (entries.length === 0) return { ok: true, links: null, reads: [], notes: [] };
    const names = (paths: string[]) => paths.map((p) => path.basename(p)).join(", ");
    const first = this.unsavedMoves(entries);
    if (first.length > 0) {
      return {
        ok: false,
        save_first: first,
        reason:
          `${path.basename(abs)} reads ${names(first)}, where rows, columns or sheets were changed and not saved yet. ` +
          `Saved now, its links would point at cells the other file does not have. Save ${names(first)} first, then this workbook. Nothing was saved.`,
      };
    }
    const inFile = this.fileLinks.get(abs) ?? [];
    const copies: LinkSave["copies"] = new Map();
    for (const e of entries) {
      const index = inFile.find((f) => f.file === e.source)?.index;
      copies.set(e.name.toUpperCase(), { source: e.source, ...(index !== undefined ? { index } : {}) });
    }
    const sources = [...new Set(entries.map((e) => e.source))];
    const unsaved = sources.filter((p) => this.sessions.get(p)?.batches.some((b) => !b.link_origin));
    const dir = path.dirname(target);
    return {
      ok: true,
      links: { copies, target: (source) => linkTarget(source, dir) },
      reads: sources,
      notes: unsaved.length > 0 ? [`The figures read from ${names(unsaved)} include changes there that are not saved yet.`] : [],
    };
  }

  /**
   * The workbooks behind `entries` that have an unsaved row, column or sheet
   * edit (a new sheet too) on a sheet that is read. Whenever the link was
   * made: the other file is still laid out as it was before the edit.
   */
  private unsavedMoves(entries: LinkEntry[]): string[] {
    const bySource = new Map<string, Set<string>>();
    for (const e of entries) {
      let sheets = bySource.get(e.source);
      if (!sheets) bySource.set(e.source, (sheets = new Set()));
      sheets.add(e.sheet.toUpperCase());
    }
    const out: string[] = [];
    for (const [source, sheets] of bySource) {
      // Entries carry the sheets' names as they are now: a rename shows up as its new name.
      const moved = this.sessions.get(source)?.batches.some(
        (b) =>
          !b.link_origin &&
          b.mutations.some((m) => {
            const sheet = m.type === "rename_sheet" ? m.new_name : m.type === "create_sheet" ? m.name : mirrorTarget(m);
            return sheet !== null && sheets.has(sheet.toUpperCase());
          }),
      );
      if (moved) out.push(source);
    }
    return out;
  }

  /**
   * The open workbooks whose files link to `abs` and that hold repeats of
   * its row, column or sheet edits: once `abs` is saved, their files point
   * at the old rows until they are saved too.
   */
  movedWith(abs: string): string[] {
    const out: string[] = [];
    for (const [reader, s] of this.sessions) {
      if (reader === abs || !s.batches.some((b) => b.link_origin?.path === abs)) continue;
      if (this.fileLinks.get(reader)?.some((f) => f.file === abs)) out.push(reader);
    }
    return out;
  }

  // -- writes --

  /**
   * Apply a tool's mutations to `s` as one batch: follow the links its
   * formulas contain, evaluate, store the batch, and bring the workbooks
   * that read `s` up to date.
   */
  apply(s: WorkbookSession, tool: string, mutations: UniverMutation[]): ApplyResult {
    this.refuseLinkSheetWrites(s, mutations);
    const at = this.stamp();
    const entries = this.entriesOf(s.path);
    // All or nothing: a write with a link that leads nowhere is refused whole.
    const check = this.prepareLinks(s, entries, mutations, { at, attach: false });
    if (check.reports.some(linkFailed)) throw new LinkError(check.reports);
    const before = new Set(entries.keys());
    const prep = this.prepareLinks(s, entries, mutations, { at });
    let applied: ReturnType<WorkbookSession["applyTracked"]>;
    try {
      applied = s.applyTracked({
        id: this.nextId(),
        prompt: tool,
        justification: tool,
        mutations: prep.mutations,
        created_at: at,
        ...(prep.links.length ? { links: prep.links } : {}),
      });
    } catch (e) {
      // The write did not land: neither do the copies taken for it.
      for (const [key, entry] of [...entries]) {
        if (before.has(key)) continue;
        entries.delete(key);
        try {
          s.detachLink(entry.name);
        } catch {
          // nothing reads it
        }
      }
      throw e;
    }
    const { batch, delta } = applied;
    saveStored(s.path, s.batches);
    const updated: string[] = [];
    this.propagate(s, batch, delta ?? { full: true }, new Set([s.path]), updated);
    return { batch, links: prep.reports, updated: [...new Set(updated)] };
  }

  /**
   * Rebuild a workbook from its file with `batches` (what is left after a
   * reject), then every open workbook that reads it.
   */
  rebuild(abs: string, batches: ChangeBatch[], above: Set<string> = new Set()): WorkbookSession {
    const old = this.sessions.get(abs);
    if (!old) throw new Error(`not open: ${abs}`);
    const fresh = WorkbookSession.open(abs, old.originalBytes);
    const entries: Entries = new Map();
    this.opening.add(abs);
    try {
      this.attachSaved(fresh, entries);
      this.replay(fresh, batches, entries);
    } finally {
      this.opening.delete(abs);
    }
    this.sessions.set(abs, fresh);
    this.links.set(abs, entries);
    saveStored(abs, fresh.batches);
    const path_ = new Set([...above, abs]);
    for (const reader of this.followersOf(abs)) {
      if (path_.has(reader)) continue;
      const d = this.sessions.get(reader);
      if (d) this.rebuild(reader, d.batches, path_);
    }
    return fresh;
  }

  // -- replay --

  /**
   * Bring a freshly opened session to the state its stored batches describe.
   * Copies of other workbooks are taken from those workbooks as they are
   * now; the stored repeats of their row, column and sheet edits then move
   * this workbook's links exactly as they moved when they happened.
   */
  private replay(s: WorkbookSession, stored: ChangeBatch[], entries: Entries): void {
    const exists = (m: UniverMutation) => {
      const target = mirrorTarget(m);
      return target !== null && s.linkSheetNames().some((n) => eq(n, target));
    };
    for (let i = 0; i < stored.length; i++) {
      const b = stored[i];
      if (b.link_origin) {
        if (this.originRejected(b.link_origin)) continue;
        // A repeat whose copy is not there (its link was rejected, the file
        // is missing) is kept as a record and not applied.
        s.apply(b, { evaluate: false, only: exists });
        this.noteLinkSheetOps(entries, b.mutations);
        continue;
      }
      const prep = this.prepareLinks(s, entries, b.mutations, { at: b.created_at, recorded: b.links, later: stored.slice(i + 1) });
      // A link that could not be followed this time (its file is away) keeps
      // its record; one whose formula is gone from the batch (the cell was
      // rejected) does not.
      const away = new Set(prep.reports.filter(linkFailed).map((r) => r.link.toUpperCase()));
      const kept = (b.links ?? []).filter((e) => away.has(e.name.toUpperCase()));
      const links = [...new Map([...kept, ...prep.links].map((e) => [e.name.toUpperCase(), e])).values()];
      const { links: _recorded, ...rest } = b;
      s.apply({ ...rest, mutations: prep.mutations, ...(links.length ? { links } : {}) });
    }
    this.catchUp(s, entries);
    this.syncAll(s, entries);
    s.settle();
    // Repeats of edits to a workbook this one no longer reads are of no use.
    const read = new Set([...entries.values()].map((e) => e.source));
    for (const b of s.batches) for (const l of b.links ?? []) read.add(l.source);
    for (let i = s.batches.length - 1; i >= 0; i--) {
      const o = s.batches[i].link_origin;
      if (o && !read.has(o.path)) s.batches.splice(i, 1);
    }
  }

  /** Repeat the source edits made while this workbook was not open. */
  private catchUp(s: WorkbookSession, entries: Entries): void {
    for (const srcPath of new Set([...entries.values()].map((e) => e.source))) {
      const src = this.sessions.get(srcPath);
      if (!src) continue;
      const have = new Set(
        s.batches.filter((b) => b.link_origin?.path === srcPath).map((b) => originKey(b.link_origin!.batch_id, b.link_origin!.created_at)),
      );
      for (const b of src.batches) {
        if (b.link_origin || have.has(originKey(b.id, b.created_at))) continue;
        this.applyMirror(s, entries, src, b);
      }
    }
  }

  /** True when the source batch a repeat came from was rejected (as opposed to saved into its file). */
  private originRejected(o: NonNullable<ChangeBatch["link_origin"]>): boolean {
    const src = this.sessions.get(o.path);
    if (!src) return false;
    if (src.batches.some((b) => b.id === o.batch_id && b.created_at === o.created_at)) return false;
    try {
      const now = baseStamp(o.path);
      return now.size === o.base.size && Math.abs(now.mtimeMs - o.base.mtimeMs) <= 1;
    } catch {
      return false;
    }
  }

  // -- links in formulas --

  /**
   * Find the links in the formulas a batch writes, attach a copy of each
   * sheet they read, and respell them the one way the engine and the copy's
   * name agree on (`'[NOW.xlsx]Model'!`). A link that cannot be followed is
   * left as written (it evaluates to #REF!) and reported.
   */
  private prepareLinks(
    s: WorkbookSession,
    entries: Entries,
    mutations: UniverMutation[],
    opts: PrepareOptions,
  ): { mutations: UniverMutation[]; links: NonNullable<ChangeBatch["links"]>; reports: LinkReport[] } {
    const seen = new Map<string, { prefix: string | null; report: LinkReport }>();
    const used = new Map<string, { name: string; source: string; sheet: string }>();
    const one = (l: FoundLink): string | null => {
      const key = `${l.dir}\u0000${l.book}\u0000${l.sheet}`.toUpperCase();
      let r = seen.get(key);
      if (!r) {
        r = this.resolveLink(s, entries, l, opts, used);
        seen.set(key, r);
      }
      return r.prefix;
    };
    const fix = (text: string) => (text.includes("[") ? rewriteLinks(text, one) : text);
    const out = mutations.map((m): UniverMutation => {
      if (m.type === "set_cell") {
        let next = m;
        if (typeof m.new_formula === "string") {
          const f = fix(m.new_formula);
          if (f !== m.new_formula) next = { ...next, new_formula: f };
        }
        if (typeof m.new_value === "string" && m.new_value.startsWith("=")) {
          const v = fix(m.new_value);
          if (v !== m.new_value) next = { ...next, new_value: v };
        }
        return next;
      }
      if (m.type === "set_range") {
        let changed = false;
        const values = m.values.map((row) =>
          row.map((v) => {
            if (typeof v !== "string" || !v.startsWith("=")) return v;
            const w = fix(v);
            if (w !== v) changed = true;
            return w;
          }),
        );
        return changed ? { ...m, values } : m;
      }
      if (m.type === "define_name" && findLinks(m.ref).length > 0) {
        throw new Error("a defined name cannot point into another workbook yet. Write the link in a cell and name that cell.");
      }
      return m;
    });
    return { mutations: out, links: [...used.values()], reports: [...seen.values()].map((r) => r.report) };
  }

  private resolveLink(
    s: WorkbookSession,
    entries: Entries,
    l: FoundLink,
    opts: PrepareOptions,
    used: Map<string, { name: string; source: string; sheet: string }>,
  ): { prefix: string | null; report: LinkReport } {
    const link = linkText(l);
    const fail = (status: LinkStatus, extra: Partial<LinkReport> = {}) => ({ prefix: null, report: { link, status, ...extra } });

    // `[1]Model`: the workbook this file's link number 1 names.
    let numbered: string | undefined;
    if (/^\d+$/.test(l.book)) {
      const n = Number(l.book);
      numbered = this.fileLinks.get(s.path)?.find((f) => f.index === n)?.file ?? undefined;
      if (!numbered) {
        const count = s.savedLinks()?.length;
        if (count !== undefined && n > count) {
          return fail("workbook_not_found", {
            detail: `this file has no link [${n}] (${count === 0 ? "it was saved reading no other workbook" : `it was saved reading ${count}`}). Write the other file's name: '[File.xlsx]Sheet'!A1.`,
          });
        }
        return fail("numbered", {
          detail:
            "[1], [2]… is how a saved file refers to the workbooks it links to, and the workbook this one means was not found, so the link keeps the value the file was saved with. To read a workbook live, write its file name: '[File.xlsx]Sheet'!A1.",
        });
      }
      l = { ...l, dir: "", book: path.basename(numbered) };
    }
    if (!s.supportsLinks()) return fail("unsupported", { detail: "this engine build cannot link workbooks (run `npm run build:engine`)" });

    // Which file. A link spelled like one this workbook already has means the same file.
    const spelled = linkSheetName(l.book, l.sheet);
    const known = numbered ?? (l.dir ? undefined : (entries.get(spelled.toUpperCase())?.source ?? opts.recorded?.find((e) => eq(e.name, spelled))?.source));
    let file: string;
    if (known) file = known;
    else {
      const where = this.locate(s.path, l);
      if ("status" in where) return fail(where.status, where.extra);
      file = where.path;
    }

    if (eq(file, s.path)) {
      const own = s.sheetNames().find((n) => eq(n, l.sheet));
      if (!own) return fail("sheet_not_found", { workbook: file, sheets: s.sheetNames(), detail: `the link points at this workbook itself, which has no sheet named ${l.sheet}` });
      return {
        prefix: quotedPrefix(own),
        report: { link, status: "same_workbook", workbook: file, sheet: own, detail: "the link points at this workbook, so it was written as an ordinary reference" },
      };
    }

    const book = path.basename(file);
    if (/[[\]]/.test(book)) return fail("unsupported", { workbook: file, detail: "a file whose name has [ or ] in it cannot be linked; rename the file" });
    const circular = () => fail("circular", { workbook: file, detail: `${book} reads this workbook (directly or through another one), so a link back would be circular` });
    if (this.opening.has(file)) return circular();
    let src: WorkbookSession;
    try {
      src = this.session(file);
    } catch (e) {
      return fail("cannot_open", { workbook: file, detail: messageOf(e) });
    }
    if (this.reads(file, s.path)) return circular();

    // Which sheet. On a replay it may have been renamed since the batch was written.
    let sheet = src.sheetNames().find((n) => eq(n, l.sheet));
    let from = sheet;
    if (!sheet && opts.later) {
      from = this.followRenames(src, book, l.sheet, opts.later, opts.at) ?? undefined;
      if (from) sheet = l.sheet;
    }
    if (!sheet || !from) return fail("sheet_not_found", { workbook: file, sheets: src.sheetNames(), detail: `${book} has no sheet named ${l.sheet}` });

    const name = linkSheetName(book, sheet);
    const key = name.toUpperCase();
    const have = entries.get(key);
    if (have && have.source !== file) {
      return fail("several_workbooks", {
        workbook: file,
        candidates: [have.source, file],
        detail: `this workbook already reads another ${book} (${have.source}); two files with the same name cannot both be linked`,
      });
    }
    if (!have && opts.attach !== false) {
      try {
        s.syncLink(name, src, from);
      } catch (e) {
        return fail("cannot_open", { workbook: file, detail: messageOf(e) });
      }
      entries.set(key, { name, source: file, sheet, since: opts.at });
      this.noteReader(file, s.path);
    }
    used.set(key, { name, source: file, sheet });
    return { prefix: quotedPrefix(name), report: { link, status: "linked", workbook: file, sheet } };
  }

  /**
   * The file a link's book name means: the folder it was written with, else
   * the reading workbook's own folder, else a workbook already open, else
   * the folders of the open workbooks.
   */
  private locate(reader: string, l: FoundLink): { path: string } | { status: LinkStatus; extra: Partial<LinkReport> } {
    const names = /\.xls[xm]$/i.test(l.book) ? [l.book] : [l.book, `${l.book}.xlsx`, `${l.book}.xlsm`];
    const inFolder = (dir: string): string[] => {
      let files: string[];
      try {
        files = fs.readdirSync(dir);
      } catch {
        return [];
      }
      // In the order of preference, with the file's own spelling of its name.
      return names.flatMap((n) => files.filter((f) => eq(f, n))).map((f) => path.join(dir, f));
    };
    const here = path.dirname(reader);
    let found: string | undefined;
    if (l.dir) {
      const home = process.env.HOME ?? process.env.USERPROFILE ?? "";
      const dir = path.resolve(here, l.dir.startsWith("~/") || l.dir.startsWith("~\\") ? path.join(home, l.dir.slice(2)) : l.dir);
      found = inFolder(dir)[0];
      if (!found) return { status: "workbook_not_found", extra: { detail: `no ${l.book} in ${dir}` } };
    } else {
      found = inFolder(here)[0];
      if (!found) {
        const open = [...this.sessions.keys()].filter((p) => names.some((n) => eq(path.basename(p), n)));
        const folders = new Set([...this.sessions.keys()].map((p) => path.dirname(p)));
        folders.delete(here);
        const elsewhere = [...folders].flatMap(inFolder);
        const candidates = [...new Map([...open, ...elsewhere].map((p) => [p.toUpperCase(), p])).values()];
        if (candidates.length > 1) {
          return { status: "several_workbooks", extra: { candidates, detail: `several files are called ${l.book}; write the folder too: '/folder/[${l.book}]${l.sheet}'!A1` } };
        }
        found = candidates[0];
      }
      if (!found) {
        return { status: "workbook_not_found", extra: { detail: `no ${l.book} next to this workbook (${here}) or among the open workbooks; write the folder too: '/folder/[${l.book}]${l.sheet}'!A1` } };
      }
    }
    // One session per file, whatever the spelling it was first opened with.
    const loaded = [...this.sessions.keys()].find((p) => eq(p, found!));
    const abs = loaded ?? (eq(found, reader) ? reader : found);
    if (!this.allowed(abs)) return { status: "not_allowed", extra: { workbook: abs, detail: `${abs} is outside the folders this server may open` } };
    return { path: abs };
  }

  /**
   * Replay: the source sheet a link written as `[book]sheet` means now,
   * following the renames recorded after that batch and the ones this
   * workbook has not seen yet. Null when it is gone.
   */
  private followRenames(src: WorkbookSession, book: string, sheet: string, later: ChangeBatch[], since: string): string | null {
    let name = sheet;
    const seen = new Set<string>();
    for (const b of later) {
      const o = b.link_origin;
      if (!o || o.path !== src.path || this.originRejected(o)) continue;
      seen.add(originKey(o.batch_id, o.created_at));
      for (const m of b.mutations) {
        if (m.type !== "rename_sheet" || !eq(m.old_name, linkSheetName(book, name))) continue;
        name = splitLinkSheetName(m.new_name)?.sheet ?? name;
      }
    }
    for (const b of src.batches) {
      if (b.link_origin || b.created_at <= since || seen.has(originKey(b.id, b.created_at))) continue;
      for (const m of b.mutations) if (m.type === "rename_sheet" && eq(m.old_name, name)) name = m.new_name;
    }
    return src.sheetNames().find((n) => eq(n, name)) ?? null;
  }

  /** Whether workbook `a` reads workbook `b`, directly or through others. */
  private reads(a: string, b: string, seen: Set<string> = new Set()): boolean {
    if (seen.has(a)) return false;
    seen.add(a);
    for (const e of this.links.get(a)?.values() ?? []) {
      if (e.source === b || this.reads(e.source, b, seen)) return true;
    }
    return false;
  }

  private refuseLinkSheetWrites(s: WorkbookSession, mutations: UniverMutation[]): void {
    for (const m of mutations) {
      const name = sheetsNamedBy(m).find(isLinkSheetName);
      if (!name) continue;
      if ((m.type === "create_sheet" && m.name === name) || (m.type === "rename_sheet" && m.new_name === name && !isLinkSheetName(m.old_name))) {
        throw new Error(`a sheet name cannot start with "[": ${name}`);
      }
      const from = this.sourceOf(s.path, name);
      throw new Error(
        `${name} is not a sheet of this workbook: it is a sheet of another workbook that formulas here read. ` +
          (from ? `Edit it there (path = ${from.workbook}, sheet = ${from.sheet}).` : "Edit it in its own file."),
      );
    }
  }

  // -- keeping readers current --

  /**
   * After `src` changed: repeat a row, column or sheet edit on the copies
   * its readers hold, carry the changed values over, recalculate each reader
   * and go on to the readers' readers.
   */
  private propagate(src: WorkbookSession, batch: ChangeBatch | null, delta: EngineDelta, above: Set<string>, updated: string[]): void {
    for (const [reader, entries] of [...this.links]) {
      if (above.has(reader)) continue;
      const d = this.sessions.get(reader);
      if (!d || ![...entries.values()].some((e) => e.source === src.path)) continue;
      const mirrored = batch ? this.applyMirror(d, entries, src, batch) : false;
      const changed = mirrored || delta.full ? this.syncAll(d, entries, src.path) : this.pushChanged(d, entries, src, delta);
      if (!mirrored && changed === 0) continue;
      const next = d.settle();
      if (mirrored) saveStored(d.path, d.batches);
      updated.push(reader);
      this.propagate(d, null, next, new Set([...above, reader]), updated);
    }
  }

  /**
   * Repeat `batch`'s row, column and sheet edits on the copies `d` holds of
   * the sheets they were made on. Not evaluated here. True when there was
   * anything to repeat.
   */
  private applyMirror(d: WorkbookSession, entries: Entries, src: WorkbookSession, batch: ChangeBatch): boolean {
    const mutations = this.mirrorOf(entries, src.path, batch);
    if (mutations.length === 0) return false;
    let base = { size: 0, mtimeMs: 0 };
    try {
      base = baseStamp(src.path);
    } catch {
      // the file is gone: the repeat is simply never taken for a rejected one
    }
    try {
      d.apply(
        {
          id: this.nextId(),
          prompt: MIRROR_PROMPT,
          justification: `${path.basename(src.path)}: ${batch.prompt}`,
          mutations,
          created_at: this.stamp(),
          link_origin: { path: src.path, batch_id: batch.id, created_at: batch.created_at, base },
        },
        { evaluate: false, only: (m) => d.linkSheetNames().some((n) => eq(n, mirrorTarget(m) ?? "")) },
      );
      this.noteLinkSheetOps(entries, mutations);
    } catch {
      // The copy is brought back in line with its source right after this.
    }
    return true;
  }

  private mirrorOf(entries: Entries, srcPath: string, batch: ChangeBatch): UniverMutation[] {
    if (batch.link_origin) return [];
    // Source sheet (upper-cased) → the copy's name, as the edits so far leave it.
    const copies = new Map<string, string>();
    for (const e of entries.values()) if (e.source === srcPath && e.since < batch.created_at) copies.set(e.sheet.toUpperCase(), e.name);
    const out: UniverMutation[] = [];
    for (const m of batch.mutations) {
      if (m.type === "insert_rows" || m.type === "insert_columns") {
        const name = copies.get(m.sheet.toUpperCase());
        if (name) out.push({ type: m.type, sheet: name, before: m.before, count: m.count });
      } else if (m.type === "delete_rows" || m.type === "delete_columns") {
        const name = copies.get(m.sheet.toUpperCase());
        if (name) out.push({ type: m.type, sheet: name, start: m.start, count: m.count });
      } else if (m.type === "rename_sheet") {
        const name = copies.get(m.old_name.toUpperCase());
        const parts = name ? splitLinkSheetName(name) : null;
        if (!name || !parts) continue;
        const renamed = linkSheetName(parts.book, m.new_name);
        out.push({ type: "rename_sheet", old_name: name, new_name: renamed });
        copies.delete(m.old_name.toUpperCase());
        copies.set(m.new_name.toUpperCase(), renamed);
      } else if (m.type === "delete_sheet") {
        const name = copies.get(m.name.toUpperCase());
        if (!name) continue;
        out.push({ type: "delete_sheet", name });
        copies.delete(m.name.toUpperCase());
      }
    }
    return out;
  }

  /** Keep the record of a workbook's copies in step with renames and deletions of them. */
  private noteLinkSheetOps(entries: Entries, mutations: UniverMutation[]): void {
    for (const m of mutations) {
      if (m.type === "rename_sheet" && isLinkSheetName(m.old_name)) {
        const e = entries.get(m.old_name.toUpperCase());
        const parts = splitLinkSheetName(m.new_name);
        if (!e || !parts) continue;
        entries.delete(m.old_name.toUpperCase());
        entries.set(m.new_name.toUpperCase(), { ...e, name: m.new_name, sheet: parts.sheet });
      } else if (m.type === "delete_sheet" && isLinkSheetName(m.name)) {
        entries.delete(m.name.toUpperCase());
      }
    }
  }

  /**
   * Make `d`'s copies equal to their sources as they are now (only those of
   * `only`, when given). A copy whose sheet no longer exists is dropped, so
   * its links show #REF!. Returns how many cells or copies changed.
   */
  private syncAll(d: WorkbookSession, entries: Entries, only?: string): number {
    let changed = 0;
    for (const [key, e] of [...entries]) {
      if (only && e.source !== only) continue;
      const src = this.sessions.get(e.source);
      if (!src) continue;
      const from = src.sheetNames().find((n) => eq(n, e.sheet));
      try {
        if (from) changed += d.syncLink(e.name, src, from).changed;
        else {
          d.detachLink(e.name);
          entries.delete(key);
          changed++;
        }
      } catch {
        // the copy keeps what it has
      }
    }
    return changed;
  }

  /** Carry the cells an evaluation of `src` changed into `d`'s copies. Returns how many differed. */
  private pushChanged(d: WorkbookSession, entries: Entries, src: WorkbookSession, delta: Extract<EngineDelta, { full: false }>): number {
    const bySheet = new Map<number, number[]>();
    for (const c of delta.cells) {
      let list = bySheet.get(c.sheet);
      if (!list) bySheet.set(c.sheet, (list = []));
      list.push(c.r, c.c);
    }
    const names = src.allSheetNames();
    let changed = 0;
    for (const [index, cells] of bySheet) {
      const name = names[index];
      if (name === undefined || isLinkSheetName(name)) continue;
      for (const e of entries.values()) {
        if (e.source === src.path && eq(e.sheet, name)) changed += d.pushLink(e.name, src, index, Int32Array.from(cells));
      }
    }
    return changed;
  }

  // -- small things --

  private entriesOf(abs: string): Entries {
    let entries = this.links.get(abs);
    if (!entries) this.links.set(abs, (entries = new Map()));
    return entries;
  }

  private nextId(): string {
    return `b${++this.seq}`;
  }

  /** New batches are numbered and stamped after every stored one (whatever the clock says now). */
  private noteIds(batches: ChangeBatch[]): void {
    for (const b of batches) {
      const n = Number(b.id.replace(/^b/, ""));
      if (Number.isFinite(n) && n > this.seq) this.seq = n;
      const t = Date.parse(b.created_at);
      if (Number.isFinite(t) && t > this.lastTick) this.lastTick = t;
    }
  }

  /** A timestamp later than every one this workspace gave before. */
  private stamp(): string {
    let t = Date.now();
    if (t <= this.lastTick) t = this.lastTick + 1;
    this.lastTick = t;
    return new Date(t).toISOString();
  }
}
