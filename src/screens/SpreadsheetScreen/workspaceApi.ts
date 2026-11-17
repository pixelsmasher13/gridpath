/**
 * The workspace: the workbooks that are worked on together. Each is a live
 * model on the Tauri side once something needs it (a tab, or a formula in
 * another workbook that reads it), and a formula in one can read a cell of
 * another: `='[NOW.xlsx]Model'!B23`.
 *
 * There can be several workspaces, each a name and a list of files, saved
 * in the app's settings. One is shown at a time: its list is the one the
 * Tauri side holds, and adding or removing a model saves it.
 */
import { invoke } from "@tauri-apps/api/core";

export interface WorkspaceMember {
  path: string;
  /** File name. */
  name: string;
  /** `open` in a tab, `loaded` without one (another workbook reads it),
   * `idle` on disk, `missing`, or `failed` (it could not be loaded). */
  state: "open" | "loaded" | "idle" | "missing" | "failed";
  /** On the list. A workbook that is only loaded because an open one reads
   * it is reported too, with `false`. */
  member: boolean;
  /** Paths of the workbooks this one reads, and of the ones that read it. */
  reads: string[];
  readBy: string[];
  /** Links of this workbook whose sheet is gone from the workbook read. */
  broken: string[];
  /** Its formulas followed a change in a workbook it reads, and its file
   * does not have that yet. */
  needsSaving: boolean;
  error: string | null;
}

// ---------------------------------------------------------------------------
// Saved workspaces
// ---------------------------------------------------------------------------

/** A saved workspace: a name and the models on it (full paths). */
export interface SavedWorkspace {
  id: string;
  name: string;
  members: string[];
}

/** Every saved workspace, and the one that is shown. Never empty. */
export interface WorkspaceStore {
  active: string;
  list: SavedWorkspace[];
}

/** Settings row (the app's database) that holds the store as JSON. */
const STORE_SETTING = "ssws_workspaces";
/** Where the first build kept its single list. */
const LEGACY_KEY = "gridpath.workspaceMembers";

const pathList = (value: unknown): string[] => (Array.isArray(value) ? value.filter((p): p is string => typeof p === "string") : []);

/** A store with one workspace in it. */
export function singleWorkspaceStore(id: string, members: string[] = []): WorkspaceStore {
  return { active: id, list: [{ id, name: "Workspace 1", members }] };
}

/** The store a settings value holds; null when it holds none (or nothing usable). */
export function parseWorkspaceStore(text: string): WorkspaceStore | null {
  if (!text) return null;
  try {
    const raw = JSON.parse(text);
    const list: SavedWorkspace[] = (Array.isArray(raw?.list) ? raw.list : [])
      .filter((w: any) => w && typeof w.id === "string" && typeof w.name === "string")
      .map((w: any) => ({ id: w.id, name: w.name, members: pathList(w.members) }));
    if (list.length === 0) return null;
    return { active: list.some((w) => w.id === raw.active) ? raw.active : list[0].id, list };
  } catch {
    return null;
  }
}

export function activeWorkspace(store: WorkspaceStore): SavedWorkspace {
  return store.list.find((w) => w.id === store.active) ?? store.list[0];
}

/** `wanted`, or `wanted 2`, `wanted 3`… when another workspace has the name. */
export function uniqueWorkspaceName(store: WorkspaceStore, wanted: string, exceptId?: string): string {
  const taken = new Set(store.list.filter((w) => w.id !== exceptId).map((w) => w.name.toLowerCase()));
  if (!taken.has(wanted.toLowerCase())) return wanted;
  for (let n = 2; ; n++) if (!taken.has(`${wanted} ${n}`.toLowerCase())) return `${wanted} ${n}`;
}

/** The store with the shown workspace's models replaced; the same object when nothing changed. */
export function withActiveMembers(store: WorkspaceStore, members: string[]): WorkspaceStore {
  const current = activeWorkspace(store);
  if (current.members.length === members.length && current.members.every((p, i) => p === members[i])) return store;
  return { ...store, list: store.list.map((w) => (w.id === current.id ? { ...w, members } : w)) };
}

/** A new, empty workspace, shown. */
export function withNewWorkspace(store: WorkspaceStore, id: string): WorkspaceStore {
  let n = store.list.length + 1;
  const taken = new Set(store.list.map((w) => w.name.toLowerCase()));
  while (taken.has(`workspace ${n}`)) n++;
  return { active: id, list: [...store.list, { id, name: `Workspace ${n}`, members: [] }] };
}

export function withRenamedWorkspace(store: WorkspaceStore, id: string, name: string): WorkspaceStore {
  const trimmed = name.trim();
  if (!trimmed) return store;
  const unique = uniqueWorkspaceName(store, trimmed, id);
  if (store.list.find((w) => w.id === id)?.name === unique) return store;
  return { ...store, list: store.list.map((w) => (w.id === id ? { ...w, name: unique } : w)) };
}

/** The store without workspace `id` (its files are not touched). The one
 * before it is shown next; deleting the last one leaves an empty one with
 * `freshId`. */
export function withoutWorkspace(store: WorkspaceStore, id: string, freshId: string): WorkspaceStore {
  const at = store.list.findIndex((w) => w.id === id);
  if (at < 0) return store;
  const list = store.list.filter((w) => w.id !== id);
  if (list.length === 0) return singleWorkspaceStore(freshId);
  const active = store.active === id ? list[Math.max(0, at - 1)].id : store.active;
  return { active, list };
}

/**
 * The saved workspaces. Throws when the settings cannot be read: the caller
 * must then not write, or a passing failure would replace what is saved.
 * `freshId` names the workspace made when none is saved yet (it takes the
 * list the first build kept, if there is one).
 */
export async function loadWorkspaceStore(freshId: string): Promise<{ store: WorkspaceStore; saved: boolean }> {
  const text = await invoke<string>("ssws_get_setting", { key: STORE_SETTING });
  const store = parseWorkspaceStore(text);
  if (store) return { store, saved: true };
  let legacy: string[] = [];
  try {
    legacy = pathList(JSON.parse(localStorage.getItem(LEGACY_KEY) ?? "[]"));
  } catch {
    /* nothing remembered */
  }
  return { store: singleWorkspaceStore(freshId, legacy), saved: false };
}

export async function saveWorkspaceStore(store: WorkspaceStore): Promise<void> {
  await invoke("ssws_set_setting", { key: STORE_SETTING, value: JSON.stringify(store) });
}

// ---------------------------------------------------------------------------
// The workspace that is shown (the Tauri side's list)
// ---------------------------------------------------------------------------

export function workspaceStatus(): Promise<WorkspaceMember[]> {
  return invoke<WorkspaceMember[]>("workspace_status");
}

/** Puts workbooks on the list (full paths; ones already there are kept). */
export function workspaceAddMembers(paths: string[]): Promise<WorkspaceMember[]> {
  return invoke<WorkspaceMember[]>("workspace_add_members", { paths });
}

/** Replaces the list: another saved workspace is shown. Tabs stay open. */
export function workspaceSetMembers(paths: string[]): Promise<WorkspaceMember[]> {
  return invoke<WorkspaceMember[]>("workspace_set_members", { paths });
}

/** Takes a workbook off the list. A workbook an open one reads stays loaded. */
export function workspaceRemoveMember(path: string): Promise<WorkspaceMember[]> {
  return invoke<WorkspaceMember[]>("workspace_remove_member", { path });
}

/** Dropped files and folders → the workbooks among them (a folder gives
 * the workbooks directly inside it). */
export function workspaceExpandPaths(paths: string[]): Promise<string[]> {
  return invoke<string[]>("workspace_expand_paths", { paths });
}

/** Writes a new, empty workbook to `path`. */
export function workspaceCreateWorkbook(path: string, sheets?: string[]): Promise<void> {
  return invoke<void>("workspace_create_workbook", { path, sheets: sheets ?? null });
}

/** File name of a path, for display. */
export function workbookName(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}
