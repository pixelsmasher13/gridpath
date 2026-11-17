import { describe, it, expect, vi, beforeEach } from "vitest";

// Saved workspaces: each a name and a list of models, kept as one settings
// row. The Tauri side is a fake settings table.
let settings: Record<string, string> = {};
let failReads = false;
vi.mock("@tauri-apps/api/core", () => ({
  invoke: async (cmd: string, args: any) => {
    if (cmd === "ssws_get_setting") {
      if (failReads) throw new Error("database is locked");
      return settings[args.key] ?? "";
    }
    if (cmd === "ssws_set_setting") {
      settings[args.key] = args.value;
      return null;
    }
    return null;
  },
}));

const api = await import("../workspaceApi");

let local: Record<string, string> = {};
beforeEach(() => {
  settings = {};
  failReads = false;
  local = {};
  (globalThis as any).localStorage = { getItem: (k: string) => local[k] ?? null, setItem: (k: string, v: string) => void (local[k] = v), removeItem: () => {} };
});

describe("saved workspaces", () => {
  it("starts with one workspace, which takes the list the first build kept", async () => {
    local["gridpath.workspaceMembers"] = JSON.stringify(["/m/NOW.xlsx", "/m/ANT.xlsx"]);
    const { store, saved } = await api.loadWorkspaceStore("w1");
    expect(saved).toBe(false);
    expect(store).toEqual({ active: "w1", list: [{ id: "w1", name: "Workspace 1", members: ["/m/NOW.xlsx", "/m/ANT.xlsx"] }] });
  });

  it("reads back what was saved", async () => {
    const store = api.withNewWorkspace(api.singleWorkspaceStore("w1", ["/m/NOW.xlsx"]), "w2");
    await api.saveWorkspaceStore(store);
    expect(await api.loadWorkspaceStore("unused")).toEqual({ store, saved: true });
  });

  it("fails loudly when the settings cannot be read, so nothing saved is replaced", async () => {
    await api.saveWorkspaceStore(api.singleWorkspaceStore("w1", ["/m/NOW.xlsx"]));
    failReads = true;
    await expect(api.loadWorkspaceStore("w9")).rejects.toThrow("database is locked");
  });

  it("ignores a saved value it cannot use", () => {
    expect(api.parseWorkspaceStore("")).toBeNull();
    expect(api.parseWorkspaceStore("{not json")).toBeNull();
    expect(api.parseWorkspaceStore(JSON.stringify({ active: "x", list: [] }))).toBeNull();
    // A shown workspace that is gone falls back to the first; bad entries are dropped.
    expect(api.parseWorkspaceStore(JSON.stringify({ active: "gone", list: [{ id: "a", name: "Comps", members: ["/m/A.xlsx", 3] }, { name: "no id" }] }))).toEqual({
      active: "a",
      list: [{ id: "a", name: "Comps", members: ["/m/A.xlsx"] }],
    });
  });

  it("makes a new workspace with a free name and shows it", () => {
    let store = api.singleWorkspaceStore("w1");
    store = api.withNewWorkspace(store, "w2");
    expect(store.active).toBe("w2");
    expect(store.list.map((w) => w.name)).toEqual(["Workspace 1", "Workspace 2"]);
    store = api.withRenamedWorkspace(store, "w1", "Workspace 3");
    store = api.withNewWorkspace(store, "w3");
    expect(store.list.map((w) => w.name)).toEqual(["Workspace 3", "Workspace 2", "Workspace 4"]);
  });

  it("renames, keeping names apart", () => {
    let store = api.withNewWorkspace(api.singleWorkspaceStore("w1"), "w2");
    store = api.withRenamedWorkspace(store, "w2", "  Software comps ");
    expect(api.activeWorkspace(store).name).toBe("Software comps");
    store = api.withRenamedWorkspace(store, "w1", "software comps");
    expect(store.list[0].name).toBe("software comps 2");
    // Its own name again, or nothing typed: no change.
    expect(api.withRenamedWorkspace(store, "w2", "Software comps")).toBe(store);
    expect(api.withRenamedWorkspace(store, "w2", "   ")).toBe(store);
  });

  it("saves the shown workspace's models, and only when they changed", () => {
    const store = api.withNewWorkspace(api.singleWorkspaceStore("w1", ["/m/NOW.xlsx"]), "w2");
    const added = api.withActiveMembers(store, ["/m/ANT.xlsx"]);
    expect(added.list).toEqual([
      { id: "w1", name: "Workspace 1", members: ["/m/NOW.xlsx"] },
      { id: "w2", name: "Workspace 2", members: ["/m/ANT.xlsx"] },
    ]);
    expect(api.withActiveMembers(added, ["/m/ANT.xlsx"])).toBe(added);
  });

  it("deletes a workspace, shows the one before it, and never leaves none", () => {
    let store = api.withNewWorkspace(api.withNewWorkspace(api.singleWorkspaceStore("w1"), "w2"), "w3");
    store = api.withoutWorkspace(store, "w3", "fresh");
    expect(store).toMatchObject({ active: "w2" });
    expect(store.list.map((w) => w.id)).toEqual(["w1", "w2"]);
    // One that is not shown: the shown one stays.
    expect(api.withoutWorkspace(store, "w1", "fresh")).toMatchObject({ active: "w2", list: [{ id: "w2" }] });
    store = api.withoutWorkspace(api.withoutWorkspace(store, "w2", "fresh"), "w1", "fresh");
    expect(store).toEqual({ active: "fresh", list: [{ id: "fresh", name: "Workspace 1", members: [] }] });
  });
});
