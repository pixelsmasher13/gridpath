import { describe, it, expect } from "vitest";
import { initialWorkspace, newTab, reduceWorkspace } from "../state/tabs";

describe("opening a tab", () => {
  it("takes the view, unless it is opened in the background", () => {
    const a = newTab("/m/SNPS.xlsx");
    const b = newTab("/m/Comps.xlsx");
    const c = newTab("/m/SNOW.xlsx");
    let state = reduceWorkspace(initialWorkspace, { type: "open", tab: a });
    expect(state.activeTabId).toBe(a.id);
    // The agent opens a workbook to write there: the user stays where they are.
    state = reduceWorkspace(state, { type: "open", tab: b, background: true });
    expect(state.tabs.map((t) => t.id)).toEqual([a.id, b.id]);
    expect(state.activeTabId).toBe(a.id);
    state = reduceWorkspace(state, { type: "open", tab: c });
    expect(state.activeTabId).toBe(c.id);
  });

  it("takes the view when nothing else is open, background or not", () => {
    const only = newTab("/m/Comps.xlsx");
    expect(reduceWorkspace(initialWorkspace, { type: "open", tab: only, background: true }).activeTabId).toBe(only.id);
  });
});
