/**
 * Bookkeeping for sheet ops the user performs through Univer's own sheet-tab
 * UI (add, delete, rename). These merge into the SaveMirror's sheetOps at
 * save time so the surgical patcher can explain every live sheet — otherwise
 * a hand-added sheet forces the lossy full-export fallback.
 */

export type SheetOp =
  | { kind: "create"; name: string; tabColor?: string | null }
  | { kind: "delete"; name: string }
  | { kind: "rename"; oldName: string; newName: string };

/**
 * Append an op to the per-tab store, compacting no-op histories in place:
 *
 * - delete of a sheet this session created (incl. via undo of the add)
 *   cancels the pending create instead of recording create+delete;
 * - rename of a session-created sheet rewrites the create op's name;
 * - consecutive renames (a→b, b→c) collapse to a→c.
 *
 * The patch replays against the ORIGINAL loaded base on every save, so the
 * store must describe the net session history, not just the latest click.
 */
export function appendManualSheetOp(store: SheetOp[], op: SheetOp): void {
  if (op.kind === "delete") {
    for (let i = store.length - 1; i >= 0; i--) {
      const prev = store[i];
      if (prev.kind === "create" && prev.name === op.name) {
        store.splice(i, 1);
        return;
      }
      // A rename chain touching this name — history isn't a simple
      // create/delete pair anymore; record the delete as-is.
      if (prev.kind === "rename" && (prev.newName === op.name || prev.oldName === op.name)) break;
    }
    store.push({ kind: "delete", name: op.name });
    return;
  }
  if (op.kind === "rename") {
    for (let i = store.length - 1; i >= 0; i--) {
      const prev = store[i];
      if (prev.kind === "create" && prev.name === op.oldName) {
        store[i] = { ...prev, name: op.newName };
        return;
      }
      if (prev.kind === "rename" && prev.newName === op.oldName) {
        store[i] = { ...prev, newName: op.newName };
        return;
      }
    }
    store.push(op);
    return;
  }
  store.push(op);
}
