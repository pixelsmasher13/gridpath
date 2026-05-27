import React, { useCallback, useEffect, useMemo, useReducer, useRef } from "react";
import styled from "styled-components";
import { useToast } from "@chakra-ui/react";
import { Settings as SettingsIcon, Save, PanelLeftOpen, PanelLeftClose, PanelRightClose, FolderOpen, FilePlus } from "lucide-react";
import { open, save } from "@tauri-apps/plugin-dialog";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { invoke } from "@tauri-apps/api/core";
import { v4 as uuidv4 } from "uuid";

import { UniverGrid, type UniverGridHandle, type SaveMirror } from "./components/UniverGrid";
import { StatusBar } from "./components/StatusBar";
import { UpdateNotification } from "./components/UpdateNotification";
import { ChatPanel } from "./components/ChatPanel";
import { TabBar } from "./components/TabBar";
import { SessionSidebar } from "./components/SessionSidebar";
import { SettingsModal } from "./components/SettingsModal";
import { ExitGuardModal } from "./components/ExitGuardModal";
import { getSettingValue, setSettingValue, SETTING_KEYS } from "./settingsApi";
import {
  readWorkbookBytes,
  writeWorkbookBytes,
  readUntitledSnapshot,
  writeUntitledSnapshot,
  appendChangeBatch,
  readChangeLog,
} from "./workbookIo";
import type { ChangeBatch, UniverMutation, FormatMutation } from "./types";
import {
  initialWorkspace,
  reduceWorkspace,
  newTab,
  findTab,
  findTabByPath,
  sessionNameFromPrompt,
  type WorkbookTab,
} from "./state/tabs";
import {
  startAgentTurn,
  stopAgentTurn,
  reportToolResult,
  subscribeAgentEvents,
  type AgentEvent,
} from "./agent/agentClient";
import { interpretToolCall } from "./agent/toolToMutation";
import { expandA1Range } from "./agent/toolToMutation";
import { captureWorkbookContext } from "./agent/captureContext";
import { buildPriorBatchesContext } from "./agent/priorContext";
import { buildFocusContext } from "./agent/selectionContext";
import {
  upsertSession,
  renameSession as renameSessionDb,
  appendMessage,
  addSessionTokens,
  listSessions,
  archiveSession,
  deleteSession,
  getMessages,
  type SessionRow,
} from "./sessionDb";

const Page = styled.div<{
  $sidebarOpen: boolean;
  $chatWidth: number;
  $chatOpen: boolean;
  $hasTabs: boolean;
}>`
  display: grid;
  grid-template-rows: 38px ${(p) => (p.$hasTabs ? "52px" : "0")} 1fr 24px;
  grid-template-columns:
    ${(p) => (p.$sidebarOpen ? "260px" : "0")}
    1fr
    ${(p) => (p.$chatOpen ? `${p.$chatWidth}px` : "0")};
  grid-template-areas:
    "header  header  header"
    "sidebar tabs    chat"
    "sidebar grid    chat"
    "status  status  status";
  height: 100vh;
  width: 100vw;
  background: #1e1e1e;
  color: #d4d4d4;
  font-family: -apple-system, "SF Pro Text", "Inter", system-ui, sans-serif;
`;

const ChatResizer = styled.div`
  position: absolute;
  left: -3px;
  top: 0;
  bottom: 0;
  width: 6px;
  cursor: col-resize;
  z-index: 5;
  background: transparent;
  &:hover { background: rgba(51, 99, 173, 0.35); }
  &:active { background: rgba(51, 99, 173, 0.6); }
`;

const SidebarArea = styled.div`
  grid-area: sidebar;
  overflow: hidden;
`;

const ChatArea = styled.div`
  grid-area: chat;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  position: relative;
`;

const Header = styled.div`
  grid-area: header;
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 0 12px;
  border-bottom: 1px solid #2a2a2a;
  font-size: 12px;
  color: #b3b3b3;
`;

const HeaderButton = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: transparent;
  border: 1px solid #333;
  color: #d4d4d4;
  border-radius: 5px;
  padding: 4px 10px;
  font-size: 12px;
  cursor: pointer;
  &:hover:not(:disabled) { background: #2a2a2a; }
  &:disabled { opacity: 0.4; cursor: not-allowed; }
`;

const TabsArea = styled.div`
  grid-area: tabs;
  overflow: hidden;
`;

const GridArea = styled.div`
  grid-area: grid;
  position: relative;
  overflow: hidden;
  background: #1e1e1e;
`;

const EmptyState = styled.div`
  display: flex;
  align-items: center;
  justify-content: center;
  height: 100%;
  color: #6f6f6f;
  font-size: 14px;
  flex-direction: column;
  gap: 16px;
  padding: 24px;
  text-align: center;
`;

const EmptyButtons = styled.div`
  display: flex;
  gap: 10px;
`;

const EmptyPrimaryBtn = styled.button`
  background: #3363AD;
  color: #fff;
  border: 0;
  border-radius: 6px;
  padding: 10px 18px;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  &:hover { background: #4275c4; }
`;

const EmptySecondaryBtn = styled.button`
  background: transparent;
  color: #d4d4d4;
  border: 1px solid #444;
  border-radius: 6px;
  padding: 10px 18px;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  &:hover { background: #2a2a2a; border-color: #3363AD; }
`;

const GridLayer = styled.div<{ $visible: boolean }>`
  position: absolute;
  inset: 0;
  visibility: ${(p) => (p.$visible ? "visible" : "hidden")};
`;

const StatusArea = styled.div`
  grid-area: status;
`;

export const SpreadsheetScreen: React.FC = () => {
  const toast = useToast();
  const [workspace, dispatch] = useReducer(reduceWorkspace, initialWorkspace);
  const [promptByTab, setPromptByTab] = React.useState<Record<string, string>>({});
  // Layout state — persisted to localStorage so it survives reloads.
  const [sidebarOpen, setSidebarOpen] = React.useState<boolean>(() => {
    try { return localStorage.getItem("ssws_sidebar_open") !== "0"; } catch { return true; }
  });
  const [chatOpen, setChatOpen] = React.useState<boolean>(() => {
    try { return localStorage.getItem("ssws_chat_open") !== "0"; } catch { return true; }
  });
  const [chatWidth, setChatWidth] = React.useState<number>(() => {
    try {
      const v = parseInt(localStorage.getItem("ssws_chat_width") || "0", 10);
      return v >= 280 && v <= 900 ? v : 400;
    } catch { return 400; }
  });
  useEffect(() => { localStorage.setItem("ssws_sidebar_open", sidebarOpen ? "1" : "0"); }, [sidebarOpen]);
  useEffect(() => { localStorage.setItem("ssws_chat_open", chatOpen ? "1" : "0"); }, [chatOpen]);
  useEffect(() => { localStorage.setItem("ssws_chat_width", String(chatWidth)); }, [chatWidth]);

  // Drag-to-resize the chat panel. Min 280, max 900 px.
  const dragStateRef = useRef<{ startX: number; startWidth: number } | null>(null);
  const startChatResize = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    dragStateRef.current = { startX: e.clientX, startWidth: chatWidth };
    const onMove = (ev: MouseEvent) => {
      const st = dragStateRef.current;
      if (!st) return;
      // Chat is on the right — dragging LEFT widens it.
      const next = Math.max(280, Math.min(900, st.startWidth - (ev.clientX - st.startX)));
      setChatWidth(next);
    };
    const onUp = () => {
      dragStateRef.current = null;
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", onUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
    document.body.style.cursor = "col-resize";
    document.body.style.userSelect = "none";
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
  }, [chatWidth]);

  const [recentSessions, setRecentSessions] = React.useState<SessionRow[]>([]);
  const [settingsOpen, setSettingsOpen] = React.useState(false);
  // Auto-apply: when on, every agent batch is auto-accepted on done — no
  // manual review step. Persisted to the settings DB so it survives
  // restart. Loaded on mount; toggling writes back.
  const [autoApply, setAutoApplyState] = React.useState(false);
  useEffect(() => {
    (async () => {
      const v = await getSettingValue(SETTING_KEYS.autoApply);
      setAutoApplyState(v === "1");
    })();
  }, []);
  const setAutoApply = useCallback((next: boolean) => {
    setAutoApplyState(next);
    setSettingValue(SETTING_KEYS.autoApply, next ? "1" : "0").catch((e) =>
      console.warn("[settings] save autoApply failed:", e),
    );
  }, []);
  // Ref mirror so async handlers (the agent done callback) see the latest
  // value without closure staleness.
  const autoApplyRef = useRef(autoApply);
  useEffect(() => { autoApplyRef.current = autoApply; }, [autoApply]);

  // (liveSelection effect lives below — it depends on `activeTab` being declared.)
  const [liveSelection, setLiveSelection] = React.useState<string | null>(null);
  /**
   * When the user dismisses the selection chip we stash the dismissed
   * label here. The chip stays hidden while `liveSelection === focusDismissedFor`.
   * Once Univer reports a different selection, the chip re-appears
   * automatically (the auto-reset effect below clears this).
   */
  const [focusDismissedFor, setFocusDismissedFor] = React.useState<string | null>(null);
  useEffect(() => {
    if (focusDismissedFor && liveSelection !== focusDismissedFor) {
      setFocusDismissedFor(null);
    }
  }, [liveSelection, focusDismissedFor]);

  const refreshSessions = useCallback(async () => {
    try {
      const rows = await listSessions(50);
      setRecentSessions(rows);
    } catch (e) {
      console.warn("[session] list failed:", e);
    }
  }, []);

  useEffect(() => { refreshSessions(); }, [refreshSessions]);
  // Cheap polling — keeps the sidebar fresh after rename/archive/new prompt.
  // Backed by indexed updated_at; ~50 rows max. Skip if no one's looking.
  useEffect(() => {
    if (!sidebarOpen) return;
    const id = setInterval(refreshSessions, 5000);
    return () => clearInterval(id);
  }, [sidebarOpen, refreshSessions]);

  // We read the latest workspace from this ref inside async event callbacks
  // that would otherwise close over stale state (e.g. the agent done handler
  // that persists the just-finished batch to the DB).
  const workspaceRef = useRef(workspace);
  useEffect(() => { workspaceRef.current = workspace; }, [workspace]);

  // Warn before quitting when any open workbook has unsaved changes. We
  // hook Tauri's `onCloseRequested`, preventDefault the close, and pop
  // an in-app modal (ExitGuardModal) with Save / Discard / Cancel — the
  // standard three-button "discard work?" pattern shipped by every
  // editor. Native dialog was ugly; the in-app one matches our color
  // scheme. Dirty currently flips on agent edits + undo/redo; manual
  // cell typing isn't tracked yet (separate fix).
  const [exitGuard, setExitGuard] = React.useState<{
    dirtyTabs: WorkbookTab[];
    saving: boolean;
  } | null>(null);
  useEffect(() => {
    let unlistenFn: (() => void) | undefined;
    (async () => {
      try {
        const win = getCurrentWindow();
        unlistenFn = await win.onCloseRequested((event) => {
          const dirtyTabs = workspaceRef.current.tabs.filter((t) => t.dirty);
          if (dirtyTabs.length === 0) return;
          event.preventDefault();
          setExitGuard({ dirtyTabs, saving: false });
        });
      } catch (e) {
        console.warn("[exit-guard] failed to install onCloseRequested:", e);
      }
    })();
    return () => { unlistenFn?.(); };
  }, []);

  // We mount one UniverGrid per tab and keep them all alive — switching
  // tabs just toggles visibility, which preserves scroll position and
  // Univer's internal state. The grid refs are kept in a map keyed by tabId.
  const gridRefs = useRef<Record<string, UniverGridHandle | null>>({});

  // For each batch we keep the cells we tinted so Accept/Reject can clean up.
  const tintedCellsByBatch = useRef<Record<string, Array<{ sheet: string; row: number; col: number }>>>({});
  // Capture pre-edit values so Reject can restore them.
  const oldValuesByBatch = useRef<
    Record<string, Array<{ sheet: string; row: number; col: number; oldValue: any; oldFormula: string | null }>>
  >({});
  // Synchronous accumulator for the agent's streaming prose, keyed by
  // batch id. We mutate this ref directly inside the text_delta handler
  // so the `done` handler can read the latest text without depending on
  // React having committed the corresponding state update — that race
  // was eating the tail of fast/short replies (the agent would emit
  // "You" then "'re welcome…" and done would fire before React synced
  // workspaceRef, persisting just "You" into batch.agent_text).
  const streamingTextByBatch = useRef<Record<string, string>>({});
  // Per-tab serial queue for tool-call processing. The Rust agent loop
  // can fire tool_calls faster than Univer can apply them — when a giant
  // set_range hits the webview, the next set_range arrives while
  // Univer/React are still committing the prior one. Without serializing,
  // mutations overwrite each other mid-commit and the reportToolResult
  // chain races. This ref holds the *last* pending task in the chain;
  // each new tool_call appends to it, guaranteeing strict ordering.
  const toolCallQueueByTab = useRef<Record<string, Promise<void>>>({});
  // Counter of in-flight tool_call tasks across all tabs. The auto-snapshot
  // loop checks this so a 5–30s exportBytes() doesn't contend with the
  // agent's readback path on the JS main thread — a snapshot mid-tool_call
  // wedges the toolCallQueue and triggers cascading 25-28s timeouts in the
  // Rust agent loop. We only snapshot when this counter is zero.
  const pendingToolCalls = useRef(0);
  const enqueueToolCallTask = (tabId: string, task: () => Promise<void>) => {
    const prior = toolCallQueueByTab.current[tabId] ?? Promise.resolve();
    pendingToolCalls.current++;
    const next = prior.catch(() => {}).then(task).finally(() => {
      pendingToolCalls.current = Math.max(0, pendingToolCalls.current - 1);
    });
    toolCallQueueByTab.current[tabId] = next;
  };

  const activeTab = findTab(workspace, workspace.activeTabId);
  const promptForActive = activeTab ? promptByTab[activeTab.id] ?? "" : "";

  // Lightweight polling of the active tab's selection so the composer chip
  // updates without wiring a Univer selection-change listener. 500ms is
  // invisible to humans and cheap.
  useEffect(() => {
    if (!activeTab) {
      setLiveSelection(null);
      return;
    }
    let cancelled = false;
    const tick = () => {
      if (cancelled) return;
      const sel = gridRefs.current[activeTab.id]?.getActiveSelection?.() ?? null;
      if (!sel) {
        setLiveSelection(null);
        return;
      }
      // Show the chip for any selection — single cell OR range. Single-cell
      // chips give the user a visible "the agent saw your cursor" cue, even
      // though we still skip shipping pure 1×1 selections as "focus" in
      // selectionContext (those are usually just where the cursor parked).
      const sameCell = sel.endRow === sel.startRow && sel.endCol === sel.startCol;
      const count = (sel.endRow - sel.startRow + 1) * (sel.endCol - sel.startCol + 1);
      const label = sameCell
        ? `${sel.sheet}!${a1Of(sel.startRow, sel.startCol)}`
        : `${sel.sheet}!${a1Of(sel.startRow, sel.startCol)}:${a1Of(sel.endRow, sel.endCol)} (${count} cells)`;
      setLiveSelection(label);
    };
    tick();
    const id = setInterval(tick, 500);
    return () => { cancelled = true; clearInterval(id); };
  }, [activeTab?.id]);

  // --- file IO -------------------------------------------------------------

  /**
   * Create a brand-new blank workbook tab without touching disk.
   * Path stays as `untitled-{uuid}` (no .xlsx suffix) so the first Save
   * naturally falls into the save-as branch and prompts for a real
   * location. We don't write a session row to the DB until save either —
   * matches VS Code's "Untitled-1" behavior, no orphan rows in the
   * sidebar for files that don't exist yet.
   */
  const createBlankWorkbook = useCallback(() => {
    const id = uuidv4();
    const path = `untitled-${id.slice(0, 8)}`;
    const tab: WorkbookTab = {
      ...newTab(path),
      id,
      filename: `Untitled-${id.slice(0, 4)}.xlsx`,
    };
    dispatch({ type: "open", tab });
    // No loadBytes needed — UniverGrid's mount creates a default blank
    // workbook. Mark dirty so the user sees the · indicator until they save.
    setTimeout(() => {
      dispatch({ type: "mark_dirty", tabId: tab.id, dirty: true });
    }, 50);
  }, []);

  const openWorkbookDialog = useCallback(async () => {
    try {
      const selected = await open({
        multiple: false,
        filters: [{ name: "Excel", extensions: ["xlsx"] }],
      });
      if (!selected || typeof selected !== "string") return;

      const existing = findTabByPath(workspace, selected);
      if (existing) {
        dispatch({ type: "activate", tabId: existing.id });
        return;
      }

      const tab = newTab(selected);
      dispatch({ type: "open", tab });

      // The new GridLayer mounts on the next React commit, and the
      // useImperativeHandle ref attaches in the commit's layout phase.
      // RAF fires *before* commit, so polling until the ref appears is the
      // reliable way to wait. ~1s cap with 25ms ticks ≈ 40 attempts.
      const waitForGrid = async (): Promise<UniverGridHandle | null> => {
        for (let i = 0; i < 40; i++) {
          const g = gridRefs.current[tab.id];
          if (g) return g;
          await new Promise((r) => setTimeout(r, 25));
        }
        return null;
      };

      (async () => {
        const grid = await waitForGrid();
        if (!grid) {
          console.error("[open] grid ref never attached for tab", tab.id);
          dispatch({ type: "close", tabId: tab.id });
          return;
        }
        try {
          const bytes = await readWorkbookBytes(selected);
          await grid.loadBytes(bytes);
          console.log("[open] xlsx loaded into tab", tab.id);
        } catch (err) {
          toast({ title: "Open failed", description: String(err), status: "error", duration: 4000 });
          dispatch({ type: "close", tabId: tab.id });
          return;
        }
        // Opening a file is "I want a fresh session on this file" — we don't
        // replay the on-disk .changes.jsonl audit log into the new session's
        // batches. The jsonl is a per-file append-only audit; if the user
        // wants to resume a prior session, they click it in the sidebar
        // (which loads that session's messages from the DB).
        //
        // We also DON'T upsert a DB row here — that's deferred to the first
        // prompt so we don't fill the sidebar with untitled entries for
        // files the user just browsed and closed.
      })();
    } catch (err) {
      toast({ title: "Open failed", description: String(err), status: "error", duration: 4000 });
    }
  }, [workspace, toast]);

  /**
   * Resume an existing session from the sidebar: open the xlsx in a new tab,
   * replay the message log into the changes panel, mark the tab as the
   * persisted session id so subsequent prompts append to the same row.
   *
   * v1: pending batches from the log are restored as "pending" — user can
   * still Accept/Reject them post-restart. Accepted/rejected batches show
   * with their final status (read-only).
   */
  const resumeSession = useCallback(
    async (row: SessionRow) => {
      const existing = findTabByPath(workspace, row.workbook_path);
      if (existing) {
        dispatch({ type: "activate", tabId: existing.id });
        return;
      }
      // Build a tab whose id matches the persisted session id so all
      // subsequent persists land in the right row.
      const tab: WorkbookTab = (() => {
        const fresh = newTab(row.workbook_path);
        return {
          ...fresh,
          id: row.id,
          name: row.name,
          // Hydrate lifetime token totals from the session row so the
          // Usage tab shows cumulative spend across app restarts, not
          // just usage from this freshly-reopened tab.
          inputTokens: row.total_input_tokens ?? 0,
          outputTokens: row.total_output_tokens ?? 0,
          cacheReadTokens: row.total_cache_read_tokens ?? 0,
          cacheCreationTokens: row.total_cache_creation_tokens ?? 0,
        };
      })();
      dispatch({ type: "open", tab });

      // Same poll-until-attached pattern as openWorkbookDialog — RAF fires
      // before React commits, so we can't rely on it for the ref.
      const waitForGrid = async (): Promise<UniverGridHandle | null> => {
        for (let i = 0; i < 40; i++) {
          const g = gridRefs.current[tab.id];
          if (g) return g;
          await new Promise((r) => setTimeout(r, 25));
        }
        return null;
      };

      (async () => {
        const grid = await waitForGrid();
        if (!grid) {
          console.error("[resume] grid ref never attached for tab", tab.id);
          dispatch({ type: "close", tabId: tab.id });
          return;
        }
        // For untitled drafts, prefer the Univer-native JSON snapshot
        // we wrote in snapshotUntitled — lossless format preservation.
        // Fall back to xlsx bytes for: real saved files (always xlsx),
        // and legacy untitled drafts that predate the .gpsnap path.
        // If NEITHER exists for an untitled draft, that's not an error —
        // the user created a tab but never prompted (no auto-snapshot
        // fire) and quit. Treat it as a blank workbook: the grid is
        // already in newTab's empty state.
        const isUntitled = row.workbook_path.startsWith("untitled-");
        try {
          const snap = isUntitled ? await readUntitledSnapshot(row.workbook_path) : null;
          if (snap) {
            await grid.loadSnapshot(snap);
          } else {
            try {
              const bytes = await readWorkbookBytes(row.workbook_path);
              await grid.loadBytes(bytes);
            } catch (err) {
              if (isUntitled) {
                // No persisted state for this untitled draft — fine, start blank.
                console.info("[resume] untitled draft has no persisted state, opening blank:", row.workbook_path);
              } else {
                throw err;
              }
            }
          }
        } catch (err) {
          toast({ title: "Open failed", description: String(err), status: "error", duration: 4000 });
          dispatch({ type: "close", tabId: tab.id });
          return;
        }
        try {
          const msgs = await getMessages(row.id);
          const replayed: ChangeBatch[] = [];
          for (const m of msgs) {
            if (m.role === "agent_batch") {
              try {
                const parsed = JSON.parse(m.payload) as { batch: ChangeBatch };
                if (parsed.batch) replayed.push(parsed.batch);
              } catch {}
            }
          }
          const latest = new Map<string, ChangeBatch>();
          for (const b of replayed) latest.set(b.id, b);
          dispatch({
            type: "batches_replace",
            tabId: tab.id,
            batches: Array.from(latest.values()),
          });
        } catch (e) {
          console.warn("[session] message replay failed:", e);
        }
        try {
          await upsertSession(tab.id, tab.name, tab.path);
        } catch {}
      })();
    },
    [workspace, toast],
  );

  // Persist an untitled draft to <app_data>/untitled_sessions/untitled-XXX
  // as a real .xlsx so it can be re-opened from the sidebar after a restart.
  // Saved-path tabs are no-op (their real .xlsx is already on disk).
  const snapshotUntitled = useCallback(async (tab: WorkbookTab) => {
    if (!tab.path.startsWith("untitled-")) return;
    const grid = gridRefs.current[tab.id];
    if (!grid) return;
    try {
      // Use Univer's native JSON snapshot (lossless) instead of xlsx
      // bytes. xlsx round-trips drop fills, custom number formats,
      // theme colors, view state, etc. — fine for the user's eventual
      // Save As (they review the file) but unacceptable for invisible
      // auto-snapshots used to restore drafts in-app.
      const snap = grid.getWorkbookSnapshot?.();
      if (!snap) return;
      await writeUntitledSnapshot(tab.path, snap);
    } catch (e) {
      console.warn("[snapshot] untitled snapshot failed:", e);
    }
  }, []);

  const closeTab = useCallback(
    async (tabId: string) => {
      const tab = findTab(workspace, tabId);
      if (tab) await snapshotUntitled(tab);
      delete gridRefs.current[tabId];
      dispatch({ type: "close", tabId });
    },
    [workspace, snapshotUntitled],
  );

  // Auto-snapshot dirty untitled drafts every 30s so a hard quit (or kernel
  // panic, OOM, etc.) doesn't lose the user's WIP. Saved-path tabs are
  // excluded — they're the user's responsibility via Cmd-S.
  //
  // Skips when the agent has tool_calls in flight: exportBytes() and the
  // tool_call readback path both hit Univer on the main thread, and
  // contention there manifests as 25-28s tool_result timeouts in the
  // Rust agent loop (see pendingToolCalls). Snapshot will fire on the
  // next tick (≤30s later) once the agent goes idle, plus closeTab
  // always snapshots regardless.
  useEffect(() => {
    const id = setInterval(() => {
      if (pendingToolCalls.current > 0) return;
      for (const t of workspace.tabs) {
        if (t.dirty && t.path.startsWith("untitled-")) {
          void snapshotUntitled(t);
        }
      }
    }, 30000);
    return () => clearInterval(id);
  }, [workspace.tabs, snapshotUntitled]);

  /**
   * Save a specific tab. Returns true on success, false if the user
   * dismissed a Save-As dialog (untitled tab) or the export failed.
   * `forceAsk` always prompts for a path (Save As semantics).
   */
  const saveTabBytes = useCallback(
    async (tab: WorkbookTab, forceAsk: boolean, quiet: boolean): Promise<boolean> => {
      try {
        const mirror = buildSaveMirror(tab.batches);
        const bytes = await gridRefs.current[tab.id]?.exportBytes(mirror);
        if (!bytes) throw new Error("nothing to export");

        let target = tab.path;
        const isUntitled = !target.toLowerCase().endsWith(".xlsx");
        if (isUntitled || forceAsk) {
          const chosen = await save({
            defaultPath: isUntitled ? tab.filename : tab.path,
            filters: [{ name: "Excel", extensions: ["xlsx"] }],
          });
          if (!chosen) return false;
          target = chosen;
        }

        await writeWorkbookBytes(target, bytes);
        const wasUntitled = tab.path.startsWith("untitled-");
        const isNewPath = target !== tab.path;
        dispatch({ type: "rename", tabId: tab.id, path: target });
        dispatch({ type: "mark_saved", tabId: tab.id, at: Date.now() });
        if (wasUntitled || isNewPath) {
          upsertSession(tab.id, tab.name || "", target).catch((e) =>
            console.warn("[session] upsert on save failed:", e),
          );
        }
        if (!quiet) {
          toast({ title: forceAsk ? "Saved as copy" : "Saved", status: "success", duration: 1800 });
        }
        return true;
      } catch (err) {
        toast({ title: "Save failed", description: String(err), status: "error", duration: 4000 });
        return false;
      }
    },
    [toast],
  );

  const doSave = useCallback(
    async (forceAsk: boolean) => {
      if (!activeTab) return;
      await saveTabBytes(activeTab, forceAsk, false);
    },
    [activeTab, saveTabBytes],
  );

  const saveActive = useCallback(() => doSave(false), [doSave]);
  const saveActiveAs = useCallback(() => doSave(true), [doSave]);

  // --- agent loop ----------------------------------------------------------

  const submitPrompt = useCallback(async () => {
    if (!activeTab) {
      toast({ title: "Open a workbook first", status: "info", duration: 2500 });
      return;
    }
    const prompt = (promptByTab[activeTab.id] ?? "").trim();
    if (!prompt) return;

    const batchId = uuidv4();
    const batch: ChangeBatch = {
      id: batchId,
      prompt,
      justification: "",
      mutations: [],
      status: "streaming",
      created_at: new Date().toISOString(),
    };
    // Auto-name the session from the very first prompt (heelix_notes-style).
    // If the user already renamed manually, we leave it alone.
    const autoName = !activeTab.name ? sessionNameFromPrompt(prompt) : null;
    if (autoName) {
      dispatch({ type: "set_name", tabId: activeTab.id, name: autoName });
    }

    // First prompt on this tab? Ensure the session exists in the DB now —
    // we deferred creation when the file was opened so the sidebar doesn't
    // fill with "untitled" rows for files the user just browsed. upsert is
    // idempotent: creates if missing, updates name + last_opened_at if not.
    //
    // For new blank workbooks (path starts with "untitled-") we ALSO upsert
    // so the session is visible in the sidebar while the user works —
    // otherwise blank-workbook sessions only appear after first save, which
    // is jarringly asymmetric vs. opened-existing files. Orphaned
    // untitled-* sessions get pruned on next app start (see the prune
    // effect on mount).
    upsertSession(activeTab.id, autoName ?? activeTab.name ?? "", activeTab.path).catch((e) =>
      console.warn("[session] upsert on first prompt failed:", e),
    );

    // Persist the user prompt as the first message of this turn.
    appendMessage(activeTab.id, "user", { prompt }).catch((e) =>
      console.warn("[session] append user message failed:", e),
    );

    dispatch({ type: "batch_add", tabId: activeTab.id, batch });
    dispatch({ type: "set_agent_running", tabId: activeTab.id, running: true });
    dispatch({ type: "stream_text_clear", tabId: activeTab.id });
    dispatch({
      type: "set_status",
      tabId: activeTab.id,
      phase: "thinking",
      message: "Connecting to Claude…",
    });
    setPromptByTab((s) => ({ ...s, [activeTab.id]: "" }));

    tintedCellsByBatch.current[batchId] = [];
    oldValuesByBatch.current[batchId] = [];

    try {
      const grid = gridRefs.current[activeTab.id] ?? null;
      const workbookContext = captureWorkbookContext(activeTab.path, grid);
      // If the user dismissed the selection chip for the currently
      // displayed range, skip the selection-as-focus block — @-mentions
      // in the prompt still flow through. We can't easily filter only
      // the selection portion from buildFocusContext, so we just skip
      // it entirely when dismissal matches the current selection. Any
      // @-mention falls through buildFocusContext's mentions path on
      // a clean re-call against the same grid (without selection).
      const focusDismissed = focusDismissedFor !== null && focusDismissedFor === liveSelection;
      const focus = focusDismissed
        ? buildFocusContext(prompt, { ...grid!, getActiveSelection: () => null } as any)
        : buildFocusContext(prompt, grid);
      if (focus) workbookContext.focus = focus.text;
      const priorBatchesContext = buildPriorBatchesContext(activeTab.batches);
      await startAgentTurn({
        tabId: activeTab.id,
        batchId,
        prompt,
        workbookContext,
        priorBatchesContext,
      });
    } catch (err) {
      dispatch({
        type: "set_status",
        tabId: activeTab.id,
        phase: "error",
        message: `Agent error: ${String(err)}`,
      });
      dispatch({ type: "set_agent_running", tabId: activeTab.id, running: false });
      dispatch({ type: "batch_finalize", tabId: activeTab.id, batchId });
      toast({ title: "Agent error", description: String(err), status: "error", duration: 5000 });
    }
  }, [activeTab, promptByTab, toast]);

  // Listen for agent events from Rust and dispatch into the right tab.
  useEffect(() => {
    const unsub = subscribeAgentEvents((ev: AgentEvent) => {
      // Verbose during v1 dev — flip to a debug flag once the loop is stable.
      console.log("[agent] event:", ev);
      switch (ev.kind) {
        case "started":
          dispatch({ type: "stream_text_clear", tabId: ev.tab_id });
          dispatch({
            type: "set_status",
            tabId: ev.tab_id,
            phase: "thinking",
            message: "Claude is thinking…",
          });
          return;
        case "text_delta":
          // Synchronously accumulate into the per-batch ref so the `done`
          // handler can read the full text without racing React's commit.
          streamingTextByBatch.current[ev.batch_id] =
            (streamingTextByBatch.current[ev.batch_id] ?? "") + ev.delta;
          dispatch({ type: "stream_text_append", tabId: ev.tab_id, delta: ev.delta });
          return;
        case "tool_call": {
          // Whole tool_call cycle (interpret → mutate → settle → report)
          // serialized per-tab. Rationale: see toolCallQueueByTab declaration.
          enqueueToolCallTask(ev.tab_id, async () => {
          const result = interpretToolCall(ev.name, ev.input);
          if (result.kind === "ignored") {
            console.warn("agent: ignored tool call:", result.reason);
            // Report the failure back so the Rust loop doesn't hang.
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({ error: result.reason }),
              );
            } catch {}
            return;
          }
          if (result.kind === "done") {
            dispatch({
              type: "batch_set_justification",
              tabId: ev.tab_id,
              batchId: ev.batch_id,
              justification: result.justification,
            });
            return;
          }
          if (result.kind === "read") {
            // Read-only tool — no mutations, no batch update, no dirty
            // flag change. Just look up the requested range from Univer
            // and ship the cells back so the agent can sanity-check its
            // own work or verify assumption-block addresses before
            // composing dependent formulas.
            const grid = gridRefs.current[ev.tab_id];
            const rangeCells = expandA1Range(result.range);
            const cells = rangeCells.slice(0, 500).map(({ row, col }) => {
              const c = grid?.getCell(result.sheet, row, col);
              return {
                cell: a1Of(row, col),
                value: c?.value ?? null,
                formula: c?.formula ?? null,
              };
            });
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({
                  sheet: result.sheet,
                  range: result.range,
                  cells: cells.filter((c) => c.value !== null || c.formula !== null),
                  truncated: rangeCells.length > 500,
                }),
              );
            } catch (e) {
              console.warn("[agent] read_range reportToolResult failed:", e);
            }
            return;
          }
          if (result.kind === "fetch") {
            dispatch({
              type: "batch_add_fetched_urls",
              tabId: ev.tab_id,
              batchId: ev.batch_id,
              urls: result.urls,
            });
            dispatch({
              type: "set_status",
              tabId: ev.tab_id,
              phase: "thinking",
              message: `Fetching ${result.urls.length} page${result.urls.length === 1 ? "" : "s"}…`,
            });
            // Rust handles the actual fetch + tool_result — we just show the chip.
            return;
          }
          // Apply mutations live: capture old state, write new.
          const grid = gridRefs.current[ev.tab_id];
          for (const raw of result.mutations) {
            if (raw.type === "set_cell") {
              const { sheet, row, col } = raw.address;
              const before = grid?.getCell(sheet, row, col);
              const m: UniverMutation = {
                ...raw,
                old_value: before?.value ?? null,
                old_formula: before?.formula ?? null,
              };
              grid?.setCell(sheet, row, col, m.new_formula ?? m.new_value);
              tintedCellsByBatch.current[ev.batch_id]?.push({ sheet, row, col });
              oldValuesByBatch.current[ev.batch_id]?.push({
                sheet, row, col,
                oldValue: before?.value ?? null,
                oldFormula: before?.formula ?? null,
              });
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: m });
              continue;
            }
            if (raw.type === "set_format") {
              // Capture each cell's pre-format snapshot so Reject can restore.
              const fm = raw as FormatMutation;
              const old_format = fm.cells.map(({ row, col }) => ({
                row, col,
                format: grid?.getCellFormat(fm.sheet, row, col) ?? null,
              }));
              // Apply the new format immediately so the user previews the look.
              for (const { row, col } of fm.cells) {
                grid?.setCellFormat(fm.sheet, row, col, fm.new_format);
                tintedCellsByBatch.current[ev.batch_id]?.push({ sheet: fm.sheet, row, col });
              }
              const m: UniverMutation = { ...fm, old_format };
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: m });
              continue;
            }
            if (raw.type === "set_column_width") {
              const old_widths = raw.columns.map((col) => ({
                col,
                width: grid?.getColumnWidth(raw.sheet, col) ?? null,
              }));
              for (const col of raw.columns) grid?.setColumnWidth(raw.sheet, col, raw.new_width);
              const m: UniverMutation = { ...raw, old_widths };
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: m });
              continue;
            }
            if (raw.type === "set_row_height") {
              const old_heights = raw.rows.map((row) => ({
                row,
                height: grid?.getRowHeight(raw.sheet, row) ?? null,
              }));
              for (const row of raw.rows) grid?.setRowHeight(raw.sheet, row, raw.new_height);
              const m: UniverMutation = { ...raw, old_heights };
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: m });
              continue;
            }
            if (raw.type === "merge_cells") {
              grid?.mergeCells(raw.sheet, raw.start_row, raw.start_col, raw.end_row, raw.end_col);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "unmerge_cells") {
              grid?.unmergeCells(raw.sheet, raw.start_row, raw.start_col, raw.end_row, raw.end_col);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "create_sheet") {
              const ok = grid?.createSheet(raw.name, raw.tab_color ?? null);
              if (ok) dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "delete_sheet") {
              const ok = grid?.deleteSheet(raw.name);
              if (ok) dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "rename_sheet") {
              const ok = grid?.renameSheet(raw.old_name, raw.new_name);
              if (ok) dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "clear_range") {
              // Capture pre-clear values for reject restoration.
              const enriched = {
                ...raw,
                cells: raw.cells.map((c) => {
                  const before = grid?.getCell(raw.sheet, c.row, c.col);
                  return {
                    row: c.row,
                    col: c.col,
                    old_value: before?.value ?? null,
                    old_formula: before?.formula ?? null,
                  };
                }),
              };
              for (const c of enriched.cells) {
                grid?.setCell(raw.sheet, c.row, c.col, null);
                tintedCellsByBatch.current[ev.batch_id]?.push({ sheet: raw.sheet, row: c.row, col: c.col });
              }
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: enriched });
              continue;
            }
            if (raw.type === "insert_rows") {
              grid?.insertRows(raw.sheet, raw.before, raw.count);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "delete_rows") {
              grid?.deleteRows(raw.sheet, raw.start, raw.count);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "insert_columns") {
              grid?.insertColumns(raw.sheet, raw.before, raw.count);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "delete_columns") {
              grid?.deleteColumns(raw.sheet, raw.start, raw.count);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "freeze_panes") {
              grid?.freezePanes(raw.sheet, raw.freeze_rows, raw.freeze_cols);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "unfreeze_panes") {
              grid?.unfreezePanes(raw.sheet);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "hide_rows") {
              grid?.hideRows(raw.sheet, raw.rows);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "show_rows") {
              grid?.showRows(raw.sheet, raw.rows);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "hide_columns") {
              grid?.hideColumns(raw.sheet, raw.columns);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
            if (raw.type === "show_columns") {
              grid?.showColumns(raw.sheet, raw.columns);
              dispatch({ type: "batch_append_mutation", tabId: ev.tab_id, batchId: ev.batch_id, mutation: raw });
              continue;
            }
          }
          dispatch({
            type: "set_status",
            tabId: ev.tab_id,
            phase: "writing",
            message: `Writing… (${tintedCellsByBatch.current[ev.batch_id]?.length ?? 0} cells)`,
          });

          // Report evaluated cell values back to the agent loop so the next
          // turn's tool_result carries the *computed* values (including
          // #VALUE! errors). Defer by one microtask + RAF so Univer's
          // formula engine has time to evaluate any formulas we just wrote.
          // AWAITED inside the queued task so the next tool_call in the
          // queue can't start until this one's report is in flight to Rust.
          await new Promise<void>((r) =>
            requestAnimationFrame(() => requestAnimationFrame(() => r())),
          );
          const evalCells: Array<{ sheet: string; row: number; col: number }> = [];
          for (const raw of result.mutations) {
            if (raw.type === "set_cell") {
              evalCells.push({ sheet: raw.address.sheet, row: raw.address.row, col: raw.address.col });
            } else if (raw.type === "set_format") {
              for (const c of raw.cells) evalCells.push({ sheet: raw.sheet, row: c.row, col: c.col });
            }
          }
          const live = gridRefs.current[ev.tab_id];
          const out = evalCells.map(({ sheet, row, col }) => {
            const c = live?.getCell(sheet, row, col);
            return {
              cell: a1Of(row, col),
              value: c?.value ?? null,
              formula: c?.formula ?? null,
            };
          });
          const content = JSON.stringify({ cells: out.slice(0, 200) });
          try {
            await reportToolResult(ev.tool_use_id, content);
          } catch (e) {
            console.warn("[agent] reportToolResult failed:", e);
          }
          }); // end enqueueToolCallTask
          return;
        }
        case "done": {
          dispatch({ type: "batch_finalize", tabId: ev.tab_id, batchId: ev.batch_id });
          dispatch({ type: "set_agent_running", tabId: ev.tab_id, running: false });
          const editedCells = tintedCellsByBatch.current[ev.batch_id]?.length ?? 0;
          dispatch({
            type: "set_status",
            tabId: ev.tab_id,
            phase: "done",
            // No edits this turn — the agent just chatted back. Don't bother
            // the user with "0 cells pending review" copy.
            message: editedCells > 0
              ? `Done · ${editedCells} cell${editedCells === 1 ? "" : "s"} pending review`
              : "Replied",
          });
          dispatch({
            type: "set_tokens",
            tabId: ev.tab_id,
            input: ev.input_tokens,
            output: ev.output_tokens,
            cacheRead: ev.cache_read_tokens,
            cacheCreation: ev.cache_creation_tokens,
          });
          // Persist the per-batch counts to the session row so the Usage
          // tab survives app restarts. Fire-and-forget — errors are
          // logged but don't block the UI flow.
          addSessionTokens(
            ev.tab_id,
            ev.input_tokens,
            ev.output_tokens,
            ev.cache_read_tokens,
            ev.cache_creation_tokens,
          ).catch((e) => console.warn("[session] add_tokens failed:", e));
          // Only mark dirty if the agent actually wrote something.
          if (editedCells > 0) {
            dispatch({ type: "mark_dirty", tabId: ev.tab_id, dirty: true });
            // Snapshot untitled drafts immediately after the agent finishes
            // writing — natural trigger, no contention with active tool calls
            // (we're past the last tool_result by definition here), and
            // guarantees the user's latest answer survives a quit before
            // the 30s background timer would fire.
            const finishedTab = findTab(workspaceRef.current, ev.tab_id);
            if (finishedTab) void snapshotUntitled(finishedTab);
          }

          // Persist what the agent produced for this turn. Read the full
          // streaming prose from the per-batch ref (mutated synchronously
          // in the text_delta handler) rather than via workspaceRef, which
          // can be a React commit behind for fast back-to-back deltas.
          const fullText = (streamingTextByBatch.current[ev.batch_id] ?? "").trim();
          delete streamingTextByBatch.current[ev.batch_id];
          (async () => {
            try {
              if (fullText) {
                dispatch({
                  type: "batch_set_agent_text",
                  tabId: ev.tab_id,
                  batchId: ev.batch_id,
                  agentText: fullText,
                });
                await appendMessage(ev.tab_id, "agent_text", { text: fullText });
              }
              const tabSnap = workspaceRef.current.tabs.find((t) => t.id === ev.tab_id);
              const batchSnap = tabSnap?.batches.find((b) => b.id === ev.batch_id);
              if (batchSnap) {
                await appendMessage(ev.tab_id, "agent_batch", {
                  batch: { ...batchSnap, agent_text: fullText || undefined, status: "pending" },
                });
              }
              // After the first batch on this tab, upgrade the heuristic
              // session name with an LLM-generated 3-5 word title. Only
              // overwrite if the user hasn't manually renamed (current name
              // still equals the heuristic of the first prompt). Best-effort
              // — failures keep the heuristic and don't surface to the user.
              if (tabSnap && batchSnap && tabSnap.batches.length === 1) {
                const heuristic = sessionNameFromPrompt(batchSnap.prompt);
                if (tabSnap.name === heuristic) {
                  try {
                    const title = await invoke<string>("generate_session_title", {
                      prompt: batchSnap.prompt,
                      justification: batchSnap.justification || null,
                    });
                    const trimmed = (title || "").trim();
                    const stillTab = workspaceRef.current.tabs.find((t) => t.id === ev.tab_id);
                    if (trimmed && stillTab && stillTab.name === heuristic) {
                      dispatch({ type: "set_name", tabId: ev.tab_id, name: trimmed });
                      upsertSession(ev.tab_id, trimmed, stillTab.path).catch((e) =>
                        console.warn("[session] upsert with llm title failed:", e),
                      );
                    }
                  } catch (e) {
                    console.warn("[session] generate_session_title failed:", e);
                  }
                }
              }
            } catch (e) {
              console.warn("[session] persist on done failed:", e);
            }
            // Auto-mode: accept the batch immediately if the user has it on,
            // and the batch actually has edits (no point auto-"accepting"
            // a pure chat-reply turn which has 0 mutations).
            if (autoApplyRef.current && editedCells > 0) {
              acceptBatchRef.current?.(ev.batch_id);
            }
          })();
          return;
        }
        case "error":
          dispatch({ type: "batch_finalize", tabId: ev.tab_id, batchId: ev.batch_id });
          dispatch({ type: "set_agent_running", tabId: ev.tab_id, running: false });
          dispatch({
            type: "set_status",
            tabId: ev.tab_id,
            phase: "error",
            message: ev.message.slice(0, 200),
          });
          toast({ title: "Agent error", description: ev.message, status: "error", duration: 6000 });
          return;
      }
    });
    return unsub;
  }, [toast]);

  // --- accept / reject ----------------------------------------------------

  const acceptBatch = useCallback(
    async (batchId: string) => {
      if (!activeTab) return;
      delete tintedCellsByBatch.current[batchId];
      delete oldValuesByBatch.current[batchId];

      dispatch({ type: "batch_accept", tabId: activeTab.id, batchId });
      dispatch({ type: "stream_text_clear", tabId: activeTab.id });

      const batch = findTab(workspace, activeTab.id)?.batches.find((b) => b.id === batchId);
      if (batch) {
        const accepted = { ...batch, status: "accepted" as const };
        try {
          await appendChangeBatch(activeTab.path, JSON.stringify(accepted));
        } catch (err) {
          toast({ title: "Change log write failed", description: String(err), status: "warning" });
        }
        // Record acceptance as a separate message so the session timeline
        // shows the full lifecycle (the prior `agent_batch` row captured the
        // mutations as pending; this row records the user's decision).
        appendMessage(activeTab.id, "agent_batch", { batch: accepted }).catch((e) =>
          console.warn("[session] persist accept failed:", e),
        );
      }
    },
    [activeTab, workspace, toast],
  );

  // Mirror acceptBatch into a ref so the agent-event handler (which is
  // declared earlier in this file and would otherwise close over an
  // undefined value) can call the latest version on auto-mode `done` events.
  const acceptBatchRef = useRef<((batchId: string) => void) | null>(null);
  useEffect(() => { acceptBatchRef.current = acceptBatch; }, [acceptBatch]);

  const rejectBatch = useCallback(
    async (batchId: string) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const olds = oldValuesByBatch.current[batchId] ?? [];
      for (const o of olds) {
        grid?.setCell(o.sheet, o.row, o.col, o.oldFormula ?? o.oldValue);
      }
      // Restore non-value mutations in this batch.
      const liveBatch = activeTab.batches.find((b) => b.id === batchId);
      if (liveBatch) {
        for (const m of liveBatch.mutations) {
          if (m.type === "set_format") {
            for (const o of m.old_format) {
              if (o.format) grid?.setCellFormat(m.sheet, o.row, o.col, o.format);
            }
          } else if (m.type === "set_column_width") {
            for (const o of m.old_widths) {
              if (o.width != null) grid?.setColumnWidth(m.sheet, o.col, o.width);
            }
          } else if (m.type === "set_row_height") {
            for (const o of m.old_heights) {
              if (o.height != null) grid?.setRowHeight(m.sheet, o.row, o.height);
            }
          } else if (m.type === "merge_cells") {
            grid?.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
          } else if (m.type === "unmerge_cells") {
            grid?.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
          }
        }
      }
      delete tintedCellsByBatch.current[batchId];
      delete oldValuesByBatch.current[batchId];

      dispatch({ type: "batch_reject", tabId: activeTab.id, batchId });
      dispatch({ type: "stream_text_clear", tabId: activeTab.id });

      const batch = findTab(workspace, activeTab.id)?.batches.find((b) => b.id === batchId);
      if (batch) {
        const rejected = { ...batch, status: "rejected" as const };
        try {
          await appendChangeBatch(activeTab.path, JSON.stringify(rejected));
        } catch {}
        appendMessage(activeTab.id, "agent_batch", { batch: rejected }).catch((e) =>
          console.warn("[session] persist reject failed:", e),
        );
      }
    },
    [activeTab, workspace],
  );

  /**
   * Undo an already-accepted batch: walk every mutation in reverse and
   * restore the captured old_* state. Sets the batch's status to "rejected"
   * (which visually flips the badge) and appends a new log entry so the
   * change log shows the full lifecycle: accepted → then undone.
   *
   * v1 limitation: undoing an OLD batch when newer batches edited the same
   * cells will produce semi-undefined results — the newer batches' "old"
   * snapshot was taken AFTER the older batch's edits. The right fix is a
   * checkpoint stack; for now the UI just allows it and we trust the user
   * to undo in roughly reverse order (or accept the messiness).
   */
  const undoBatch = useCallback(
    async (batchId: string) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const batch = activeTab.batches.find((b) => b.id === batchId);
      if (!batch || batch.status !== "accepted") return;

      // Restore set_cell mutations from their old_value / old_formula.
      // Iterate reversed so writes-on-top-of-writes within the same batch
      // wind back correctly.
      const reversed = [...batch.mutations].reverse();
      for (const m of reversed) {
        if (m.type === "set_cell") {
          grid?.setCell(m.address.sheet, m.address.row, m.address.col, m.old_formula ?? m.old_value);
        } else if (m.type === "set_format") {
          for (const o of m.old_format) {
            if (o.format) grid?.setCellFormat(m.sheet, o.row, o.col, o.format);
          }
        } else if (m.type === "set_column_width") {
          for (const o of m.old_widths) {
            if (o.width != null) grid?.setColumnWidth(m.sheet, o.col, o.width);
          }
        } else if (m.type === "set_row_height") {
          for (const o of m.old_heights) {
            if (o.height != null) grid?.setRowHeight(m.sheet, o.row, o.height);
          }
        } else if (m.type === "merge_cells") {
          grid?.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        } else if (m.type === "unmerge_cells") {
          grid?.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        }
      }

      dispatch({ type: "batch_reject", tabId: activeTab.id, batchId });
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });

      const undone = { ...batch, status: "rejected" as const };
      try {
        await appendChangeBatch(activeTab.path, JSON.stringify(undone));
      } catch (err) {
        toast({ title: "Change log write failed", description: String(err), status: "warning" });
      }
      appendMessage(activeTab.id, "agent_batch", { batch: undone }).catch((e) =>
        console.warn("[session] persist undo failed:", e),
      );

      toast({ title: "Undone", description: `Reverted ${batch.mutations.length} edit${batch.mutations.length === 1 ? "" : "s"}.`, status: "info", duration: 2000 });
    },
    [activeTab, toast],
  );

  /**
   * Inverse of undoBatch — re-apply every mutation's NEW state and flip
   * the batch back to "accepted". Works for any rejected batch (whether
   * the user initially Rejected a pending one or Undid an accepted one)
   * because the mutations still carry the agent's intended new values.
   */
  const redoBatch = useCallback(
    async (batchId: string) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const batch = activeTab.batches.find((b) => b.id === batchId);
      if (!batch || batch.status !== "rejected") return;

      for (const m of batch.mutations) {
        if (m.type === "set_cell") {
          grid?.setCell(m.address.sheet, m.address.row, m.address.col, m.new_formula ?? m.new_value);
        } else if (m.type === "set_format") {
          for (const c of m.cells) grid?.setCellFormat(m.sheet, c.row, c.col, m.new_format);
        } else if (m.type === "set_column_width") {
          for (const col of m.columns) grid?.setColumnWidth(m.sheet, col, m.new_width);
        } else if (m.type === "set_row_height") {
          for (const row of m.rows) grid?.setRowHeight(m.sheet, row, m.new_height);
        } else if (m.type === "merge_cells") {
          grid?.mergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        } else if (m.type === "unmerge_cells") {
          grid?.unmergeCells(m.sheet, m.start_row, m.start_col, m.end_row, m.end_col);
        }
      }

      dispatch({ type: "batch_accept", tabId: activeTab.id, batchId });
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });

      const restored = { ...batch, status: "accepted" as const };
      try {
        await appendChangeBatch(activeTab.path, JSON.stringify(restored));
      } catch (err) {
        toast({ title: "Change log write failed", description: String(err), status: "warning" });
      }
      appendMessage(activeTab.id, "agent_batch", { batch: restored }).catch((e) =>
        console.warn("[session] persist redo failed:", e),
      );

      toast({ title: "Redone", description: `Re-applied ${batch.mutations.length} edit${batch.mutations.length === 1 ? "" : "s"}.`, status: "info", duration: 2000 });
    },
    [activeTab, toast],
  );

  // --- keyboard shortcuts -------------------------------------------------

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toLowerCase().includes("mac");
      const mod = isMac ? e.metaKey : e.ctrlKey;
      if (!mod) return;
      const key = e.key.toLowerCase();
      if (key === "t") {
        e.preventDefault();
        openWorkbookDialog();
      } else if (key === "n" && !e.shiftKey) {
        e.preventDefault();
        createBlankWorkbook();
      } else if (key === "w") {
        if (activeTab) {
          e.preventDefault();
          closeTab(activeTab.id);
        }
      } else if (key === "s") {
        e.preventDefault();
        if (e.shiftKey) saveActiveAs();
        else saveActive();
      } else if (e.key === "." && activeTab?.agentRunning) {
        e.preventDefault();
        const running = activeTab.batches.find((b) => b.status === "streaming");
        if (running) stopAgentTurn(running.id).catch(console.warn);
      } else if (key === "z" && activeTab) {
        // ⌘Z = undo the most recent accepted batch.
        // ⌘⇧Z = redo (re-apply) the most recent rejected/undone batch.
        // Skip when a text field has focus — preserve native text undo there.
        const target = e.target as HTMLElement | null;
        if (target && (target.tagName === "TEXTAREA" || target.tagName === "INPUT")) return;
        if (e.shiftKey) {
          const lastRejected = [...activeTab.batches].reverse().find((b) => b.status === "rejected");
          if (lastRejected) {
            e.preventDefault();
            redoBatch(lastRejected.id);
          }
        } else {
          const lastAccepted = [...activeTab.batches].reverse().find((b) => b.status === "accepted");
          if (lastAccepted) {
            e.preventDefault();
            undoBatch(lastAccepted.id);
          }
        }
      } else if (/^[1-9]$/.test(key)) {
        const idx = parseInt(key, 10) - 1;
        const tab = workspace.tabs[idx];
        if (tab) {
          e.preventDefault();
          dispatch({ type: "activate", tabId: tab.id });
        }
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeTab, workspace.tabs, openWorkbookDialog, createBlankWorkbook, closeTab, saveActive, saveActiveAs, undoBatch, redoBatch]);

  // --- render -------------------------------------------------------------

  const batches = activeTab?.batches ?? [];

  const statusInfo = useMemo(() => {
    if (!activeTab) return { phase: "idle" as const, message: "Open a workbook to start (⌘T)" };
    return { phase: activeTab.statusPhase, message: activeTab.statusMessage };
  }, [activeTab]);

  return (
    <Page
      $sidebarOpen={sidebarOpen}
      $chatOpen={chatOpen}
      $chatWidth={chatWidth}
      $hasTabs={workspace.tabs.length > 0}
    >
      <SidebarArea>
        {sidebarOpen && (
          <SessionSidebar
            sessions={recentSessions}
            activePath={activeTab?.path ?? null}
            onOpen={(s) => resumeSession(s)}
            onNewBlank={createBlankWorkbook}
            onOpenExisting={openWorkbookDialog}
            onArchive={async (id) => {
              await archiveSession(id);
              refreshSessions();
            }}
            onDelete={async (id) => {
              await deleteSession(id);
              refreshSessions();
              // Also close the matching tab if open — its DB rows are gone now.
              const open = workspace.tabs.find((t) => t.id === id);
              if (open) dispatch({ type: "close", tabId: id });
            }}
          />
        )}
      </SidebarArea>
      <Header>
        <HeaderButton
          onClick={() => setSidebarOpen((v) => !v)}
          title={sidebarOpen ? "Hide sessions sidebar" : "Show sessions sidebar"}
          style={{ padding: "4px 8px" }}
        >
          {sidebarOpen ? <PanelLeftClose size={13} /> : <PanelLeftOpen size={13} />}
        </HeaderButton>
        <span style={{ fontWeight: 600, color: "#e4e4e4", letterSpacing: 0.2 }}>
          GridPath
        </span>
        {activeTab && (
          <>
            <span style={{ color: "#3a3a3a" }}>/</span>
            <span style={{ color: activeTab.name ? "#d4d4d4" : "#6f6f6f", fontStyle: activeTab.name ? "normal" : "italic" }}>
              {activeTab.name || "Untitled session"}
            </span>
            <span style={{ color: "#666", fontSize: 11 }}>· {activeTab.filename}</span>
          </>
        )}
        <span style={{ marginLeft: "auto", color: "#7c7c7c", fontSize: 11 }}>
          {workspace.tabs.length === 0
            ? "no session"
            : `${workspace.tabs.length} session${workspace.tabs.length === 1 ? "" : "s"}${workspace.tabs.filter((t) => t.agentRunning).length > 0 ? ` · ${workspace.tabs.filter((t) => t.agentRunning).length} running` : ""}`}
        </span>
        <div style={{ display: "flex", gap: 8 }}>
          <HeaderButton onClick={saveActive} disabled={!activeTab}>
            <Save size={12} /> Save (⌘S)
          </HeaderButton>
          <HeaderButton onClick={saveActiveAs} disabled={!activeTab} title="Save a copy to a new location (⌘⇧S)">
            Save As…
          </HeaderButton>
          <HeaderButton
            onClick={() => setChatOpen((v) => !v)}
            title={chatOpen ? "Hide chat panel" : "Show chat panel"}
            style={{ padding: "4px 8px" }}
          >
            <PanelRightClose size={13} />
          </HeaderButton>
          <HeaderButton onClick={() => setSettingsOpen(true)} title="API keys, Claude connection">
            <SettingsIcon size={12} /> Settings
          </HeaderButton>
        </div>
      </Header>

      <TabsArea>
        {workspace.tabs.length > 0 && <TabBar
          tabs={workspace.tabs}
          activeTabId={workspace.activeTabId}
          onActivate={(id) => dispatch({ type: "activate", tabId: id })}
          onClose={closeTab}
          onRename={(id, name) => {
            dispatch({ type: "set_name", tabId: id, name });
            renameSessionDb(id, name).catch((e) =>
              console.warn("[session] rename failed:", e),
            );
          }}
        />}
      </TabsArea>

      <GridArea>
        {workspace.tabs.length === 0 && (
          <EmptyState>
            <div style={{ fontSize: 15, color: "#9b9b9b" }}>No workbook open</div>
            <EmptyButtons>
              <EmptyPrimaryBtn onClick={createBlankWorkbook}>
                <FilePlus size={14} /> New blank (⌘N)
              </EmptyPrimaryBtn>
              <EmptySecondaryBtn onClick={openWorkbookDialog}>
                <FolderOpen size={14} /> Open existing (⌘T)
              </EmptySecondaryBtn>
            </EmptyButtons>
            <div style={{ fontSize: 11, color: "#555" }}>
              or pick one from <strong style={{ color: "#7c7c7c" }}>Recent sessions</strong> on the left
            </div>
          </EmptyState>
        )}
        {workspace.tabs.map((t) => (
          <GridLayer key={t.id} $visible={t.id === workspace.activeTabId}>
            <UniverGrid
              ref={(h) => {
                if (h) gridRefs.current[t.id] = h;
                else delete gridRefs.current[t.id];
              }}
              workbookPath={t.path}
            />
          </GridLayer>
        ))}
      </GridArea>

      <ChatArea>
        {chatOpen && <ChatResizer onMouseDown={startChatResize} title="Drag to resize" />}
        <ChatPanel
          tab={activeTab}
          prompt={promptForActive}
          onPromptChange={(v) => activeTab && setPromptByTab((s) => ({ ...s, [activeTab.id]: v }))}
          onSubmit={submitPrompt}
          onStop={() => {
            const running = activeTab?.batches.find((b) => b.status === "streaming");
            if (running) stopAgentTurn(running.id).catch(console.warn);
          }}
          onAccept={acceptBatch}
          onReject={rejectBatch}
          onUndo={undoBatch}
          onRedo={redoBatch}
          selectionLabel={focusDismissedFor === liveSelection ? null : liveSelection}
          onDismissSelection={() => setFocusDismissedFor(liveSelection)}
          autoApply={autoApply}
          onSetAutoApply={(next) => setAutoApply(next)}
        />
      </ChatArea>

      <StatusArea>
        <StatusBar
          workbookPath={activeTab?.path ?? null}
          dirty={activeTab?.dirty ?? false}
          phase={statusInfo.phase}
          message={statusInfo.message}
          lastSavedAt={activeTab?.lastSavedAt ?? null}
          inputTokens={activeTab?.inputTokens}
          outputTokens={activeTab?.outputTokens}
        />
      </StatusArea>

      <SettingsModal
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        tabs={workspace.tabs}
      />
      <UpdateNotification />
      <ExitGuardModal
        open={!!exitGuard}
        dirtyTabs={exitGuard?.dirtyTabs ?? []}
        saving={exitGuard?.saving ?? false}
        onCancel={() => setExitGuard(null)}
        onDiscard={async () => {
          // destroy() bypasses our onCloseRequested handler — otherwise
          // closing programmatically would re-prompt forever.
          setExitGuard(null);
          try { await getCurrentWindow().destroy(); } catch (e) { console.warn("[exit-guard] destroy failed:", e); }
        }}
        onSave={async () => {
          if (!exitGuard) return;
          setExitGuard({ ...exitGuard, saving: true });
          // Save sequentially so each untitled tab's Save-As dialog can
          // resolve without overlapping. If any save returns false (user
          // cancelled the dialog), stay in the modal so they can retry
          // or pick Discard.
          let allSaved = true;
          for (const t of exitGuard.dirtyTabs) {
            // Re-read the tab from the latest workspace in case prior
            // iterations renamed it (untitled → real path).
            const latest = workspaceRef.current.tabs.find((x) => x.id === t.id) ?? t;
            const ok = await saveTabBytes(latest, false, true);
            if (!ok) { allSaved = false; break; }
          }
          if (allSaved) {
            setExitGuard(null);
            try { await getCurrentWindow().destroy(); } catch (e) { console.warn("[exit-guard] destroy failed:", e); }
          } else {
            setExitGuard((g) => g ? { ...g, saving: false } : null);
          }
        }}
      />
    </Page>
  );
};

/**
 * Build the agent-mutation mirror we ship to UniverGrid.exportBytes so the
 * ExcelJS workbook gets the agent's format / width / height / merge changes
 * patched into it before we writeBuffer. Walks every ACCEPTED batch — the
 * batches themselves carry old/new state per mutation, so we don't have to
 * re-read from Univer.
 *
 * If a cell got multiple format edits across batches, the last one wins
 * (later accepted batch overwrites earlier). Same for column widths etc.
 */
function buildSaveMirror(batches: any[]): SaveMirror {
  const cellFormats: SaveMirror["cellFormats"] = [];
  const columnWidths: SaveMirror["columnWidths"] = [];
  const rowHeights: SaveMirror["rowHeights"] = [];
  const merges: SaveMirror["merges"] = [];
  const sheetOps: SaveMirror["sheetOps"] = [];
  const clears: SaveMirror["clears"] = [];
  const rowColOps: SaveMirror["rowColOps"] = [];
  const freezePanes: SaveMirror["freezePanes"] = [];
  const visibility: SaveMirror["visibility"] = [];
  for (const b of batches) {
    if (b.status !== "accepted") continue;
    for (const m of b.mutations ?? []) {
      if (m.type === "set_format") {
        // Save mirror keeps `background` as a sibling field on each cell
        // entry (the ExcelJS side reads it from there in applyStyleMirror).
        // Hoist background_color out of the format object so the saved
        // xlsx actually carries the fill, not just the in-app display.
        const bg = m.new_format?.background_color ?? null;
        for (const c of m.cells ?? []) {
          cellFormats!.push({
            sheet: m.sheet,
            row: c.row,
            col: c.col,
            format: m.new_format,
            background: bg,
          });
        }
      } else if (m.type === "set_column_width") {
        for (const col of m.columns ?? []) {
          columnWidths!.push({ sheet: m.sheet, col, widthPx: m.new_width });
        }
      } else if (m.type === "set_row_height") {
        for (const row of m.rows ?? []) {
          rowHeights!.push({ sheet: m.sheet, row, heightPx: m.new_height });
        }
      } else if (m.type === "merge_cells") {
        merges!.push({ sheet: m.sheet, range: m.range, merge: true });
      } else if (m.type === "unmerge_cells") {
        merges!.push({ sheet: m.sheet, range: m.range, merge: false });
      } else if (m.type === "create_sheet") {
        sheetOps!.push({ kind: "create", name: m.name, tabColor: m.tab_color });
      } else if (m.type === "delete_sheet") {
        sheetOps!.push({ kind: "delete", name: m.name });
      } else if (m.type === "rename_sheet") {
        sheetOps!.push({ kind: "rename", oldName: m.old_name, newName: m.new_name });
      } else if (m.type === "clear_range") {
        for (const c of m.cells ?? []) clears!.push({ sheet: m.sheet, row: c.row, col: c.col });
      } else if (m.type === "insert_rows") {
        rowColOps!.push({ kind: "insertRows", sheet: m.sheet, before: m.before, count: m.count });
      } else if (m.type === "delete_rows") {
        rowColOps!.push({ kind: "deleteRows", sheet: m.sheet, start: m.start, count: m.count });
      } else if (m.type === "insert_columns") {
        rowColOps!.push({ kind: "insertColumns", sheet: m.sheet, before: m.before, count: m.count });
      } else if (m.type === "delete_columns") {
        rowColOps!.push({ kind: "deleteColumns", sheet: m.sheet, start: m.start, count: m.count });
      } else if (m.type === "freeze_panes") {
        freezePanes!.push({ sheet: m.sheet, freezeRows: m.freeze_rows, freezeCols: m.freeze_cols });
      } else if (m.type === "unfreeze_panes") {
        freezePanes!.push({ sheet: m.sheet, freezeRows: 0, freezeCols: 0 });
      } else if (m.type === "hide_rows") {
        visibility!.push({ kind: "hideRows", sheet: m.sheet, rows: m.rows });
      } else if (m.type === "show_rows") {
        visibility!.push({ kind: "showRows", sheet: m.sheet, rows: m.rows });
      } else if (m.type === "hide_columns") {
        visibility!.push({ kind: "hideColumns", sheet: m.sheet, columns: m.columns });
      } else if (m.type === "show_columns") {
        visibility!.push({ kind: "showColumns", sheet: m.sheet, columns: m.columns });
      }
    }
  }
  return { cellFormats, columnWidths, rowHeights, merges, sheetOps, clears, rowColOps, freezePanes, visibility };
}

function a1Of(row: number, col: number): string {
  let n = col;
  let letters = "";
  do {
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return `${letters}${row + 1}`;
}
