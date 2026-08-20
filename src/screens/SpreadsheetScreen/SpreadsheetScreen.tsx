import React, { useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
import styled from "styled-components";
import { useToast } from "@chakra-ui/react";
import { Settings as SettingsIcon, Save, PanelLeftOpen, PanelLeftClose, PanelRightClose, FolderOpen, FilePlus, Maximize2, X } from "lucide-react";
import { open, save } from "@tauri-apps/plugin-dialog";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { invoke } from "@tauri-apps/api/core";
import { v4 as uuidv4 } from "uuid";

import { UniverGrid, type UniverGridHandle, type SaveMirror, type BordersShape, type ManualSheetOp } from "./components/UniverGrid";
import { FormatToolbar, type InsertAction } from "./components/FormatToolbar";
import { StatusBar, type SelectionStats } from "./components/StatusBar";
import { UpdateNotification } from "./components/UpdateNotification";
import { ChatPanel } from "./components/ChatPanel";
import { TabBar } from "./components/TabBar";
import { SessionSidebar } from "./components/SessionSidebar";
import { SettingsModal } from "./components/SettingsModal";
import { ExitGuardModal } from "./components/ExitGuardModal";
import { FidelitySaveModal } from "./components/FidelitySaveModal";
import { ReviewModal } from "./components/ReviewModal";
import { appendManualSheetOp } from "./manualSheetOps";
import { colLetters } from "./agent/reviewGroups";
import { getSettingValue, setSettingValue, getModel, getEffort, SETTING_KEYS } from "./settingsApi";
import { getEvalConfig, evalFinish, type EvalConfig, type EvalMeta } from "./agent/evalDriver";
import {
  readWorkbookBytes,
  writeWorkbookBytes,
  saveWorkbookPatched,
  saveWorkbookPatchedAs,
  readUntitledSnapshot,
  writeUntitledSnapshot,
} from "./workbookIo";
import { describePatchFallback } from "./surgicalPatch";
import { auditExportLoss, describeFidelityRisks, describeMissingParts, gateFallbackSave } from "./fidelityScan";
import { auditDefinedNameLoss, repairDefinedNames } from "./definedNamePreserve";
import { sanitizeExportedPackage } from "./exportSanitize";

/**
 * Surgical save is the default; `localStorage.setItem("gridpath.surgicalSave",
 * "off")` reverts to the ExcelJS full-export path as an escape hatch.
 */
function surgicalSaveEnabled(): boolean {
  try {
    return localStorage.getItem("gridpath.surgicalSave") !== "off";
  } catch {
    return true;
  }
}

/**
 * Await `frames` animation frames, but never longer than `ms`. WKWebView
 * suspends requestAnimationFrame while the window is hidden or fully
 * occluded, so a bare rAF await parks its caller until the window is next
 * visible — observed as a 49s "settle" on a 1-cell write, with the whole
 * per-tab tool queue stalled behind it while the Rust loop's readback
 * timeout substituted fake "ok" results. The timer keeps hidden-window
 * batches flowing (throttled timers still fire); visible windows resolve on
 * the real frame(s) as before.
 */
function nextFrame(frames = 1, ms = 150): Promise<void> {
  return new Promise<void>((resolve) => {
    const timer = setTimeout(resolve, ms);
    let left = frames;
    const tick = () =>
      requestAnimationFrame(() => {
        left -= 1;
        if (left <= 0) {
          clearTimeout(timer);
          resolve();
        } else {
          tick();
        }
      });
    tick();
  });
}
import type { ChangeBatch, UniverMutation, FormatMutation, CellFormat } from "./types";
import { ironcalcShadowLoad } from "./ironcalcShadow";
import {
  ironcalcEngineOnBlank,
  ironcalcEngineOnClose,
  ironcalcEngineOnOpen,
  ironcalcEngineOnRestore,
  ironcalcEngineSetActiveTab,
} from "./ironcalcEngine";
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
  reportToolStarted,
  subscribeAgentEvents,
  type AgentEvent,
} from "./agent/agentClient";
import { interpretToolCall } from "./agent/toolToMutation";
import { layoutStageBlock } from "./agent/stageLayout";
import { expandA1Range } from "./agent/toolToMutation";
import { expandCopy } from "./agent/copyRange";
import { executeSheetScript } from "./agent/scriptRunner";
import { captureWorkbookContext } from "./agent/captureContext";
import { fileSavedIfDivergent, formatCalcHealthLine } from "./agent/calcHealth";
import {
  describeWorkbookPayload,
  findRowsInIndex,
  firstLabelIn,
  getWorkbookIndex,
  isLabelText,
} from "./agent/workbookIndex";
import {
  reverseMutation,
  applyMutationForward,
  validateBatchLayout,
  isLiteralOnlyBatch,
} from "./agent/batchMutations";
import {
  loadReferenceWorkbook,
  captureReferenceContext,
  readReferenceRange,
  evictReference,
  getCachedReference,
  referenceLabelFromPath,
  MAX_REFERENCES,
  type ParsedReference,
} from "./agent/referenceContext";
import { buildPriorBatchesContext } from "./agent/priorContext";
import { buildFocusContext } from "./agent/selectionContext";
import {
  upsertSession,
  renameSession as renameSessionDb,
  setSessionReferences,
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
  $hasActiveTab: boolean;
}>`
  display: grid;
  grid-template-rows: 38px ${(p) => (p.$hasTabs ? "52px" : "0")} ${(p) => (p.$hasActiveTab ? "minmax(34px, auto)" : "0")} 1fr 24px;
  /* The middle (sheet/ribbon) column must be minmax(0, 1fr), NOT a bare 1fr:
     a plain 1fr track has an implicit min-width:auto, so the wide formatting
     toolbar would force the column to its min-content width, overflow the
     viewport, and squeeze the fixed-width chat panel into a thin strip. */
  grid-template-columns:
    ${(p) => (p.$sidebarOpen ? "260px" : "0")}
    minmax(0, 1fr)
    ${(p) => (p.$chatOpen ? `${p.$chatWidth}px` : "0")};
  grid-template-areas:
    "header  header  header"
    "sidebar tabs    chat"
    "sidebar ribbon  chat"
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

const RibbonArea = styled.div`
  grid-area: ribbon;
  /* The toolbar's dropdown menus open downward, past this row, into the grid
     region. overflow must stay visible (otherwise the menus are clipped to a
     sliver and look broken), and a stacking context above the grid keeps them
     on top of the Univer canvas. min-width:0 lets the column shrink so the
     toolbar wraps instead of pushing the layout wider than the viewport. */
  position: relative;
  z-index: 30;
  min-width: 0;
  border-bottom: 1px solid #cdd1d6;
`;

const GridArea = styled.div`
  grid-area: grid;
  position: relative;
  overflow: hidden;
  min-width: 0;
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

// Opaque cover over the grid while a workbook parses/builds — without it a
// big model shows a blank white grid for seconds, then "pops" into existence.
const GridLoadOverlay = styled.div`
  position: absolute;
  inset: 0;
  z-index: 10;
  background: #fafbfc;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
`;

const LoadSpinner = styled.div`
  width: 26px;
  height: 26px;
  border: 3px solid #d8dde3;
  border-top-color: #3363AD;
  border-radius: 50%;
  animation: gridload-spin 0.8s linear infinite;
  @keyframes gridload-spin { to { transform: rotate(360deg); } }
`;

const LoadFilename = styled.div`
  font-size: 13px;
  font-weight: 600;
  color: #3a4148;
`;

const LoadStage = styled.div`
  font-size: 12px;
  color: #8a939c;
`;

const StatusArea = styled.div`
  grid-area: status;
`;

/** Floating pill shown while the review modal is minimized for a jump-to-cell. */
const ReviewChip = styled.button`
  position: fixed;
  bottom: 46px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 150;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  padding: 7px 14px;
  border-radius: 16px;
  border: 1px solid rgba(51, 99, 173, 0.6);
  background: #1d2735;
  color: #93c5fd;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  box-shadow: 0 6px 24px rgba(0, 0, 0, 0.45);
  &:hover { background: #223042; }
`;

const ReviewChipClose = styled.span`
  display: inline-flex;
  align-items: center;
  border-radius: 50%;
  padding: 2px;
  color: #7ba6d8;
  &:hover { background: rgba(51, 99, 173, 0.35); color: #cfe1f8; }
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
  // Excel-style live selection stats (Sum / Average / Count) for the status bar.
  const [selectionStats, setSelectionStats] = React.useState<SelectionStats | null>(null);
  const statsSelKeyRef = useRef<string>("");
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

  /**
   * Load errors for attached reference workbooks, keyed by path (e.g. the
   * file was deleted since it was attached). Errored references render as
   * red chips and are skipped when building the agent's context.
   */
  const [referenceErrors, setReferenceErrors] = React.useState<Record<string, string>>({});

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
  // Per-tab load overlay: stage label while a workbook parses/builds, null
  // when the grid is live.
  const [gridLoading, setGridLoading] = React.useState<Record<string, string | null>>({});
  const setLoadStage = useCallback((tabId: string, stage: string | null) => {
    setGridLoading((s) => {
      if (stage === null) {
        if (!(tabId in s)) return s;
        const { [tabId]: _drop, ...rest } = s;
        return rest;
      }
      return { ...s, [tabId]: stage };
    });
  }, []);

  const [exitGuard, setExitGuard] = React.useState<{
    dirtyTabs: WorkbookTab[];
    saving: boolean;
  } | null>(null);

  // Reduced-fidelity Save As requires explicit consent: the modal names
  // what the copy will lose and awaits the user's choice. Promise-backed so
  // saveTabBytes can simply `await` the answer mid-flow.
  const [fidelityPrompt, setFidelityPrompt] = React.useState<{
    filename: string;
    note: string | null;
    losses: string;
  } | null>(null);
  const fidelityResolveRef = useRef<((ok: boolean) => void) | null>(null);
  const confirmReducedFidelitySave = useCallback(
    (filename: string, note: string | null, losses: string) =>
      new Promise<boolean>((resolve) => {
        fidelityResolveRef.current = resolve;
        setFidelityPrompt({ filename, note, losses });
      }),
    [],
  );
  const settleFidelityPrompt = useCallback((ok: boolean) => {
    fidelityResolveRef.current?.(ok);
    fidelityResolveRef.current = null;
    setFidelityPrompt(null);
  }, []);
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

  // UniverGrid is wrapped in React.memo so a re-render of this screen (e.g. a
  // keystroke in a chat/AI input elsewhere on screen) doesn't re-render every
  // mounted tab's grid. That only holds if `ref` / `onUserEdit` keep the SAME
  // function identity across renders — an inline arrow function per tab would
  // recreate on every render and defeat the memo via a "changed prop" on
  // every comparison. Since we render ALL tabs via .map() (not just one
  // active tab), a plain useCallback keyed on a single id doesn't fit; cache
  // one stable callback per tabId instead, built lazily and reused for the
  // life of the tab. Cleared in closeTab so it doesn't grow across a session
  // of many opened/closed tabs.
  const gridRefCallbacks = useRef<Map<string, (h: UniverGridHandle | null) => void>>(new Map());
  const getGridRefCallback = (tabId: string) => {
    let cb = gridRefCallbacks.current.get(tabId);
    if (!cb) {
      cb = (h) => {
        if (h) gridRefs.current[tabId] = h;
        else delete gridRefs.current[tabId];
      };
      gridRefCallbacks.current.set(tabId, cb);
    }
    return cb;
  };
  const userEditCallbacks = useRef<Map<string, () => void>>(new Map());
  const getUserEditCallback = (tabId: string) => {
    let cb = userEditCallbacks.current.get(tabId);
    if (!cb) {
      cb = () => handleUserEdit(tabId);
      userEditCallbacks.current.set(tabId, cb);
    }
    return cb;
  };
  const manualSheetOpCallbacks = useRef<Map<string, (op: ManualSheetOp) => void>>(new Map());
  const getManualSheetOpCallback = (tabId: string) => {
    let cb = manualSheetOpCallbacks.current.get(tabId);
    if (!cb) {
      cb = (op) => handleManualSheetOp(tabId, op);
      manualSheetOpCallbacks.current.set(tabId, cb);
    }
    return cb;
  };

  // Manual formatting toolbar edits, keyed by tabId. These never go through
  // the agent batch/accept-reject flow — the user clicked a button — so we
  // accumulate them here and merge into the SaveMirror at save time so they
  // round-trip to xlsx. (In-session only; the saved file is the durable record.)
  const manualFormatsRef = useRef<Record<string, NonNullable<SaveMirror["cellFormats"]>>>({});
  const manualBordersRef = useRef<Record<string, NonNullable<SaveMirror["cellBorders"]>>>({});
  /**
   * Sheet add/delete/rename the user performed through Univer's own sheet-tab
   * UI. Without these the surgical patcher meets a live sheet it can't
   * explain and every save of the session falls back to the lossy export.
   */
  const manualSheetOpsRef = useRef<Record<string, NonNullable<SaveMirror["sheetOps"]>>>({});
  // Manual (toolbar) merge and row/column structure edits, mirrored to xlsx on
  // save the same way manual formats/borders are.
  const manualMergesRef = useRef<Record<string, NonNullable<SaveMirror["merges"]>>>({});
  const manualRowColRef = useRef<Record<string, NonNullable<SaveMirror["rowColOps"]>>>({});
  // Manual (toolbar) freeze-pane state per tab — last write wins, folded into
  // the save mirror so the freeze survives round-trip to xlsx.
  const manualFreezeRef = useRef<Record<string, NonNullable<SaveMirror["freezePanes"]>[number]>>({});
  // Manual (toolbar) AutoFilter range per tab — folded into the save mirror so
  // the filter range round-trips to xlsx.
  const manualFilterRef = useRef<Record<string, NonNullable<SaveMirror["autoFilters"]>[number]>>({});
  // Row/column visibility changes from the outline (grouping) level buttons,
  // folded into the save mirror like the other manual edits.
  const manualVisibilityRef = useRef<Record<string, NonNullable<SaveMirror["visibility"]>>>({});
  // Outline (grouping) depth of the active sheet — drives the "1 2 3" toolbar
  // buttons; null hides them entirely.
  const [outlineLevels, setOutlineLevels] = useState<{ cols: number; rows: number } | null>(null);
  const outlineKeyRef = useRef("");

  // Format Painter: when armed, holds the copied format + the selection key it
  // was copied from. The selection poll applies it to the next distinct range.
  const formatPainterRef = useRef<{ tabId: string; format: CellFormat; sourceKey: string } | null>(null);
  const [painterArmed, setPainterArmed] = React.useState(false);

  // Per-tab flag: has the user made a direct grid edit (via the formatting
  // toolbar) since the last agent batch? When true, ⌘Z falls through to Univer's
  // native undo (which reverts the user's edit) instead of hijacking it to undo
  // the whole last agent batch — see the keyboard handler. Reset whenever a new
  // agent turn starts or a batch is accepted/redone.
  const userEditedSinceBatchRef = useRef<Record<string, boolean>>({});

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
  // Same synchronous-accumulator trick for the agent's reasoning/plan stream,
  // so the `done` handler can persist the full plan onto the batch without
  // racing React's commit. See streamingTextByBatch above.
  const streamingReasoningByBatch = useRef<Record<string, string>>({});
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

  // Background tabs are hidden with visibility:hidden, which hides pixels
  // but NOT work: Univer's rAF render loop keeps repainting every mutation
  // on grids nobody can see. Freeze the render unit of every non-active
  // tab (official activate/deactivate API — data model, command pipeline
  // and IronCalc mirror keep running) so agent writes to background tabs
  // skip all layout/paint. The active tab's first frame after switching
  // repaints whatever changed while it was frozen.
  useEffect(() => {
    for (const t of workspace.tabs) {
      gridRefs.current[t.id]?.setRenderActive?.(t.id === workspace.activeTabId);
    }
  }, [workspace.activeTabId, workspace.tabs]);

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
        setSelectionStats(null);
        statsSelKeyRef.current = "";
        return;
      }

      // Excel-style Sum/Average/Count for the selection. Only recompute when
      // the selection rectangle changes (keyed) so the 500ms poll doesn't
      // re-scan cells every tick; cap the scan so a huge selection can't jank.
      const selKey = `${sel.sheet}:${sel.startRow},${sel.startCol},${sel.endRow},${sel.endCol}`;

      // Outline (grouping) buttons follow the active sheet. Keyed so the
      // poll only touches state when the sheet actually changes.
      if (outlineKeyRef.current !== `${activeTab.id}:${sel.sheet}`) {
        outlineKeyRef.current = `${activeTab.id}:${sel.sheet}`;
        const summary = gridRefs.current[activeTab.id]?.getOutlineSummary?.(sel.sheet) ?? null;
        setOutlineLevels(summary && (summary.cols > 0 || summary.rows > 0) ? summary : null);
      }

      // Format Painter: once armed, apply the copied format to the next
      // distinct selection, then disarm (Excel's single-use painter).
      const painter = formatPainterRef.current;
      if (painter && painter.tabId === activeTab.id && selKey !== painter.sourceKey) {
        const g = gridRefs.current[activeTab.id];
        g?.setRangeFormat(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, painter.format);
        formatPainterRef.current = null;
        setPainterArmed(false);
        userEditedSinceBatchRef.current[activeTab.id] = true;
        dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
      }

      const selArea = (sel.endRow - sel.startRow + 1) * (sel.endCol - sel.startCol + 1);
      if (selArea <= 1) {
        setSelectionStats(null);
        statsSelKeyRef.current = selKey;
      } else if (selKey !== statsSelKeyRef.current) {
        statsSelKeyRef.current = selKey;
        if (selArea > 5000) {
          setSelectionStats(null);
        } else {
          const grid = gridRefs.current[activeTab.id];
          let nonEmpty = 0, numericCount = 0, sum = 0, min = Infinity, max = -Infinity;
          for (let r = sel.startRow; r <= sel.endRow; r++) {
            for (let c = sel.startCol; c <= sel.endCol; c++) {
              const v = grid?.getCell(sel.sheet, r, c)?.value;
              if (v === null || v === undefined || v === "") continue;
              nonEmpty++;
              const n =
                typeof v === "number"
                  ? v
                  : typeof v === "string" && v.trim() !== "" && !isNaN(Number(v))
                    ? Number(v)
                    : null;
              if (n !== null) { numericCount++; sum += n; if (n < min) min = n; if (n > max) max = n; }
            }
          }
          setSelectionStats(
            nonEmpty === 0
              ? null
              : {
                  count: nonEmpty,
                  numericCount,
                  sum,
                  avg: numericCount ? sum / numericCount : 0,
                  min: numericCount ? min : 0,
                  max: numericCount ? max : 0,
                },
          );
        }
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

  // Diagnostic harness: replicate the agent's read_range tool exactly —
  // same ready/settle gates, same getCell reads, same JSON shape — without
  // burning an agent turn. Usage from the devtools console:
  //   await gridpathReadRange("Model", "A55:CL55")
  useEffect(() => {
    (window as any).gridpathReadRange = async (sheetName: string, range: string) => {
      const tab = workspaceRef.current.tabs.find((t) => t.id === workspaceRef.current.activeTabId);
      if (!tab) return { ok: false, error: "no active tab" };
      const grid = gridRefs.current[tab.id];
      const ready = grid ? await grid.whenReady() : false;
      if (!ready) return { ok: false, error: "grid_not_ready" };
      await grid?.whenCalculated();
      const rangeCells = expandA1Range(range);
      const cells = rangeCells.slice(0, 500).map(({ row, col }) => {
        const c = grid?.getCell(sheetName, row, col);
        const out: { cell: string; value: any; display?: string; formula: string | null; file_saved?: any } = {
          cell: a1Of(row, col),
          value: c?.value ?? null,
          formula: c?.formula ?? null,
        };
        if (c?.display !== undefined) out.display = c.display;
        const fs = fileSavedIfDivergent(out.value, out.formula, grid?.getFileSavedCell?.(sheetName, row, col));
        if (fs !== undefined) out.file_saved = fs;
        return out;
      });
      const payload = {
        sheet: sheetName,
        range,
        cells: cells.filter((c) => c.value !== null || c.formula !== null),
        truncated: rangeCells.length > 500,
      };
      console.log(`[read-replicate] ${sheetName}!${range}:`, payload);
      return payload;
    };
    return () => {
      delete (window as any).gridpathReadRange;
    };
  }, []);

  // Sync the persisted engine flag (settings DB) into its localStorage
  // any computed values that arrived while the tab was in the background.
  useEffect(() => {
    if (activeTab?.id) ironcalcEngineSetActiveTab(activeTab.id);
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
    // Engine mode: register an empty IronCalc model for this untitled tab
    // (matches UniverGrid's default blank workbook: one "Sheet1").
    ironcalcEngineOnBlank(tab.id, ["Sheet1"]);
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
          setLoadStage(tab.id, "Reading file…");
          const bytes = await readWorkbookBytes(selected);
          await grid.loadBytes(bytes, (stage) => setLoadStage(tab.id, stage));
          console.log("[open] xlsx loaded into tab", tab.id);
          void ironcalcShadowLoad(selected, tab.id);
          void ironcalcEngineOnOpen(tab.id, selected);
        } catch (err) {
          toast({ title: "Open failed", description: String(err), status: "error", duration: 4000 });
          dispatch({ type: "close", tabId: tab.id });
          return;
        } finally {
          setLoadStage(tab.id, null);
        }
        // Opening a file is "I want a fresh session on this file" — prior
        // batches aren't replayed. If the user wants to resume a prior
        // session, they click it in the sidebar (which loads that session's
        // messages from the DB).
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
   * Parse (or re-validate) reference workbooks in the background, updating
   * the per-path error map that drives chip styling. Called on attach, on
   * session resume, and before each agent turn — the underlying loader
   * caches by path so repeat calls are cheap. Sequential on purpose: each
   * parse spikes transient memory, and concurrent parses of several analyst
   * models can OOM the webview.
   */
  const warmLoadReferences = useCallback(async (paths: string[]) => {
    for (const p of paths) {
      try {
        await loadReferenceWorkbook(p);
        setReferenceErrors((s) => {
          if (!(p in s)) return s;
          const { [p]: _gone, ...rest } = s;
          return rest;
        });
      } catch (e) {
        console.warn("[reference] load failed:", p, e);
        setReferenceErrors((s) => ({ ...s, [p]: String(e) }));
      }
    }
  }, []);

  /** Attach reference workbooks to the active session via the file picker. */
  const attachReferenceDialog = useCallback(async () => {
    if (!activeTab) return;
    try {
      const selected = await open({
        multiple: true,
        filters: [{ name: "Excel", extensions: ["xlsx"] }],
      });
      if (!selected) return;
      const picked = (Array.isArray(selected) ? selected : [selected]).filter(
        (p): p is string => typeof p === "string",
      );
      if (picked.length === 0) return;
      const current = activeTab.referencePaths;
      // Don't allow attaching the workbook being edited as its own reference.
      const additions = picked.filter((p) => !current.includes(p) && p !== activeTab.path);
      const next = [...current, ...additions].slice(0, MAX_REFERENCES);
      if (current.length + additions.length > MAX_REFERENCES) {
        toast({
          title: `At most ${MAX_REFERENCES} references per session`,
          status: "info",
          duration: 3000,
        });
      }
      if (next.length === current.length) return;
      dispatch({ type: "set_reference_paths", tabId: activeTab.id, paths: next });
      // Fire-and-forget persist — no-ops if the session row doesn't exist
      // yet (submitPrompt re-persists after the first upsert).
      setSessionReferences(activeTab.id, next).catch((e) =>
        console.warn("[reference] persist failed:", e),
      );
      void warmLoadReferences(next);
    } catch (err) {
      toast({ title: "Attach failed", description: String(err), status: "error", duration: 4000 });
    }
  }, [activeTab, toast, warmLoadReferences]);

  const removeReference = useCallback(
    (path: string) => {
      if (!activeTab) return;
      const next = activeTab.referencePaths.filter((p) => p !== path);
      dispatch({ type: "set_reference_paths", tabId: activeTab.id, paths: next });
      // Only evict the parse cache if no other open tab still references it.
      const stillUsed = workspaceRef.current.tabs.some(
        (t) => t.id !== activeTab.id && t.referencePaths.includes(path),
      );
      if (!stillUsed) evictReference(path);
      setReferenceErrors((s) => {
        if (!(path in s)) return s;
        const { [path]: _gone, ...rest } = s;
        return rest;
      });
      setSessionReferences(activeTab.id, next).catch((e) =>
        console.warn("[reference] persist failed:", e),
      );
    },
    [activeTab],
  );

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
      // Restore attached reference workbooks from the session row. Malformed
      // JSON (shouldn't happen) degrades to "no references".
      const referencePaths: string[] = (() => {
        try {
          const parsed = JSON.parse(row.reference_paths ?? "[]");
          return Array.isArray(parsed) ? parsed.filter((p) => typeof p === "string") : [];
        } catch {
          return [];
        }
      })();
      // Build a tab whose id matches the persisted session id so all
      // subsequent persists land in the right row.
      const tab: WorkbookTab = (() => {
        const fresh = newTab(row.workbook_path);
        return {
          ...fresh,
          id: row.id,
          name: row.name,
          referencePaths,
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
      // Warm the reference parse cache in the background so the first prompt
      // doesn't pay the parse cost, and missing files show error chips now.
      if (referencePaths.length > 0) void warmLoadReferences(referencePaths);

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
          setLoadStage(tab.id, "Restoring session…");
          const snap = isUntitled ? await readUntitledSnapshot(row.workbook_path) : null;
          if (snap) {
            await grid.loadSnapshot(snap);
            // Engine mode: seed a model from the restored snapshot — there is
            // no file to import, and without this, formulas added after a
            // restore would never calculate (and restored formulas would
            // have no computed values).
            ironcalcEngineOnRestore(tab.id, snap);
          } else {
            try {
              setLoadStage(tab.id, "Reading file…");
              const bytes = await readWorkbookBytes(row.workbook_path);
              await grid.loadBytes(bytes, (stage) => setLoadStage(tab.id, stage));
              void ironcalcShadowLoad(row.workbook_path, tab.id);
              void ironcalcEngineOnOpen(tab.id, row.workbook_path);
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
        } finally {
          setLoadStage(tab.id, null);
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
            // `persisted` gates buildSaveMirror. For FILE-based tabs the
            // grid was just loaded from the xlsx, which already contains
            // everything these batches did — replaying them into the save
            // mirror would apply structural ops a second time, so they are
            // history (persisted: true). UNTITLED drafts have no underlying
            // file: the fresh-export path writes styles/widths/merges ONLY
            // from the mirror, so restored batches must stay mirror-eligible
            // or a restart before save-as silently drops all formatting
            // (the tesla.xlsx bug: grid showed 21 styled cells, the export
            // carried one). The fresh-export path applies no structure
            // mirror, so there is no double-apply risk on this branch.
            batches: Array.from(latest.values()).map((b) => ({ ...b, persisted: !isUntitled })),
          });
        } catch (e) {
          console.warn("[session] message replay failed:", e);
        }
        try {
          await upsertSession(tab.id, tab.name, tab.path);
        } catch {}
      })();
    },
    [workspace, toast, warmLoadReferences],
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
      gridRefCallbacks.current.delete(tabId);
      userEditCallbacks.current.delete(tabId);
      manualSheetOpCallbacks.current.delete(tabId);
      delete manualSheetOpsRef.current[tabId];
      ironcalcEngineOnClose(tabId);
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
        // Let any in-flight recalculation settle before snapshotting: the
        // patch caches formula RESULTS into the file, and a snapshot taken
        // mid-calc would persist half-updated numbers that reopen as-is
        // (the cache suppresses the load-time recalc that would fix them).
        await gridRefs.current[tab.id]?.whenCalculated?.(8000);
        const mirror = buildSaveMirror(tab.batches);
        // Fold in manual toolbar edits. They go AFTER the agent's accepted
        // formats so on overlapping cells the user's explicit click wins
        // (applyStyleMirror merges per-cell in iteration order).
        //
        // Re-read each manual-format cell from the LIVE grid rather than trusting
        // the value we recorded at click time: if the user later undid the format
        // with ⌘Z (Univer native undo), the live cell reflects the reverted state
        // and we must not write the stale format back into the file.
        const grid = gridRefs.current[tab.id];
        const mf = manualFormatsRef.current[tab.id] ?? [];
        if (mf.length) {
          const refreshed = grid
            ? mf.map((e) => {
                const live = grid.getCellFormat(e.sheet, e.row, e.col);
                return { sheet: e.sheet, row: e.row, col: e.col, format: live as CellFormat, background: live.background_color ?? null };
              })
            : mf;
          mirror.cellFormats = [...(mirror.cellFormats ?? []), ...refreshed];
        }
        const mb = manualBordersRef.current[tab.id] ?? [];
        if (mb.length) mirror.cellBorders = [...(mirror.cellBorders ?? []), ...mb];
        // Manual structural edits (merges, row/col insert/delete) from the
        // toolbar. Appended after the agent's accepted ops; for purely-manual
        // edits this preserves click order, which matches the live grid.
        const mm = manualMergesRef.current[tab.id] ?? [];
        if (mm.length) mirror.merges = [...(mirror.merges ?? []), ...mm];
        const mrc = manualRowColRef.current[tab.id] ?? [];
        if (mrc.length) mirror.rowColOps = [...(mirror.rowColOps ?? []), ...mrc];
        const mfreeze = manualFreezeRef.current[tab.id];
        if (mfreeze) mirror.freezePanes = [...(mirror.freezePanes ?? []), mfreeze];
        const mfilter = manualFilterRef.current[tab.id];
        if (mfilter) mirror.autoFilters = [...(mirror.autoFilters ?? []), mfilter];
        const mvis = manualVisibilityRef.current[tab.id] ?? [];
        if (mvis.length) mirror.visibility = [...(mirror.visibility ?? []), ...mvis];
        // Sheets the user added/deleted/renamed via Univer's own sheet-tab UI.
        // Without these the patch builder finds a live sheet with no baseline
        // and no create op and falls back to the lossy full export.
        const msheets = manualSheetOpsRef.current[tab.id] ?? [];
        if (msheets.length) mirror.sheetOps = [...(mirror.sheetOps ?? []), ...msheets];
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

        // Surgical save first: apply the patch to the original package —
        // in place when saving over the file we opened, or onto a clone of
        // it for Save As. Either way every part we don't model (charts,
        // external links, drawings…) survives byte-identical. Untitled tabs
        // have no source package and need the full export (nothing exists
        // to lose). exportPatch handles structure edits too (sheet ops,
        // row/col insert-delete); when it can't represent the edits it
        // returns a typed reason, which drives the content-aware gate below.
        let saved = false;
        const isSaveAs = target !== tab.path;
        // Why the surgical path wasn't used — shown verbatim to the user.
        let fallbackNote: string | null = null;
        if (isUntitled) {
          // Fresh workbook: the full export IS the high-fidelity path.
        } else if (!surgicalSaveEnabled()) {
          fallbackNote = "In-place saving is disabled (gridpath.surgicalSave=off)";
        } else {
          const result = grid?.exportPatch(mirror) ?? {
            ok: false as const,
            reason: "missing_baseline" as const,
          };
          if (!result.ok) {
            fallbackNote = describePatchFallback(result);
            console.warn(
              "[save] surgical patch unavailable:",
              result.reason,
              (result as any).detail ?? "",
            );
          } else {
            try {
              if (isSaveAs) {
                await saveWorkbookPatchedAs(tab.path, target, JSON.stringify(result.patch));
              } else {
                await saveWorkbookPatched(target, JSON.stringify(result.patch));
              }
              saved = true;
            } catch (err) {
              const msg = String(err);
              if (msg.includes("BASE_CHANGED") && !isSaveAs) {
                // The file changed outside GridPath. Never clobber it
                // silently — the user decides via Save As.
                toast({
                  title: "File changed on disk",
                  description: "This workbook was modified by another program. Use Save As to keep your version.",
                  status: "error",
                  duration: 6000,
                });
                return false;
              }
              // Save As with a changed-on-disk source can't clone safely;
              // that and patcher failures reach the gated fallback below.
              console.warn("[save] surgical save failed:", msg);
              fallbackNote = msg.toLowerCase().includes("unsupported")
                ? "The workbook contains features (pivot tables, Excel tables, slicers…) the in-place saver can't restructure yet"
                : msg.toLowerCase().includes("failed validation")
                  ? "The in-place saver produced output that failed its safety check, so nothing was written"
                  : "The in-place save failed unexpectedly";
            }
          }
        }

        if (!saved) {
          // Content-aware gate. Plain files (no charts/comments/links/…)
          // lose nothing meaningful in the full export — proceed silently.
          // At-risk files must never be overwritten in place by the lossy
          // writer; the user gets a copy flow that names the exact losses.
          const risks = !isUntitled ? grid?.getFidelityRisks() ?? null : null;
          const gate = gateFallbackSave({ isUntitled, isSaveAs, risks });
          if (gate === "block") {
            toast({
              title: "Save blocked to protect workbook content",
              description:
                `${fallbackNote ?? "The in-place save path is unavailable"}. ` +
                `Overwriting this file through the fallback writer would remove ${describeFidelityRisks(risks!)}. ` +
                `Use Save As to write a reduced-fidelity copy — the original stays intact.`,
              status: "error",
              duration: 10000,
            });
            return false;
          }
          let bytes = await gridRefs.current[tab.id]?.exportBytes(mirror);
          if (!bytes) throw new Error("nothing to export");
          const source = grid?.getSourceBytes() ?? null;
          // ExcelJS writes schema-invalid cfRule operators for
          // contains-errors/blanks rules — Excel opens the file only via
          // its repair flow, which strips the rules. Fix the XML in place.
          // (Applies to untitled exports too: same writer, same bug.)
          bytes = await sanitizeExportedPackage(bytes);
          if (!isUntitled && source) {
            // ExcelJS guts defined names it can't parse as cell ranges
            // (constant-valued vendor flags like `Name = "TRUE"`) — either
            // dropping them or emitting empty, schema-invalid elements that
            // trip Excel's repair dialog. Restore them from the source
            // package; names the user redefined this session are exempt.
            const editedNames = new Set((mirror.definedNames ?? []).map((d) => d.name));
            const repair = await repairDefinedNames(source, bytes, editedNames);
            if (repair) {
              bytes = repair.bytes;
              if (repair.restored.length || repair.grafted.length || repair.dropped.length) {
                console.info(
                  `[save] defined-name repair: restored ${repair.restored.length}, ` +
                    `re-added ${repair.grafted.length}, dropped ${repair.dropped.length} empty`,
                );
              }
            }
            // Names the repair couldn't save (formula-backed refs the
            // export mangled) are real loss the part-based risk scan can't
            // see — a chartless workbook full of defined names would
            // otherwise overwrite "silently lossless". In place: block.
            if (!isSaveAs) {
              const lostNames = await auditDefinedNameLoss(source, bytes);
              if (lostNames?.length) {
                const sample = lostNames.slice(0, 3).join(", ");
                toast({
                  title: "Save blocked to protect workbook content",
                  description:
                    `${fallbackNote ?? "The in-place save path is unavailable"}. ` +
                    `Overwriting this file would drop ${lostNames.length} defined name${lostNames.length === 1 ? "" : "s"} ` +
                    `(${sample}${lostNames.length > 3 ? ", …" : ""}). ` +
                    `Use Save As to write a reduced-fidelity copy — the original stays intact.`,
                  status: "error",
                  duration: 10000,
                });
                return false;
              }
            }
          }
          if (gate === "export_with_warning") {
            // Reduced-fidelity copy: audit the output against the source,
            // name precisely what this copy would be missing, and get an
            // explicit go-ahead BEFORE anything is written. A toast after
            // the fact is not consent.
            let lost = describeFidelityRisks(risks!);
            if (source) {
              const missing = await auditExportLoss(source, bytes);
              if (missing) lost = missing.length ? describeMissingParts(missing) : "";
              const lostNames = await auditDefinedNameLoss(source, bytes);
              if (lostNames?.length) {
                const namesDesc = `${lostNames.length} defined name${lostNames.length === 1 ? "" : "s"}`;
                lost = lost ? `${lost}, ${namesDesc}` : namesDesc;
              }
            }
            const filename = target.split("/").pop() ?? target;
            const proceed = await confirmReducedFidelitySave(filename, fallbackNote, lost);
            if (!proceed) return false;
          }
          try {
            await writeWorkbookBytes(target, bytes);
          } catch (err) {
            // The Rust write boundary validates full-export bytes against
            // the same Excel-strict invariants as surgical saves. A failure
            // means the exporter produced a file Excel would repair —
            // nothing was written.
            const msg = String(err);
            if (msg.includes("EXPORT_INVALID")) {
              toast({
                title: "Save blocked by the validity check",
                description:
                  "The exported file failed Excel-validity checks, so nothing was written and the original is untouched. " +
                  `Detail: ${msg.replace(/^.*EXPORT_INVALID:\s*/, "")}`,
                status: "error",
                duration: 12000,
              });
              return false;
            }
            throw err;
          }
          // The export baked every batch and manual op into the file, and a
          // full export is NOT replay-idempotent (row inserts / sheet
          // deletes splice again on the next save). Reset the save baseline
          // to the exported bytes and retire the consumed history. Untitled
          // workbooks are exempt: they re-export from the live snapshot
          // each save, and their mirror is the only carrier of formats.
          if (!isUntitled && source && grid) {
            await grid.commitSavedBaseline(bytes);
            dispatch({ type: "batches_mark_persisted", tabId: tab.id });
            delete manualFormatsRef.current[tab.id];
            delete manualBordersRef.current[tab.id];
            delete manualMergesRef.current[tab.id];
            delete manualRowColRef.current[tab.id];
            delete manualFreezeRef.current[tab.id];
            delete manualFilterRef.current[tab.id];
            delete manualVisibilityRef.current[tab.id];
            delete manualSheetOpsRef.current[tab.id];
          }
        }
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
    [toast, confirmReducedFidelitySave],
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

  // --- manual formatting toolbar -------------------------------------------

  /**
   * Apply a partial cell format to the live selection. `toggleKey` flips a
   * boolean attribute (bold/italic/…) based on the anchor cell's current
   * state, so clicking Bold on already-bold cells un-bolds them — matching
   * Excel. Edits go straight onto the Univer cells and are recorded for the
   * save mirror so they persist to xlsx.
   */
  const applyFormatToSelection = useCallback(
    (patch: Partial<CellFormat>, opts?: { toggleKey?: "bold" | "italic" | "underline" | "strike" }) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      if (!grid) return;
      const sel = grid.getActiveSelection?.() ?? null;
      if (!sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }

      let finalPatch: Partial<CellFormat> = patch;
      if (opts?.toggleKey) {
        const cur = grid.getCellFormat(sel.sheet, sel.startRow, sel.startCol);
        finalPatch = { ...patch, [opts.toggleKey]: !cur[opts.toggleKey] };
      }

      const bg = finalPatch.background_color ?? null;
      // Apply to the whole selection in ONE Univer command so a single ⌘Z
      // reverts the entire change (per-cell would need N undos).
      grid.setRangeFormat(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, finalPatch);
      // Still record per-cell for the save mirror (bookkeeping only — these are
      // re-read from the live grid at save time, so an undone format won't
      // reappear in the file).
      const store = (manualFormatsRef.current[activeTab.id] ??= []);
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          store.push({ sheet: sel.sheet, row: r, col: c, format: finalPatch as CellFormat, background: bg });
        }
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * Increase/decrease the decimal places of the selection's number format,
   * mirroring Excel's "Increase/Decrease Decimal" buttons. Reads the anchor
   * cell's current pattern and applies the adjusted one to the whole range.
   */
  const adjustDecimalsSelection = useCallback(
    (delta: 1 | -1) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      if (!grid) return;
      const sel = grid.getActiveSelection?.() ?? null;
      if (!sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      const cur = grid.getCellFormat(sel.sheet, sel.startRow, sel.startCol);
      const next = adjustDecimalPlaces(cur.number_format, delta);
      applyFormatToSelection({ number_format: next });
    },
    [activeTab, applyFormatToSelection, toast],
  );

  /**
   * Step the font size up/down one rung on the Excel size ladder, anchored to
   * the current size of the selection's top-left cell (defaults to 11pt).
   */
  const adjustFontSizeSelection = useCallback(
    (delta: 1 | -1) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      if (!grid) return;
      const sel = grid.getActiveSelection?.() ?? null;
      if (!sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      const cur = grid.getCellFormat(sel.sheet, sel.startRow, sel.startCol);
      applyFormatToSelection({ font_size: stepFontSize(cur.font_size ?? 11, delta) });
    },
    [activeTab, applyFormatToSelection, toast],
  );

  /**
   * Sort the rows of the current selection by its left-most column, ascending
   * or descending (Excel's quick A→Z / Z→A). Cell formatting travels with each
   * row; blanks always sink to the bottom.
   */
  const sortSelection = useCallback(
    (direction: "asc" | "desc") => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid || !sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      if (sel.startRow === sel.endRow) {
        toast({ title: "Select multiple rows to sort", status: "info", duration: 1800 });
        return;
      }
      grid.sortRange(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, sel.startCol, direction);
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * Apply borders to the live selection. "all" gives every cell a full box;
   * "outer" only draws the selection's perimeter; "none" clears all sides.
   */
  const applyBordersToSelection = useCallback(
    (kind: "all" | "outer" | "none") => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      if (!grid) return;
      const sel = grid.getActiveSelection?.() ?? null;
      if (!sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }

      // Apply to the live grid as a single command so one ⌘Z reverts the
      // whole border action (Excel-style), and so the correct Univer facade
      // signature is used.
      grid.setRangeBorders(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, kind);

      // Record per-cell entries so the borders round-trip through ExcelJS on
      // save (ExcelJS stores borders per cell).
      const thin = { style: "thin", color: "#000000" };
      const store = (manualBordersRef.current[activeTab.id] ??= []);
      for (let r = sel.startRow; r <= sel.endRow; r++) {
        for (let c = sel.startCol; c <= sel.endCol; c++) {
          let borders: BordersShape;
          if (kind === "none") {
            borders = { top: null, bottom: null, left: null, right: null };
          } else if (kind === "all") {
            borders = { top: thin, bottom: thin, left: thin, right: thin };
          } else {
            // outer: only the sides on the selection perimeter.
            borders = {
              top: r === sel.startRow ? thin : undefined,
              bottom: r === sel.endRow ? thin : undefined,
              left: c === sel.startCol ? thin : undefined,
              right: c === sel.endCol ? thin : undefined,
            };
          }
          store.push({ sheet: sel.sheet, row: r, col: c, borders });
        }
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * Clear the selection. "contents" wipes values (keeps formatting), "formats"
   * resets formatting (keeps values), "all" does both.
   */
  const clearSelection = useCallback(
    (kind: "all" | "formats" | "contents") => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid || !sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      if (kind === "formats" || kind === "all") {
        // Reset every property we manage back to Excel defaults.
        grid.setRangeFormat(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol, {
          bold: false,
          italic: false,
          underline: false,
          strike: false,
          font_color: "#000000",
          background_color: null,
          number_format: "General",
          wrap_text: false,
          horizontal_align: "left",
          vertical_align: "bottom",
        });
        // Drop any recorded manual formats/borders for these cells so save
        // doesn't re-apply them.
        const inSel = (e: { sheet: string; row: number; col: number }) =>
          !(e.sheet === sel.sheet && e.row >= sel.startRow && e.row <= sel.endRow && e.col >= sel.startCol && e.col <= sel.endCol);
        if (manualFormatsRef.current[activeTab.id])
          manualFormatsRef.current[activeTab.id] = manualFormatsRef.current[activeTab.id].filter(inSel);
        if (manualBordersRef.current[activeTab.id])
          manualBordersRef.current[activeTab.id] = manualBordersRef.current[activeTab.id].filter(inSel);
      }
      if (kind === "contents" || kind === "all") {
        grid.clearRange(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * Copy the anchor cell's format and arm the Format Painter; the next distinct
   * selection (applied by the selection poll) receives that format. Clicking
   * again while armed cancels.
   */
  const toggleFormatPainter = useCallback(() => {
    if (!activeTab) return;
    if (formatPainterRef.current) {
      formatPainterRef.current = null;
      setPainterArmed(false);
      return;
    }
    const grid = gridRefs.current[activeTab.id];
    const sel = grid?.getActiveSelection?.() ?? null;
    if (!grid || !sel) {
      toast({ title: "Select a cell to copy its format", status: "info", duration: 1800 });
      return;
    }
    const fmt = grid.getCellFormat(sel.sheet, sel.startRow, sel.startCol) as CellFormat;
    const sourceKey = `${sel.sheet}:${sel.startRow},${sel.startCol},${sel.endRow},${sel.endCol}`;
    formatPainterRef.current = { tabId: activeTab.id, format: fmt, sourceKey };
    setPainterArmed(true);
  }, [activeTab, toast]);

  /**
   * Freeze panes. "topRow"/"firstCol" freeze the first row/column; "atSelection"
   * freezes everything above and left of the anchor; "none" unfreezes.
   */
  const freezeSelection = useCallback(
    (kind: "topRow" | "firstCol" | "atSelection" | "none") => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid) return;
      let rows = 0;
      let cols = 0;
      if (kind === "topRow") rows = 1;
      else if (kind === "firstCol") cols = 1;
      else if (kind === "atSelection") {
        if (!sel) {
          toast({ title: "Select a cell first", status: "info", duration: 1800 });
          return;
        }
        rows = sel.startRow;
        cols = sel.startCol;
      }
      const sheetName = sel?.sheet ?? activeTab.filename;
      if (kind === "none") grid.unfreezePanes(sheetName);
      else grid.freezePanes(sheetName, rows, cols);
      manualFreezeRef.current[activeTab.id] = { sheet: sheetName, freezeRows: rows, freezeCols: cols };
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /** Open Univer's Find & Replace dialog for the active workbook. */
  const openFindReplace = useCallback(() => {
    if (!activeTab) return;
    gridRefs.current[activeTab.id]?.openFindReplace?.(true);
  }, [activeTab]);

  /** Open Univer's conditional-formatting panel for the active workbook. */
  const openConditionalFormatting = useCallback(() => {
    if (!activeTab) return;
    gridRefs.current[activeTab.id]?.openConditionalFormatting?.();
  }, [activeTab]);

  /** Open Univer's data-validation panel for the active workbook. */
  const openDataValidation = useCallback(() => {
    if (!activeTab) return;
    gridRefs.current[activeTab.id]?.openDataValidation?.();
  }, [activeTab]);

  /**
   * Toggle an AutoFilter over the current selection. The filter range is folded
   * into the save mirror so it round-trips to xlsx (ExcelJS autoFilter).
   */
  const toggleFilterSelection = useCallback(() => {
    if (!activeTab) return;
    const grid = gridRefs.current[activeTab.id];
    const sel = grid?.getActiveSelection?.() ?? null;
    if (!grid || !sel) {
      toast({ title: "Select the data range to filter", status: "info", duration: 1800 });
      return;
    }
    const active = grid.toggleFilter(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
    const range = `${a1Of(sel.startRow, sel.startCol)}:${a1Of(sel.endRow, sel.endCol)}`;
    manualFilterRef.current[activeTab.id] = { sheet: sel.sheet, range: active ? range : null };
    userEditedSinceBatchRef.current[activeTab.id] = true;
    dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
  }, [activeTab, toast]);

  /**
   * Excel's outline level buttons ("1 2 3"): collapse/expand the active
   * sheet's column or row groups to the given depth. Visibility changes are
   * recorded for the save mirror so they round-trip to the file.
   */
  const applyOutlineLevel = useCallback(
    (axis: "cols" | "rows", level: number) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid || !sel) return;
      const res = grid.applyOutlineLevel(sel.sheet, axis, level);
      if (!res) return;
      const ops = (manualVisibilityRef.current[activeTab.id] =
        manualVisibilityRef.current[activeTab.id] ?? []);
      if (res.show.length) {
        ops.push(
          axis === "cols"
            ? { kind: "showColumns", sheet: sel.sheet, columns: res.show }
            : { kind: "showRows", sheet: sel.sheet, rows: res.show },
        );
      }
      if (res.hide.length) {
        ops.push(
          axis === "cols"
            ? { kind: "hideColumns", sheet: sel.sheet, columns: res.hide }
            : { kind: "hideRows", sheet: sel.sheet, rows: res.hide },
        );
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab],
  );

  /** Toggle wrap-text across the selection (off→on based on the anchor cell). */
  const toggleWrapSelection = useCallback(() => {
    if (!activeTab) return;
    const grid = gridRefs.current[activeTab.id];
    const sel = grid?.getActiveSelection?.() ?? null;
    if (!grid || !sel) {
      toast({ title: "Select cells first", status: "info", duration: 1800 });
      return;
    }
    const cur = grid.getCellFormat(sel.sheet, sel.startRow, sel.startCol);
    applyFormatToSelection({ wrap_text: !cur.wrap_text });
  }, [activeTab, applyFormatToSelection, toast]);

  /**
   * Merge / unmerge the selection. "center" also center-aligns the anchor.
   * Recorded into manualMergesRef so the change survives save to xlsx.
   */
  const mergeSelection = useCallback(
    (kind: "center" | "merge" | "unmerge") => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid || !sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      const range = `${a1Of(sel.startRow, sel.startCol)}:${a1Of(sel.endRow, sel.endCol)}`;
      const store = (manualMergesRef.current[activeTab.id] ??= []);
      if (kind === "unmerge") {
        grid.unmergeCells(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
        store.push({ sheet: sel.sheet, range, merge: false });
      } else {
        grid.mergeCells(sel.sheet, sel.startRow, sel.startCol, sel.endRow, sel.endCol);
        store.push({ sheet: sel.sheet, range, merge: true });
        if (kind === "center") {
          grid.setCellFormat(sel.sheet, sel.startRow, sel.startCol, { horizontal_align: "center" });
          (manualFormatsRef.current[activeTab.id] ??= []).push({
            sheet: sel.sheet,
            row: sel.startRow,
            col: sel.startCol,
            format: { horizontal_align: "center" } as CellFormat,
            background: null,
          });
        }
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * Insert/delete rows or columns relative to the selection. Mirrors the
   * agent's structural-edit path: shift Univer live, then record a rowColOp
   * so the saved xlsx (which preserves charts/CF/merges via ExcelJS) shifts
   * its metadata to match.
   */
  const insertSelection = useCallback(
    (action: InsertAction) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const sel = grid?.getActiveSelection?.() ?? null;
      if (!grid || !sel) {
        toast({ title: "Select cells first", status: "info", duration: 1800 });
        return;
      }
      const store = (manualRowColRef.current[activeTab.id] ??= []);
      const rowCount = sel.endRow - sel.startRow + 1;
      const colCount = sel.endCol - sel.startCol + 1;
      if (action === "rowAbove") {
        grid.insertRows(sel.sheet, sel.startRow, rowCount);
        store.push({ kind: "insertRows", sheet: sel.sheet, before: sel.startRow, count: rowCount });
      } else if (action === "rowBelow") {
        grid.insertRows(sel.sheet, sel.endRow + 1, rowCount);
        store.push({ kind: "insertRows", sheet: sel.sheet, before: sel.endRow + 1, count: rowCount });
      } else if (action === "colLeft") {
        grid.insertColumns(sel.sheet, sel.startCol, colCount);
        store.push({ kind: "insertColumns", sheet: sel.sheet, before: sel.startCol, count: colCount });
      } else if (action === "colRight") {
        grid.insertColumns(sel.sheet, sel.endCol + 1, colCount);
        store.push({ kind: "insertColumns", sheet: sel.sheet, before: sel.endCol + 1, count: colCount });
      } else if (action === "deleteRows") {
        grid.deleteRows(sel.sheet, sel.startRow, rowCount);
        store.push({ kind: "deleteRows", sheet: sel.sheet, start: sel.startRow, count: rowCount });
      } else if (action === "deleteColumns") {
        grid.deleteColumns(sel.sheet, sel.startCol, colCount);
        store.push({ kind: "deleteColumns", sheet: sel.sheet, start: sel.startCol, count: colCount });
      }
      userEditedSinceBatchRef.current[activeTab.id] = true;
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, toast],
  );

  /**
   * A raw user edit reported by the grid (typing in a cell, paste, native
   * context-menu format). Marks this edit as the most recent action so ⌘Z
   * reverts it (not the last agent batch), and flags the tab dirty. The grid
   * only fires this for genuine user edits — programmatic ones are suppressed
   * by its depth guard.
   */
  const handleUserEdit = useCallback((tabId: string) => {
    userEditedSinceBatchRef.current[tabId] = true;
    const t = workspaceRef.current.tabs.find((x) => x.id === tabId);
    if (t && !t.dirty) dispatch({ type: "mark_dirty", tabId, dirty: true });
  }, []);

  /**
   * Record a sheet add/delete/rename the user made through Univer's own UI
   * so the save mirror can explain it to the surgical patcher. Compacted so
   * undo (add then delete) and rename-after-add don't leave phantom ops.
   */
  const handleManualSheetOp = useCallback((tabId: string, op: ManualSheetOp) => {
    const store = (manualSheetOpsRef.current[tabId] ??= []);
    appendManualSheetOp(store, op);
    handleUserEdit(tabId);
  }, [handleUserEdit]);

  // --- agent loop ----------------------------------------------------------

  const submitPrompt = useCallback(async () => {
    if (!activeTab) {
      toast({ title: "Open a workbook first", status: "info", duration: 2500 });
      return;
    }
    const prompt = (promptByTab[activeTab.id] ?? "").trim();
    if (!prompt) return;

    // A new agent turn becomes the most recent action, so ⌘Z should target its
    // batch again (clears any prior "user hand-edited" state for this tab).
    userEditedSinceBatchRef.current[activeTab.id] = false;

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
    upsertSession(activeTab.id, autoName ?? activeTab.name ?? "", activeTab.path)
      .then(() => {
        // Re-persist reference paths now that the row definitely exists —
        // attaches made before the first prompt hit a no-op UPDATE.
        if (activeTab.referencePaths.length > 0) {
          return setSessionReferences(activeTab.id, activeTab.referencePaths);
        }
      })
      .catch((e) => console.warn("[session] upsert on first prompt failed:", e));

    // Persist the user prompt as the first message of this turn.
    appendMessage(activeTab.id, "user", { prompt }).catch((e) =>
      console.warn("[session] append user message failed:", e),
    );

    dispatch({ type: "batch_add", tabId: activeTab.id, batch });
    dispatch({ type: "set_agent_running", tabId: activeTab.id, running: true });
    dispatch({ type: "stream_text_clear", tabId: activeTab.id });
    dispatch({ type: "stream_reasoning_clear", tabId: activeTab.id });
    dispatch({
      type: "set_status",
      tabId: activeTab.id,
      phase: "thinking",
      message: "Connecting to Claude…",
    });
    setPromptByTab((s) => ({ ...s, [activeTab.id]: "" }));

    tintedCellsByBatch.current[batchId] = [];
    oldValuesByBatch.current[batchId] = [];

    // Let React commit and the browser PAINT the user bubble + thinking
    // indicator before the context capture below. captureWorkbookContext is
    // synchronous and heavy on a cache miss (full snapshot walk + content
    // hash + preview build — hundreds of ms on a big model), and without
    // this yield it runs in the same task as the dispatches above, so Enter
    // feels dead until the capture finishes. Double rAF = "a frame was
    // definitely presented" (time-capped: no frames come while hidden).
    await nextFrame(2);

    try {
      const grid = gridRefs.current[activeTab.id] ?? null;
      // Same mid-recalc gate as the read tools: a prompt sent while the
      // formula engine is still settling (first prompt seconds after opening
      // a dense model is the common case) would pin the base+delta capture
      // on half-computed values — the preview then lacks evaluated samples
      // and the next settled capture produces different bytes, re-billing
      // the full cached context block for no edit. Bounded wait (15s cap).
      await grid?.whenCalculated();
      const workbookContext = captureWorkbookContext(activeTab.path, grid);
      // One-time scan for live results that are unusable vs Excel's cache
      // (cached per load — stable bytes, safe for the cached context block).
      // Warns the agent which sheets collapsed to 0/error so it prefers
      // file_saved instead of distrusting the user's model (rule 17b).
      const calcHealthLine = formatCalcHealthLine((await grid?.getCalcHealth?.()) ?? null);
      if (calcHealthLine) workbookContext.calc_health = calcHealthLine;
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
      // Always tell the agent which sheet the user is looking at — without
      // this, a prompt like "fix this column" sent with no multi-cell
      // selection gave the agent no idea which of the N sheets is on screen.
      const activeSheetName = grid?.getActiveSheetName?.() ?? null;
      const focusParts: string[] = [];
      if (activeSheetName) {
        focusParts.push(`# User view\nThe user is currently viewing sheet "${activeSheetName}".`);
      }
      if (focus) focusParts.push(focus.text);
      if (focusParts.length > 0) workbookContext.focus = focusParts.join("\n\n");
      // Attach reference workbooks (read-only context). Load through the
      // cache — first prompt after attach/resume pays the parse, later
      // turns are free. Files that fail to load (deleted, unreadable) are
      // skipped and their chips flip to the error state.
      if (activeTab.referencePaths.length > 0) {
        await warmLoadReferences(activeTab.referencePaths);
        const loaded = activeTab.referencePaths
          .map((p) => getCachedReference(p))
          .filter((r): r is ParsedReference => r !== null);
        if (loaded.length > 0) {
          workbookContext.references = captureReferenceContext(loaded);
        }
      }
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
  }, [activeTab, promptByTab, toast, warmLoadReferences]);

  // Listen for agent events from Rust and dispatch into the right tab.
  useEffect(() => {
    const unsub = subscribeAgentEvents((ev: AgentEvent) => {
      // Verbose during v1 dev — flip to a debug flag once the loop is stable.
      console.log("[agent] event:", ev);
      switch (ev.kind) {
        case "started":
          dispatch({ type: "stream_text_clear", tabId: ev.tab_id });
          dispatch({ type: "stream_reasoning_clear", tabId: ev.tab_id });
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
        case "reasoning":
          // Same synchronous accumulation for the model's plan/thinking.
          streamingReasoningByBatch.current[ev.batch_id] =
            (streamingReasoningByBatch.current[ev.batch_id] ?? "") + ev.delta;
          dispatch({ type: "stream_reasoning_append", tabId: ev.tab_id, delta: ev.delta });
          return;
        case "tool_call": {
          // keep_pages / read_source are handled ENTIRELY in Rust (context
          // pruning); their tool_call events reach the webview only as
          // stream notifications. Interpreting them here produced a spurious
          // "unknown tool" warning plus an error result Rust ignores.
          if (ev.name === "keep_pages" || ev.name === "read_source" || ev.name === "edgar_lookup") return;
          // Whole tool_call cycle (interpret → mutate → settle → report)
          // serialized per-tab. Rationale: see toolCallQueueByTab declaration.
          const runToolTask = async () => {
          // First thing in the task, before any await that could delay it:
          // this is the moment the tool stops waiting in line and starts
          // costing execution time, which is what Rust's readback budget is
          // actually sized for.
          void reportToolStarted(ev.tool_use_id);
          let result = interpretToolCall(ev.name, ev.input);
          // Populated by the run_script branch; merged into the final tool
          // result so the model sees its own log() output for debugging.
          let scriptLogs: string[] = [];
          let wasScript = false;
          // Populated by the stage_data branch; merged into the final tool
          // result so the model gets the staged block's address map.
          let stageReport: import("./agent/stageLayout").StageReport | null = null;
          const perfT0 = performance.now();
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
            // Gate Accept/Reject presentation on layout validation: scan every
            // cell this batch touched for #REF! / self-refs / formula errors.
            // If anything's broken, reject the `done` and ship the errors back
            // so the agent must patch before the batch is finalized.
            const grid = gridRefs.current[ev.tab_id];
            const tabSnap = workspaceRef.current.tabs.find((t) => t.id === ev.tab_id);
            const batchSnap = tabSnap?.batches.find((b) => b.id === ev.batch_id);
            // Let formula engine settle before scanning.
            await nextFrame(2);
            const errors = validateBatchLayout(grid, batchSnap?.mutations ?? []);
            if (errors.length > 0) {
              const sample = errors.slice(0, 25).map((e) => ({
                sheet: e.sheet,
                cell: e.cell,
                kind: e.kind,
                value: e.value,
                formula: e.formula,
              }));
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "layout_validation_failed",
                    message:
                      "Batch has formula/layout errors — do NOT call done yet. " +
                      "Fix every listed cell (compose against row_map; patch off-by-one / circular refs), " +
                      "then call done again.",
                    errors: sample,
                    truncated: errors.length > 25,
                    error_count: errors.length,
                  }),
                );
              } catch (e) {
                console.warn("[agent] done validation report failed:", e);
              }
              dispatch({
                type: "set_status",
                tabId: ev.tab_id,
                phase: "writing",
                message: `Fixing ${errors.length} formula error${errors.length === 1 ? "" : "s"}…`,
              });
              return;
            }
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({ ok: true }),
              );
            } catch (e) {
              console.warn("[agent] done ok report failed:", e);
            }
            dispatch({
              type: "batch_set_justification",
              tabId: ev.tab_id,
              batchId: ev.batch_id,
              justification: result.justification,
              turnSummary: result.turn_summary,
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
            // A not-ready grid reads back as "everything is empty" — that's
            // misinformation, not data. Same gate as the write path.
            const readReady = grid ? await grid.whenReady() : false;
            if (!readReady) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "grid_not_ready",
                    message: "The spreadsheet is still initializing — retry this exact tool call.",
                  }),
                );
              } catch {}
              return;
            }
            // Reads must not observe mid-recalc state: Univer clears the
            // computed value of dirty formula cells while the worker engine
            // recalculates, so a read racing a recalc reports nulls (or
            // pre-write values) as if they were data. Bounded wait — on
            // timeout we serve whatever is computed rather than failing.
            await grid?.whenCalculated();
            // Const copies: `result` is a `let` (copy_range reassigns it), so
            // closures can't rely on its narrowed type.
            const readSheet = result.sheet;
            const rangeCells = expandA1Range(result.range);
            const cells = rangeCells.slice(0, 500).map(({ row, col }) => {
              const c = grid?.getCell(readSheet, row, col);
              const out: { cell: string; value: any; display?: string; formula: string | null; file_saved?: any } = {
                cell: a1Of(row, col),
                value: c?.value ?? null,
                formula: c?.formula ?? null,
              };
              // Raw value is the source of truth; `display` (present only
              // when a number format changes the rendering, e.g. 0.7235 →
              // "72.35%") tells the agent what the user sees without
              // tempting it to treat the data as text.
              if (c?.display !== undefined) out.display = c.display;
              // Live-vs-file (rule 17b): on unmodified formula cells whose
              // live result is unusable (0, empty, error — INDIRECT,
              // external refs…), surface Excel's last-saved value so the
              // agent reports the right figure instead of distrusting the
              // user's model. Two plausible numbers that merely disagree
              // are not flagged — the user is looking at the live grid.
              const fs = fileSavedIfDivergent(
                out.value,
                out.formula,
                grid?.getFileSavedCell?.(readSheet, row, col),
              );
              if (fs !== undefined) out.file_saved = fs;
              return out;
            });
            const kept = cells.filter((c) => c.value !== null || c.formula !== null);
            const divergent = kept.filter((c) => c.file_saved !== undefined).length;
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({
                  sheet: result.sheet,
                  range: result.range,
                  cells: kept,
                  truncated: rangeCells.length > 500,
                  ...(divergent > 0
                    ? {
                        divergent_cells: divergent,
                        note:
                          "file_saved = Excel's last-saved value where live evaluation is unusable " +
                          "(0, empty, or error — engine limitation, NOT a defect in the workbook). " +
                          "Prefer file_saved for reporting; see rule 17b",
                      }
                    : {}),
                }),
              );
            } catch (e) {
              console.warn("[agent] read_range reportToolResult failed:", e);
            }
            return;
          }
          if (result.kind === "describe_workbook" || result.kind === "find_rows") {
            // Read-only index lookups. Served from the memoized structural
            // index (rebuilt only when the content hash changes), so both
            // are near-instant regardless of workbook size.
            const grid = gridRefs.current[ev.tab_id];
            const idxReady = grid ? await grid.whenReady() : false;
            if (!idxReady) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "grid_not_ready",
                    message: "The spreadsheet is still initializing — retry this exact tool call.",
                  }),
                );
              } catch {}
              return;
            }
            // Same mid-recalc gate as read_range: the snapshot feeds a
            // content-hashed index (formulas + stored values), so cleared
            // dirty-cell values would serve stale row samples AND defeat the
            // index memoization by changing the hash every recalc pass.
            await grid?.whenCalculated();
            const tabSnap = workspaceRef.current.tabs.find((t) => t.id === ev.tab_id);
            const snapshot = grid?.getWorkbookSnapshot?.();
            const index = snapshot
              ? getWorkbookIndex(tabSnap?.path ?? ev.tab_id, snapshot)
              : null;
            let payload: object;
            if (!index) {
              payload = { ok: false, error: "no_workbook", message: "No workbook is loaded." };
            } else if (result.kind === "describe_workbook") {
              payload = describeWorkbookPayload(index, result.sheet);
            } else {
              const { matches, total } = findRowsInIndex(
                index,
                result.query,
                result.sheet,
                result.max_results,
              );
              // Attach one evaluated sample per match, read from the LIVE
              // grid at result time — the index stores addresses, never
              // values, so samples can't be stale.
              payload = {
                matches: matches.map((m) => {
                  const out: any = {
                    sheet: m.sheet,
                    row: m.row,
                    label: m.label,
                  };
                  if (m.section !== null) out.section = m.section;
                  if (m.column !== undefined) out.column = m.column;
                  if (m.sampleCol !== null) {
                    const c = grid?.getCell(m.sheet, m.row - 1, m.sampleCol);
                    if (c && (c.value !== null || c.formula !== null)) {
                      const addr = a1Of(m.row - 1, m.sampleCol);
                      out.sample = c.formula
                        ? `${addr} = ${c.formula} → ${c.value ?? ""}`
                        : `${addr} = ${c.value}`;
                    }
                  }
                  return out;
                }),
                total,
                truncated: total > matches.length,
              };
            }
            try {
              await reportToolResult(ev.tool_use_id, JSON.stringify(payload));
            } catch (e) {
              console.warn(`[agent] ${result.kind} reportToolResult failed:`, e);
            }
            return;
          }
          if (result.kind === "read_reference") {
            // Read-only lookup against an attached reference workbook. Served
            // from the in-memory parse cache — never touches Univer or disk.
            const tab = workspaceRef.current.tabs.find((t) => t.id === ev.tab_id);
            const payload = readReferenceRange(
              tab?.referencePaths ?? [],
              result.workbook,
              result.sheet,
              result.range,
            );
            try {
              await reportToolResult(ev.tool_use_id, payload);
            } catch (e) {
              console.warn("[agent] read_reference reportToolResult failed:", e);
            }
            return;
          }
          if (result.kind === "copy") {
            // copy_range: read the source rectangle from the live grid and
            // expand into ordinary set_cell/set_format mutations (relative
            // refs shifted, formats grouped), then fall through to the
            // normal mutation-apply path below — diff tinting, undo,
            // Accept/Reject and the save mirror all work unchanged.
            const copyGrid = gridRefs.current[ev.tab_id];
            const copyReady = copyGrid ? await copyGrid.whenReady() : false;
            if (!copyReady) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "grid_not_ready",
                    message: "The spreadsheet is still initializing — retry this exact tool call.",
                  }),
                );
              } catch {}
              return;
            }
            const copySheets = new Set(copyGrid!.getSheetNames());
            const copyMissing = [result.sheet, result.dest_sheet].filter((s) => !copySheets.has(s));
            if (copyMissing.length > 0) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "unknown_sheet",
                    missing: copyMissing,
                    available: [...copySheets],
                    message: "Nothing was copied. Use an exact sheet name from `available` (or create_sheet first).",
                  }),
                );
              } catch {}
              return;
            }
            const expansion = expandCopy(copyGrid!, result);
            if (!expansion.ok) {
              try {
                await reportToolResult(ev.tool_use_id, JSON.stringify(expansion));
              } catch {}
              return;
            }
            result = { kind: "mutations", mutations: expansion.mutations };
          }
          if (result.kind === "script") {
            // run_script: execute the model's program in a sandboxed worker
            // against a snapshot-derived read model, then convert its
            // recorded writes into ordinary mutations and fall through to
            // the standard apply path — tinting, review, undo and the save
            // mirror never know a script was involved.
            const scriptGrid = gridRefs.current[ev.tab_id];
            const scriptReady = scriptGrid ? await scriptGrid.whenReady() : false;
            if (!scriptReady) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "grid_not_ready",
                    message: "The spreadsheet is still initializing — retry this exact tool call.",
                  }),
                );
              } catch {}
              return;
            }
            dispatch({
              type: "set_status",
              tabId: ev.tab_id,
              phase: "writing",
              message: "Running script…",
            });
            // The script's read model is built from getWorkbookSnapshot();
            // a snapshot taken mid-recalc still carries pre-write values for
            // dependent formula cells (literal-only write batches skip the
            // engine-idle wait, so a recalc is often still running when the
            // agent's follow-up script arrives). Wait for idle first.
            await scriptGrid!.whenCalculated();
            const exec = await executeSheetScript(scriptGrid!, result.code);
            if (!exec.ok) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "script_error",
                    message:
                      `Script failed — NOTHING was written. Fix the script and call run_script again. Error: ${exec.error}`,
                    logs: exec.logs,
                  }),
                );
              } catch {}
              return;
            }
            scriptLogs = exec.logs;
            wasScript = true;
            result = { kind: "mutations", mutations: exec.mutations };
          }
          if (result.kind === "stage") {
            // stage_data: the harness (not the model) picks placement on the
            // staging sheet — two rows below its used range — lays out
            // title/provenance/header/data as ordinary mutations, and the
            // final tool result carries an address map keyed by the labels
            // the model supplied. Falls through the standard apply path, so
            // review, undo, auto-create-sheet, and readback all just work.
            const stageGrid = gridRefs.current[ev.tab_id];
            const stageReady = stageGrid ? await stageGrid.whenReady() : false;
            if (!stageReady) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "grid_not_ready",
                    message: "The spreadsheet is still initializing — retry this exact tool call.",
                  }),
                );
              } catch {}
              return;
            }
            let stageStartRow = 0;
            const stageSnap = stageGrid!.getWorkbookSnapshot();
            if (stageSnap?.sheets) {
              for (const id of Object.keys(stageSnap.sheets)) {
                const sh = stageSnap.sheets[id];
                if (sh?.name !== result.sheet) continue;
                let last = -1;
                const cellData = sh.cellData ?? {};
                for (const r of Object.keys(cellData)) {
                  const rowIdx = Number(r);
                  if (!Number.isFinite(rowIdx) || rowIdx <= last) continue;
                  const rowCells = cellData[r] ?? {};
                  if (Object.keys(rowCells).length > 0) last = rowIdx;
                }
                stageStartRow = last < 0 ? 0 : last + 2;
                break;
              }
            }
            const laid = layoutStageBlock(result, stageStartRow);
            stageReport = laid.report;
            result = { kind: "mutations", mutations: laid.mutations };
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
          // GATE: never mutate a grid that isn't ready. On a fresh tab the
          // first tool call can beat Univer's dynamic imports; every write
          // below is optional-chained, so without this the whole batch
          // silently no-ops and the agent builds on a phantom layout.
          const gridReady = grid ? await grid.whenReady() : false;
          if (!gridReady) {
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({
                  ok: false,
                  error: "grid_not_ready",
                  message:
                    "The spreadsheet is still initializing — nothing was written. Retry this exact tool call.",
                }),
              );
            } catch {}
            return;
          }
          // Writes to sheets that don't exist: near-miss spellings of an
          // existing sheet (case/whitespace variants) still fail loudly as
          // probable typos — resolveRange would otherwise fall back to the
          // ACTIVE sheet. Genuinely NEW names auto-create the sheet, so
          // organizing a build across purposeful tabs costs nothing.
          const knownSheets = new Set(grid!.getSheetNames());
          const missingSheets = new Set<string>();
          for (const raw of result.mutations) {
            if (raw.type === "create_sheet") continue;
            const sheetNames: string[] = [];
            if ("sheet" in raw && typeof (raw as any).sheet === "string") sheetNames.push((raw as any).sheet);
            if ("dest_sheet" in raw && typeof (raw as any).dest_sheet === "string") sheetNames.push((raw as any).dest_sheet);
            if (raw.type === "set_cell") sheetNames.push(raw.address.sheet);
            for (const n of sheetNames) {
              if (n && !knownSheets.has(n)) missingSheets.add(n);
            }
          }
          if (missingSheets.size > 0) {
            const normalize = (s: string) => s.toLowerCase().replace(/\s+/g, "");
            const byNorm = new Map<string, string>();
            for (const k of knownSheets) byNorm.set(normalize(k), k);
            const typos: Array<{ requested: string; did_you_mean: string }> = [];
            const toCreate: string[] = [];
            for (const n of missingSheets) {
              const near = byNorm.get(normalize(n));
              if (near) typos.push({ requested: n, did_you_mean: near });
              else toCreate.push(n);
            }
            if (typos.length > 0) {
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "unknown_sheet",
                    probable_typos: typos,
                    available: [...knownSheets],
                    message:
                      "Nothing was written. These sheet names look like near-misses of existing sheets — use the exact existing name (new, distinct names would have been auto-created).",
                  }),
                );
              } catch {}
              return;
            }
            // Prepend creations so the standard apply path lands them (and
            // records them in the batch for review/undo) before any write.
            result = {
              ...result,
              mutations: [
                ...toCreate.map((name) => ({ type: "create_sheet", name, tab_color: null } as any)),
                ...result.mutations,
              ],
            };
          }
          // One state update per tool call, not per cell: mutations are
          // collected here and appended in a single dispatch at the end.
          // Per-cell dispatches cloned the whole mutation array each time
          // (O(n²) over a batch) and re-rendered the entire screen per cell.
          const appended: UniverMutation[] = [];
          const append = (m: UniverMutation) => { appended.push(m); };

          // Consecutive set_cell ops on one sheet are buffered and written
          // through ONE Univer command (grid.setCells) — one command
          // pipeline pass, one undo entry, one formula dirty pass — instead
          // of a full cycle per cell. Flushed in chunks with a yielded
          // frame in between so the browser paints (tints, progress) while
          // a big batch lands, instead of freezing until it's done.
          const totalCells = result.mutations.reduce(
            (n: number, r: any) => n + (r.type === "set_cell" ? 1 : 0), 0);
          let appliedCells = 0;
          let cellBuf: { sheet: string; cells: Array<{ row: number; col: number; value: any }> } | null = null;
          // 1000, up from 250: the formula trigger controller debounces
          // dirty-marking with a 100ms timer, so chunk count never multiplied
          // recalc passes — smaller chunks only bought more React commits and
          // paint cycles between them.
          const CELL_CHUNK = 1000;
          const flushCells = async () => {
            const buf = cellBuf;
            cellBuf = null;
            if (!buf || buf.cells.length === 0) return;
            for (let i = 0; i < buf.cells.length; i += CELL_CHUNK) {
              const chunk = buf.cells.slice(i, i + CELL_CHUNK);
              const ok = grid?.setCells(buf.sheet, chunk) ?? false;
              if (!ok) {
                // Batched command unavailable/failed — per-cell fallback.
                for (const c of chunk) grid?.setCell(buf.sheet, c.row, c.col, c.value);
              }
              appliedCells += chunk.length;
              if (totalCells > CELL_CHUNK) {
                dispatch({
                  type: "set_status",
                  tabId: ev.tab_id,
                  phase: "writing",
                  message: `Writing… (${appliedCells}/${totalCells} cells)`,
                });
                await nextFrame();
              }
            }
          };

          // Bulk pre-capture: one bounding-box read per touched sheet (2
          // facade calls) replaces a resolveRange round-trip per cell. Only
          // safe when nothing in the batch shifts coordinates or clears
          // mid-apply — structural/clear batches keep per-cell reads at
          // write time, which see the shifted grid correctly.
          const bulkCaptureSafe = result.mutations.every(
            (m: any) => m.type === "set_cell" || m.type === "set_format",
          );
          const preRead = new Map<string, { value: any; formula: string | null } | null>();
          if (bulkCaptureSafe) {
            const boxes = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
            for (const raw of result.mutations) {
              if (raw.type !== "set_cell") continue;
              const { sheet, row, col } = raw.address;
              const b = boxes.get(sheet);
              if (!b) boxes.set(sheet, { r0: row, c0: col, r1: row, c1: col });
              else {
                b.r0 = Math.min(b.r0, row);
                b.c0 = Math.min(b.c0, col);
                b.r1 = Math.max(b.r1, row);
                b.c1 = Math.max(b.c1, col);
              }
            }
            for (const [sheet, b] of boxes) {
              // A pathologically sparse batch (A1 + ZZ9999) would sweep a
              // huge rectangle for a handful of cells — per-cell is cheaper.
              if ((b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1) > 50_000) continue;
              const data = grid?.getRangeData(sheet, b.r0, b.c0, b.r1, b.c1);
              if (!data) continue;
              for (const raw of result.mutations) {
                if (raw.type !== "set_cell" || raw.address.sheet !== sheet) continue;
                const { row, col } = raw.address;
                const v = data.values[row - b.r0]?.[col - b.c0] ?? null;
                const f = data.formulas[row - b.r0]?.[col - b.c0] ?? null;
                preRead.set(`${sheet}!${row},${col}`, v === null && f === null ? null : { value: v, formula: f });
              }
            }
          }

          for (const raw of result.mutations) {
            if (raw.type === "set_cell") {
              const { sheet, row, col } = raw.address;
              const preKey = `${sheet}!${row},${col}`;
              const before = preRead.has(preKey) ? preRead.get(preKey) : grid?.getCell(sheet, row, col);
              const m: UniverMutation = {
                ...raw,
                old_value: before?.value ?? null,
                old_formula: before?.formula ?? null,
              };
              if (cellBuf && cellBuf.sheet !== sheet) await flushCells();
              (cellBuf ??= { sheet, cells: [] }).cells.push({ row, col, value: m.new_formula ?? m.new_value });
              tintedCellsByBatch.current[ev.batch_id]?.push({ sheet, row, col });
              oldValuesByBatch.current[ev.batch_id]?.push({
                sheet, row, col,
                oldValue: before?.value ?? null,
                oldFormula: before?.formula ?? null,
              });
              append(m);
              continue;
            }
            // Any structural/format op below must see prior cell writes
            // applied (and may shift coordinates) — flush the buffer first.
            await flushCells();
            if (raw.type === "set_format") {
              // Capture each cell's pre-format snapshot so Reject can restore.
              const fm = raw as FormatMutation;
              const old_format = fm.cells.map(({ row, col }) => ({
                row, col,
                format: grid?.getCellFormat(fm.sheet, row, col) ?? null,
              }));
              // Apply ONCE at range level. The per-cell loop this replaces
              // dispatched every facade setter (font, fill, numfmt, …) per
              // CELL — a 3-property format over B2:Z100 was ~7,400 command
              // executions; range-level it's ~3. fm.cells comes from
              // expandA1Range so its bounding box IS the range.
              const rows = fm.cells.map((c) => c.row);
              const cols = fm.cells.map((c) => c.col);
              grid?.setRangeFormat(
                fm.sheet,
                Math.min(...rows),
                Math.min(...cols),
                Math.max(...rows),
                Math.max(...cols),
                fm.new_format,
              );
              for (const { row, col } of fm.cells) {
                tintedCellsByBatch.current[ev.batch_id]?.push({ sheet: fm.sheet, row, col });
              }
              const m: UniverMutation = { ...fm, old_format };
              append(m);
              continue;
            }
            if (raw.type === "set_column_width") {
              const old_widths = raw.columns.map((col) => ({
                col,
                width: grid?.getColumnWidth(raw.sheet, col) ?? null,
              }));
              for (const col of raw.columns) grid?.setColumnWidth(raw.sheet, col, raw.new_width);
              const m: UniverMutation = { ...raw, old_widths };
              append(m);
              continue;
            }
            if (raw.type === "set_row_height") {
              const old_heights = raw.rows.map((row) => ({
                row,
                height: grid?.getRowHeight(raw.sheet, row) ?? null,
              }));
              for (const row of raw.rows) grid?.setRowHeight(raw.sheet, row, raw.new_height);
              const m: UniverMutation = { ...raw, old_heights };
              append(m);
              continue;
            }
            if (raw.type === "merge_cells") {
              grid?.mergeCells(raw.sheet, raw.start_row, raw.start_col, raw.end_row, raw.end_col);
              append(raw);
              continue;
            }
            if (raw.type === "unmerge_cells") {
              grid?.unmergeCells(raw.sheet, raw.start_row, raw.start_col, raw.end_row, raw.end_col);
              append(raw);
              continue;
            }
            if (raw.type === "create_sheet") {
              const ok = grid?.createSheet(raw.name, raw.tab_color ?? null);
              if (ok) append(raw);
              continue;
            }
            if (raw.type === "delete_sheet") {
              // Snapshot the sheet BEFORE delete so Reject can restore it.
              const sheet_snapshot = grid?.getSheetSnapshot(raw.name) ?? null;
              const ok = grid?.deleteSheet(raw.name);
              if (ok) append({ ...raw, sheet_snapshot });
              continue;
            }
            if (raw.type === "rename_sheet") {
              const ok = grid?.renameSheet(raw.old_name, raw.new_name);
              if (ok) append(raw);
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
                // Also stash in oldValuesByBatch so the legacy value-restore
                // path covers clears (and Reject's reverseMutation does too).
                oldValuesByBatch.current[ev.batch_id]?.push({
                  sheet: raw.sheet,
                  row: c.row,
                  col: c.col,
                  oldValue: c.old_value,
                  oldFormula: c.old_formula,
                });
              }
              append(enriched);
              continue;
            }
            if (raw.type === "insert_rows") {
              grid?.insertRows(raw.sheet, raw.before, raw.count);
              append(raw);
              continue;
            }
            if (raw.type === "delete_rows") {
              // Capture non-empty cells in the band before deleting so Reject
              // can re-insert the rows and restore content.
              const deleted_cells = grid?.captureCellBand(
                raw.sheet,
                raw.start,
                0,
                raw.start + raw.count - 1,
                80,
              ) ?? [];
              grid?.deleteRows(raw.sheet, raw.start, raw.count);
              append({ ...raw, deleted_cells });
              continue;
            }
            if (raw.type === "insert_columns") {
              grid?.insertColumns(raw.sheet, raw.before, raw.count);
              append(raw);
              continue;
            }
            if (raw.type === "delete_columns") {
              const deleted_cells = grid?.captureCellBand(
                raw.sheet,
                0,
                raw.start,
                2000,
                raw.start + raw.count - 1,
              ) ?? [];
              grid?.deleteColumns(raw.sheet, raw.start, raw.count);
              append({ ...raw, deleted_cells });
              continue;
            }
            if (raw.type === "freeze_panes") {
              const prior = grid?.getFreezePanes(raw.sheet) ?? { freezeRows: 0, freezeCols: 0 };
              grid?.freezePanes(raw.sheet, raw.freeze_rows, raw.freeze_cols);
              append({ ...raw, old_freeze_rows: prior.freezeRows, old_freeze_cols: prior.freezeCols });
              continue;
            }
            if (raw.type === "unfreeze_panes") {
              const prior = grid?.getFreezePanes(raw.sheet) ?? { freezeRows: 0, freezeCols: 0 };
              grid?.unfreezePanes(raw.sheet);
              append({ ...raw, old_freeze_rows: prior.freezeRows, old_freeze_cols: prior.freezeCols });
              continue;
            }
            if (raw.type === "set_note") {
              const old_text = grid?.getNote(raw.sheet, raw.row, raw.col) ?? null;
              grid?.setNote(raw.sheet, raw.row, raw.col, raw.text);
              append({ ...raw, old_text });
              continue;
            }
            if (raw.type === "delete_note") {
              const old_text = grid?.getNote(raw.sheet, raw.row, raw.col) ?? null;
              grid?.deleteNote(raw.sheet, raw.row, raw.col);
              append({ ...raw, old_text });
              continue;
            }
            if (raw.type === "hide_rows") {
              grid?.hideRows(raw.sheet, raw.rows);
              append(raw);
              continue;
            }
            if (raw.type === "show_rows") {
              grid?.showRows(raw.sheet, raw.rows);
              append(raw);
              continue;
            }
            if (raw.type === "hide_columns") {
              grid?.hideColumns(raw.sheet, raw.columns);
              append(raw);
              continue;
            }
            if (raw.type === "show_columns") {
              grid?.showColumns(raw.sheet, raw.columns);
              append(raw);
              continue;
            }
            if (raw.type === "define_name") {
              // Capture the prior ref (if any) so Reject can restore or
              // delete the name as appropriate.
              const old_ref = grid?.getDefinedNameRef(raw.name) ?? null;
              grid?.defineName(raw.name, raw.ref);
              const m: UniverMutation = { ...raw, old_ref };
              append(m);
              continue;
            }
          }
          await flushCells();
          // The whole tool call lands in the chat state as ONE update.
          dispatch({ type: "batch_append_mutations", tabId: ev.tab_id, batchId: ev.batch_id, mutations: appended });
          dispatch({
            type: "set_status",
            tabId: ev.tab_id,
            phase: "writing",
            message: `Writing… (${tintedCellsByBatch.current[ev.batch_id]?.length ?? 0} cells)`,
          });

          // Report evaluated cell values back to the agent loop so the next
          // turn's tool_result carries the *computed* values (including
          // #VALUE! errors). Recalculation now runs in the worker, so wait
          // for the engine to go idle (bounded) instead of guessing with
          // fixed RAF delays; one RAF after so the mirrored results have
          // landed in the main-thread model.
          // AWAITED inside the queued task so the next tool_call in the
          // queue can't start until this one's report is in flight to Rust.
          // Literal-only data lays skip the engine-idle wait entirely — see
          // isLiteralOnlyBatch. One frame stays either way so the command's
          // main-thread mirror has landed before we read.
          const perfApplied = performance.now();
          const literalOnly = isLiteralOnlyBatch(appended);
          if (!literalOnly) await grid?.whenCalculated();
          await nextFrame();
          const perfSettled = performance.now();
          const evalCells: Array<{ sheet: string; row: number; col: number }> = [];
          for (const raw of result.mutations) {
            if (raw.type === "set_cell") {
              evalCells.push({ sheet: raw.address.sheet, row: raw.address.row, col: raw.address.col });
            } else if (raw.type === "set_format") {
              for (const c of raw.cells) evalCells.push({ sheet: raw.sheet, row: c.row, col: c.col });
            }
          }
          const live = gridRefs.current[ev.tab_id];
          // Bulk readback — same bounding-box trick as the pre-capture: for
          // a 2,000-cell script batch this is 2 facade calls per sheet
          // instead of 2,000 resolveRange round-trips. Cells outside a
          // fetched box (sparse batch, failed read) fall back to getCell.
          const afterRead = new Map<string, { value: any; display?: string; formula: string | null } | null>();
          {
            const boxes = new Map<string, { r0: number; c0: number; r1: number; c1: number }>();
            for (const { sheet, row, col } of evalCells) {
              const b = boxes.get(sheet);
              if (!b) boxes.set(sheet, { r0: row, c0: col, r1: row, c1: col });
              else {
                b.r0 = Math.min(b.r0, row);
                b.c0 = Math.min(b.c0, col);
                b.r1 = Math.max(b.r1, row);
                b.c1 = Math.max(b.c1, col);
              }
            }
            for (const [sheet, b] of boxes) {
              if ((b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1) > 50_000) continue;
              const data = live?.getRangeData(sheet, b.r0, b.c0, b.r1, b.c1);
              if (!data) continue;
              for (const { sheet: s, row, col } of evalCells) {
                if (s !== sheet) continue;
                const v = data.values[row - b.r0]?.[col - b.c0] ?? null;
                const f = data.formulas[row - b.r0]?.[col - b.c0] ?? null;
                if (v === null && f === null) {
                  afterRead.set(`${sheet}!${row},${col}`, null);
                  continue;
                }
                const entry: { value: any; display?: string; formula: string | null } = { value: v, formula: f };
                const d = data.displays?.[row - b.r0]?.[col - b.c0];
                if (typeof d === "string" && d !== "" && v !== null && d !== String(v)) entry.display = d;
                afterRead.set(`${sheet}!${row},${col}`, entry);
              }
            }
          }
          // Qualify the reported address with its sheet when the batch spans
          // more than one — otherwise two sheets writing the same A1 address
          // (which parallel-structured models do constantly) come back as two
          // entries both labelled "H8" with different values. Observed live:
          // the agent read another sheet's #DIV/0! against its own address,
          // concluded its write had failed, and fell back to slow per-cell
          // calls. Single-sheet batches keep the bare "H8" form.
          const readbackSheets = new Set(evalCells.map((c) => c.sheet));
          const qualify = readbackSheets.size > 1;
          const out = evalCells.map(({ sheet, row, col }) => {
            const key = `${sheet}!${row},${col}`;
            const c = afterRead.has(key) ? afterRead.get(key) : live?.getCell(sheet, row, col);
            const entry: { cell: string; value: any; display?: string; formula: string | null } = {
              cell: qualify ? `${sheet}!${a1Of(row, col)}` : a1Of(row, col),
              value: c?.value ?? null,
              formula: c?.formula ?? null,
            };
            const d = (c as { display?: string } | null | undefined)?.display;
            if (d !== undefined) entry.display = d;
            return entry;
          });
          // Address-echo: report the ACTUAL row each written row landed at,
          // keyed by its leftmost non-empty text label. The agent composes
          // formulas against its mental layout, which drifts (an extra header
          // row, a duplicated line, an anchor miscount) and produces
          // off-by-row / circular-ref bugs. Handing back "Revenue = row 58,
          // COGS = row 59 …" lets the next turn reference verified positions
          // instead of guesses. One entry per labeled row (not per cell), so
          // it's small and UNCAPPED — unlike `cells`, which is capped for
          // magnitude sanity only.
          const seenRows = new Map<string, { sheet: string; row: number }>();
          for (const raw of result.mutations) {
            if (raw.type === "set_cell") {
              const key = `${raw.address.sheet}!${raw.address.row}`;
              if (!seenRows.has(key)) {
                seenRows.set(key, { sheet: raw.address.sheet, row: raw.address.row });
              }
            }
          }
          // Bulk label reads for the row_map: one A:D band read per sheet
          // instead of 4 getCell round-trips per labeled row.
          const labelBands = new Map<string, { r0: number; values: any[][] }>();
          {
            const rowSpans = new Map<string, { min: number; max: number }>();
            for (const { sheet, row } of seenRows.values()) {
              const s = rowSpans.get(sheet);
              if (!s) rowSpans.set(sheet, { min: row, max: row });
              else {
                s.min = Math.min(s.min, row);
                s.max = Math.max(s.max, row);
              }
            }
            for (const [sheet, s] of rowSpans) {
              if (s.max - s.min + 1 > 10_000) continue;
              const data = live?.getRangeData(sheet, s.min, 0, s.max, 3);
              if (data) labelBands.set(sheet, { r0: s.min, values: data.values });
            }
          }
          const rowMap = Array.from(seenRows.values())
            .map(({ sheet, row }) => {
              // Prefer column A, but fall back across the first few columns —
              // some templates put section labels in B/C with a blank A.
              const band = labelBands.get(sheet);
              const label = firstLabelIn(
                band
                  ? band.values[row - band.r0] ?? []
                  : Array.from({ length: 4 }, (_, col) => live?.getCell(sheet, row, col)?.value),
              );
              return label
                ? { row: row + 1, label } // 1-indexed A1 row
                : null;
            })
            .filter((e): e is { row: number; label: string } => e !== null)
            .sort((a, b) => a.row - b.row);
          // WRITE VERIFICATION: cells we wrote content into must read back
          // non-empty. A batch that reads back entirely empty means the
          // writes didn't land (engine hiccup, disposed sheet, future
          // regression of the readiness race) — report that as an explicit
          // error so the agent retries instead of building on nothing.
          const expectedContent = new Set<string>();
          for (const raw of result.mutations) {
            if (raw.type === "set_cell" && (raw.new_value != null || raw.new_formula != null)) {
              expectedContent.add(a1Of(raw.address.row, raw.address.col));
            }
          }
          const landed = out.filter(
            (c) => expectedContent.has(c.cell) && (c.value !== null || c.formula !== null),
          ).length;
          const writeCheck =
            expectedContent.size > 0 && landed === 0
              ? "FAILED"
              : landed < expectedContent.size
                ? `partial: ${landed}/${expectedContent.size} cells landed`
                : "ok";
          if (writeCheck === "FAILED") {
            try {
              await reportToolResult(
                ev.tool_use_id,
                JSON.stringify({
                  ok: false,
                  error: "write_verification_failed",
                  message:
                    `Wrote ${expectedContent.size} cells but ALL read back empty — the write did not land. ` +
                    "Retry this exact tool call; if it fails again, read the range first.",
                }),
              );
            } catch (e) {
              console.warn("[agent] write-verification report failed:", e);
            }
            return;
          }
          const content = JSON.stringify({
            ok: true,
            write_check: writeCheck,
            cells: out.slice(0, 200),
            row_map: rowMap,
            ...(stageReport ? { staged: stageReport } : {}),
            // Scripts ALWAYS get a script_logs field: a bare success on a
            // script that never called log() reads as "my logs were dropped"
            // and erodes the model's trust in the whole tool channel.
            ...(wasScript
              ? {
                  script_logs:
                    scriptLogs.length > 0
                      ? scriptLogs
                      : ["<script completed but produced no log() output>"],
                }
              : {}),
          });
          // Per-phase timing so apply-path work is measured, not guessed.
          // apply = interpret + capture + Univer commands; settle = formula
          // engine wait (skipped for literal-only); readback = eval read +
          // row_map + serialization.
          console.log(
            `[agent-perf] ${ev.name}: cells=${totalCells}` +
              ` apply=${Math.round(perfApplied - perfT0)}ms` +
              ` settle=${Math.round(perfSettled - perfApplied)}ms` +
              ` readback=${Math.round(performance.now() - perfSettled)}ms` +
              (literalOnly ? " (literal-only: calc wait skipped)" : ""),
          );
          try {
            await reportToolResult(ev.tool_use_id, content);
          } catch (e) {
            console.warn("[agent] reportToolResult failed:", e);
          }
          }; // end runToolTask
          enqueueToolCallTask(ev.tab_id, async () => {
            try {
              await runToolTask();
            } catch (err) {
              // An uncaught exception here previously reported NOTHING —
              // the Rust loop burned its full readback timeout, substituted
              // "ok", and the agent built on writes that never landed.
              // Fail loudly instead.
              console.error("[agent] tool task crashed:", ev.name, err);
              try {
                await reportToolResult(
                  ev.tool_use_id,
                  JSON.stringify({
                    ok: false,
                    error: "internal_error",
                    message:
                      `Tool '${ev.name}' crashed inside the app: ${String(err).slice(0, 300)}. ` +
                      "Its writes may not have applied — read_range to verify before building on them.",
                  }),
                );
              } catch {}
            }
          });
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
          const fullReasoning = (streamingReasoningByBatch.current[ev.batch_id] ?? "").trim();
          delete streamingReasoningByBatch.current[ev.batch_id];
          (async () => {
            try {
              if (fullReasoning) {
                dispatch({
                  type: "batch_set_reasoning",
                  tabId: ev.tab_id,
                  batchId: ev.batch_id,
                  reasoning: fullReasoning,
                });
              }
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
                  batch: {
                    ...batchSnap,
                    agent_text: fullText || undefined,
                    reasoning: fullReasoning || undefined,
                    status: "pending",
                  },
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

      // The accepted batch is now the most recent action → ⌘Z reverts it.
      userEditedSinceBatchRef.current[activeTab.id] = false;

      dispatch({ type: "batch_accept", tabId: activeTab.id, batchId });
      dispatch({ type: "stream_text_clear", tabId: activeTab.id });

      const batch = findTab(workspace, activeTab.id)?.batches.find((b) => b.id === batchId);
      if (batch) {
        const accepted = { ...batch, status: "accepted" as const };
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

  // --- eval driver ---------------------------------------------------------
  // Self-driving mode for eval/run-gridpath.mjs (GRIDPATH_EVAL_* env). Runs
  // the REAL product path: open the pre-created output file, submit the task
  // prompt verbatim, accept every settled batch, save IN PLACE, write
  // meta.json via eval_finish and exit. Without the env vars getEvalConfig
  // returns null and every hook below is inert.
  const evalRunRef = useRef<{
    cfg: EvalConfig;
    tabId: string;
    t0: number;
    submitted: boolean;
    finishing: boolean;
  } | null>(null);

  const buildEvalMeta = useCallback(async (error: string | null): Promise<EvalMeta> => {
    const ev = evalRunRef.current!;
    const tab = workspaceRef.current.tabs.find((t) => t.id === ev.tabId);
    const choice = await getSettingValue(SETTING_KEYS.apiChoice).catch(() => "");
    const provider = choice === "openai-codex" ? ("openai-codex" as const) : ("claude" as const);
    const [model, effort] = await Promise.all([
      getModel(provider).catch(() => "unknown"),
      getEffort(provider).catch(() => ""),
    ]);
    return {
      harness: "gridpath",
      prompt: ev.cfg.prompt,
      model,
      effort: effort || "default",
      duration_ms: Date.now() - ev.t0,
      batches: tab?.batches.length ?? 0,
      accepted_batches: tab?.batches.filter((b) => b.status === "accepted").length ?? 0,
      input_tokens: tab?.inputTokens ?? 0,
      output_tokens: tab?.outputTokens ?? 0,
      cache_read_tokens: tab?.cacheReadTokens ?? 0,
      cache_creation_tokens: tab?.cacheCreationTokens ?? 0,
      saved: false,
      status: tab?.statusPhase ?? "unknown",
      error,
    };
  }, []);

  // Mount: detect eval launch, open the pre-created output file, stage the
  // prompt. Submission happens in the effect below once the prompt state
  // has committed and the tab is active.
  useEffect(() => {
    (async () => {
      const cfg = await getEvalConfig();
      if (!cfg || evalRunRef.current) return;
      console.log("[eval] driver active:", cfg.start_file);
      const tab = newTab(cfg.start_file);
      evalRunRef.current = { cfg, tabId: tab.id, t0: Date.now(), submitted: false, finishing: false };
      dispatch({ type: "open", tab });
      // Same grid-attach wait as openWorkbookDialog — RAF fires before the
      // commit that attaches the imperative ref, so poll.
      let grid: UniverGridHandle | null = null;
      for (let i = 0; i < 40 && !grid; i++) {
        grid = gridRefs.current[tab.id] ?? null;
        if (!grid) await new Promise((r) => setTimeout(r, 25));
      }
      if (!grid) {
        await evalFinish(await buildEvalMeta("grid ref never attached"), false);
        return;
      }
      try {
        setLoadStage(tab.id, "Reading file…");
        const bytes = await readWorkbookBytes(cfg.start_file);
        await grid.loadBytes(bytes, (stage) => setLoadStage(tab.id, stage));
        void ironcalcShadowLoad(cfg.start_file, tab.id);
        void ironcalcEngineOnOpen(tab.id, cfg.start_file);
      } catch (err) {
        await evalFinish(await buildEvalMeta(`open failed: ${String(err)}`), false);
        return;
      } finally {
        setLoadStage(tab.id, null);
      }
      await grid.whenReady();
      setPromptByTab((s) => ({ ...s, [tab.id]: cfg.prompt }));
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Submit once the staged prompt has committed and the eval tab is active.
  useEffect(() => {
    const ev = evalRunRef.current;
    if (!ev || ev.submitted) return;
    if (!activeTab || activeTab.id !== ev.tabId) return;
    if ((promptByTab[ev.tabId] ?? "") !== ev.cfg.prompt) return;
    ev.submitted = true;
    console.log("[eval] submitting prompt");
    submitPrompt();
  }, [promptByTab, activeTab, submitPrompt]);

  // Completion watcher: when the agent stops, accept everything it produced
  // ("reviewer accepts" is the eval protocol), then save in place and exit.
  // Accepting dispatches workspace updates, which re-enter this effect until
  // nothing is left pending.
  useEffect(() => {
    const ev = evalRunRef.current;
    if (!ev || !ev.submitted || ev.finishing) return;
    const tab = workspace.tabs.find((t) => t.id === ev.tabId);
    if (!tab || tab.agentRunning) return;
    if (tab.batches.length === 0) return;
    if (tab.batches.some((b) => b.status === "streaming")) return;
    const pending = tab.batches.filter((b) => b.status === "pending" && b.mutations.length > 0);
    if (pending.length > 0 && tab.statusPhase !== "error") {
      for (const b of pending) acceptBatch(b.id);
      return;
    }
    ev.finishing = true;
    (async () => {
      // Let post-accept persistence settle before the save reads batches.
      await new Promise((r) => setTimeout(r, 500));
      const liveTab = workspaceRef.current.tabs.find((t) => t.id === ev.tabId);
      const failed = liveTab?.statusPhase === "error";
      let saved = false;
      if (liveTab && !failed) {
        try {
          saved = await saveTabBytes(liveTab, false, true);
        } catch (e) {
          console.error("[eval] save failed:", e);
        }
      } else if (liveTab && failed) {
        // Save anyway so a partial output can be inspected — the run still
        // reports failure via exit code and meta.error.
        try {
          saved = await saveTabBytes(liveTab, false, true);
        } catch {}
      }
      const meta = await buildEvalMeta(failed ? liveTab?.statusMessage ?? "agent error" : null);
      meta.saved = saved;
      console.log("[eval] finishing:", meta);
      await evalFinish(meta, saved && !failed);
    })();
  }, [workspace, acceptBatch, saveTabBytes, buildEvalMeta]);

  const rejectBatch = useCallback(
    async (batchId: string) => {
      if (!activeTab) return;
      const grid = gridRefs.current[activeTab.id];
      const liveBatch = activeTab.batches.find((b) => b.id === batchId);
      // Walk mutations in reverse so nested structural shifts unwind correctly
      // (e.g. insert_rows then set_cell → restore cell first, then delete rows).
      if (liveBatch) {
        const reversed = [...liveBatch.mutations].reverse();
        for (const m of reversed) reverseMutation(grid, m);
      }
      // Legacy value restore for any set_cell that only lived in oldValuesByBatch
      // (defensive — reverseMutation already covers set_cell via old_* on the mutation).
      const olds = oldValuesByBatch.current[batchId] ?? [];
      for (const o of olds) {
        // Skip if the batch mutation already restored this cell.
        const already = liveBatch?.mutations.some(
          (m) =>
            m.type === "set_cell" &&
            m.address.sheet === o.sheet &&
            m.address.row === o.row &&
            m.address.col === o.col,
        );
        if (!already) grid?.setCell(o.sheet, o.row, o.col, o.oldFormula ?? o.oldValue);
      }
      delete tintedCellsByBatch.current[batchId];
      delete oldValuesByBatch.current[batchId];

      dispatch({ type: "batch_reject", tabId: activeTab.id, batchId });
      dispatch({ type: "stream_text_clear", tabId: activeTab.id });
      // The user may have saved while this batch was pending (values + styles
      // land in the file — see buildSaveMirror). The revert above diverges the
      // grid from that file, so the tab needs saving again.
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });

      const batch = findTab(workspace, activeTab.id)?.batches.find((b) => b.id === batchId);
      if (batch) {
        const rejected = { ...batch, status: "rejected" as const };
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
   * (which visually flips the badge) and appends a new session message so
   * the timeline shows the full lifecycle: accepted → then undone.
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

      // Restore every mutation from its old_* / deleted_* / sheet_snapshot
      // state. Iterate reversed so writes-on-top-of-writes and nested
      // structural shifts within the same batch wind back correctly.
      const reversed = [...batch.mutations].reverse();
      for (const m of reversed) reverseMutation(grid, m);

      dispatch({ type: "batch_reject", tabId: activeTab.id, batchId });
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });

      const undone = { ...batch, status: "rejected" as const };
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

      for (const m of batch.mutations) applyMutationForward(grid, m);

      userEditedSinceBatchRef.current[activeTab.id] = false;
      dispatch({ type: "batch_accept", tabId: activeTab.id, batchId });
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });

      const restored = { ...batch, status: "accepted" as const };
      appendMessage(activeTab.id, "agent_batch", { batch: restored }).catch((e) =>
        console.warn("[session] persist redo failed:", e),
      );

      toast({ title: "Redone", description: `Re-applied ${batch.mutations.length} edit${batch.mutations.length === 1 ? "" : "s"}.`, status: "info", duration: 2000 });
    },
    [activeTab, toast],
  );

  /** Toolbar Undo/Redo buttons — drive Univer programmatically (focus is on
   * the button, never the grid) when the last action was a manual edit, else
   * fall back to agent-batch undo/redo. */
  const performUndo = useCallback(() => {
    if (!activeTab) return;
    const grid = gridRefs.current[activeTab.id];
    if (userEditedSinceBatchRef.current[activeTab.id]) {
      grid?.undo();
      return;
    }
    const lastAccepted = [...activeTab.batches].reverse().find((b) => b.status === "accepted");
    if (lastAccepted) undoBatch(lastAccepted.id);
  }, [activeTab, undoBatch]);

  const performRedo = useCallback(() => {
    if (!activeTab) return;
    const grid = gridRefs.current[activeTab.id];
    if (userEditedSinceBatchRef.current[activeTab.id]) {
      grid?.redo();
      return;
    }
    const lastRejected = [...activeTab.batches].reverse().find((b) => b.status === "rejected");
    if (lastRejected) redoBatch(lastRejected.id);
  }, [activeTab, redoBatch]);

  // --- changes review modal -----------------------------------------------

  /** Which batch the full-screen review is open on, and whether the user
   *  minimized it to a chip after a jump-to-cell. */
  const [review, setReview] = useState<{ batchId: string; minimized: boolean } | null>(null);
  const reviewBatch = review
    ? activeTab?.batches.find((b) => b.id === review.batchId) ?? null
    : null;

  const openReview = useCallback((batchId: string) => {
    setReview({ batchId, minimized: false });
  }, []);

  /**
   * Human label for a review region: leftmost text cell of its first row
   * (columns A–D), same heuristic the agent's row_map readback uses.
   */
  const reviewLabelOf = useCallback(
    (sheet: string, row: number): string | null => {
      const grid = activeTab ? gridRefs.current[activeTab.id] : null;
      if (!grid) return null;
      for (let col = 0; col < 4; col++) {
        try {
          const v = grid.getCell(sheet, row, col)?.value;
          if (isLabelText(v)) return v.trim().slice(0, 60);
        } catch {}
      }
      return null;
    },
    [activeTab],
  );

  /**
   * Workbook index snapshot for the open review — computed once per modal
   * open (getWorkbookIndex memoizes on content, so this is at most one
   * hash walk). Supplies per-column headers ("BU" → "FY2026E") without a
   * per-column grid scan.
   */
  const reviewIndex = useMemo(() => {
    if (!review || !activeTab) return null;
    const snapshot = gridRefs.current[activeTab.id]?.getWorkbookSnapshot?.();
    return snapshot ? getWorkbookIndex(activeTab.path, snapshot) : null;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [review?.batchId, activeTab?.id]);

  /** Column header for a review cell, from the index's detected header row. */
  const reviewHeaderOf = useCallback(
    (sheet: string, col: number): string | null => {
      const si = reviewIndex?.sheets.find((s) => s.name === sheet);
      return si?.headers[colLetters(col)] ?? null;
    },
    [reviewIndex],
  );

  /** Jump-to-cell from the review: minimize the modal so the grid is visible. */
  const jumpFromReview = useCallback(
    (sheet: string, row: number, col: number) => {
      if (!activeTab) return;
      gridRefs.current[activeTab.id]?.jumpToCell(sheet, row, col);
      setReview((r) => (r ? { ...r, minimized: true } : r));
    },
    [activeTab],
  );

  /**
   * Per-cell / per-region reject from the review modal: reverse just those
   * set_cell mutations on the grid and drop them from the pending batch, so
   * an eventual Accept applies only what survived.
   */
  const rejectReviewCells = useCallback(
    (cells: Array<{ sheet: string; row: number; col: number }>) => {
      if (!activeTab || !review) return;
      const grid = gridRefs.current[activeTab.id];
      const batch = activeTab.batches.find((b) => b.id === review.batchId);
      if (!batch || batch.status !== "pending") return;
      const keys = new Set(cells.map((c) => `${c.sheet}!${c.row},${c.col}`));
      const targets = batch.mutations.filter(
        (m) =>
          m.type === "set_cell" &&
          keys.has(`${m.address.sheet}!${m.address.row},${m.address.col}`),
      );
      // Reverse in reverse order so repeated writes to one cell unwind to
      // the original pre-batch value.
      for (const m of [...targets].reverse()) reverseMutation(grid, m);
      dispatch({ type: "batch_drop_cells", tabId: activeTab.id, batchId: batch.id, cells });
      dispatch({ type: "mark_dirty", tabId: activeTab.id, dirty: true });
    },
    [activeTab, review],
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
        // If the user has made a direct grid edit (formatting toolbar, etc.)
        // since the last agent batch, that edit — not the batch — is the most
        // recent action. Don't hijack ⌘Z to undo the whole batch.
        if (userEditedSinceBatchRef.current[activeTab.id]) {
          const grid = gridRefs.current[activeTab.id];
          // When the grid canvas itself has focus (e.g. right after typing in a
          // cell), Univer's own ⌘Z keybinding handles it — let it through.
          if (grid?.containsFocus?.()) return;
          // Otherwise focus is on a toolbar button / elsewhere, so Univer never
          // receives the key event. Drive its undo/redo programmatically so a
          // toolbar format/merge/border/sort still reverts with one ⌘Z.
          e.preventDefault();
          if (e.shiftKey) grid?.redo?.();
          else grid?.undo?.();
          return;
        }
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
      $hasActiveTab={!!activeTab}
    >
      <SidebarArea>
        {sidebarOpen && (
          <SessionSidebar
            sessions={recentSessions}
            activePath={activeTab?.path ?? null}
            runningIds={new Set(workspace.tabs.filter((t) => t.agentRunning).map((t) => t.id))}
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

      <RibbonArea>
        {activeTab && (
          <FormatToolbar
            disabled={!activeTab}
            onUndo={performUndo}
            onRedo={performRedo}
            onFormatPainter={toggleFormatPainter}
            painterActive={painterArmed}
            onClear={(kind) => clearSelection(kind)}
            onFreeze={(kind) => freezeSelection(kind)}
            onToggleFilter={toggleFilterSelection}
            onConditionalFormatting={openConditionalFormatting}
            onDataValidation={openDataValidation}
            onFindReplace={openFindReplace}
            outline={outlineLevels}
            onOutlineLevel={applyOutlineLevel}
            onToggle={(key) => applyFormatToSelection({}, { toggleKey: key })}
            onFontColor={(color) => applyFormatToSelection({ font_color: color })}
            onFillColor={(color) => applyFormatToSelection({ background_color: color })}
            onAlign={(a) => applyFormatToSelection({ horizontal_align: a })}
            onVerticalAlign={(a) => applyFormatToSelection({ vertical_align: a })}
            onWrap={() => toggleWrapSelection()}
            onFontFamily={(family) => applyFormatToSelection({ font_family: family })}
            onFontSize={(size) => applyFormatToSelection({ font_size: size })}
            onAdjustFontSize={(delta) => adjustFontSizeSelection(delta)}
            onSort={(direction) => sortSelection(direction)}
            onMerge={(kind) => mergeSelection(kind)}
            onInsert={(action) => insertSelection(action)}
            onNumberFormat={(pattern) => applyFormatToSelection({ number_format: pattern })}
            onAdjustDecimals={(delta) => adjustDecimalsSelection(delta)}
            onBorders={(kind) => applyBordersToSelection(kind)}
          />
        )}
      </RibbonArea>

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
              ref={getGridRefCallback(t.id)}
              tabId={t.id}
              workbookPath={t.path}
              onUserEdit={getUserEditCallback(t.id)}
              onManualSheetOp={getManualSheetOpCallback(t.id)}
            />
            {gridLoading[t.id] && (
              <GridLoadOverlay>
                <LoadSpinner />
                <LoadFilename>{t.filename}</LoadFilename>
                <LoadStage>{gridLoading[t.id]}</LoadStage>
              </GridLoadOverlay>
            )}
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
          onOpenReview={openReview}
          selectionLabel={focusDismissedFor === liveSelection ? null : liveSelection}
          onDismissSelection={() => setFocusDismissedFor(liveSelection)}
          autoApply={autoApply}
          onSetAutoApply={(next) => setAutoApply(next)}
          references={(activeTab?.referencePaths ?? []).map((p) => ({
            path: p,
            label: referenceLabelFromPath(p),
            error: referenceErrors[p],
          }))}
          onAttachReference={attachReferenceDialog}
          onRemoveReference={removeReference}
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
          selectionStats={selectionStats}
          agentRunning={!!activeTab?.agentRunning}
          onOpenSettings={() => setSettingsOpen(true)}
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
      <FidelitySaveModal
        open={!!fidelityPrompt}
        filename={fidelityPrompt?.filename ?? ""}
        note={fidelityPrompt?.note ?? null}
        losses={fidelityPrompt?.losses ?? ""}
        onConfirm={() => settleFidelityPrompt(true)}
        onCancel={() => settleFidelityPrompt(false)}
      />
      {reviewBatch && review && !review.minimized && (
        <ReviewModal
          batch={reviewBatch}
          filename={activeTab?.filename ?? ""}
          labelOf={reviewLabelOf}
          headerOf={reviewHeaderOf}
          onClose={() => setReview(null)}
          onAcceptAll={() => { setReview(null); acceptBatch(reviewBatch.id); }}
          onRejectAll={() => { setReview(null); rejectBatch(reviewBatch.id); }}
          onJump={jumpFromReview}
          onRejectCells={rejectReviewCells}
        />
      )}
      {reviewBatch && review?.minimized && (
        <ReviewChip onClick={() => setReview((r) => (r ? { ...r, minimized: false } : r))}>
          <Maximize2 size={12} />
          Back to review ({reviewBatch.mutations.length})
          <ReviewChipClose
            onClick={(e) => { e.stopPropagation(); setReview(null); }}
            title="Close review"
          >
            <X size={12} />
          </ReviewChipClose>
        </ReviewChip>
      )}
    </Page>
  );
};

/**
 * Build the agent-mutation mirror we ship to UniverGrid.exportBytes so the
 * ExcelJS workbook gets the agent's format / width / height / merge changes
 * patched into it before we writeBuffer. Walks every ACCEPTED and PENDING
 * batch — pending edits are already live on the grid (and their VALUES save
 * unconditionally via the Univer snapshot), so skipping their styling would
 * write a half-applied batch: data without formatting. Save is WYSIWYG.
 * Rejected batches were reverted on the grid, so they're excluded. The
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
  // Defined names: last (accepted/pending) write per name wins, so collapse by name.
  const definedNameMap = new Map<string, { name: string; ref: string }>();
  for (const b of batches) {
    if (b.status !== "accepted" && b.status !== "pending") continue;
    // Already accounted for by the file + save baseline (replayed session
    // history, or batches consumed by a full-export save). Replaying them
    // would apply structural ops a second time.
    if (b.persisted) continue;
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
      } else if (m.type === "define_name") {
        definedNameMap.set(m.name, { name: m.name, ref: m.ref });
      }
    }
  }
  return {
    cellFormats, columnWidths, rowHeights, merges, sheetOps, clears, rowColOps, freezePanes, visibility,
    definedNames: Array.from(definedNameMap.values()),
  };
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

/**
 * Adjust the decimal-place count of an Excel number-format pattern by `delta`
 * (clamped to 0–10), preserving the integer grouping and any trailing suffix
 * like `%`. Operates on the first (positive) section only — good enough for the
 * Increase/Decrease Decimal toolbar buttons. Falls back to a plain integer
 * pattern when the cell has no explicit format ("General").
 */
function adjustDecimalPlaces(pattern: string | undefined, delta: number): string {
  const main = (!pattern || pattern === "General" ? "0" : pattern).split(";")[0];
  const dot = main.indexOf(".");

  let intPart: string;
  let suffix: string;
  let curDecimals: number;
  if (dot >= 0) {
    intPart = main.slice(0, dot);
    const afterDot = main.slice(dot + 1);
    const m = afterDot.match(/^(0*)(.*)$/);
    curDecimals = m ? m[1].length : 0;
    suffix = m ? m[2] : "";
  } else {
    // No decimals yet — split the integer body from any trailing suffix.
    const m = main.match(/^([#0,]*)(.*)$/);
    intPart = m ? m[1] : main;
    suffix = m ? m[2] : "";
    curDecimals = 0;
  }
  if (!intPart) intPart = "0";

  const nextDecimals = Math.max(0, Math.min(10, curDecimals + delta));
  const decimals = nextDecimals > 0 ? "." + "0".repeat(nextDecimals) : "";
  return `${intPart}${decimals}${suffix}` || "0";
}

/** Excel's font-size ladder; the +/- buttons step to the next/previous rung. */
const FONT_SIZE_LADDER = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

/** Bump a font size to the next (delta=1) or previous (delta=-1) ladder rung. */
function stepFontSize(current: number, delta: 1 | -1): number {
  if (delta > 0) {
    const up = FONT_SIZE_LADDER.find((s) => s > current);
    return up ?? Math.min(409, Math.round(current) + 1);
  }
  const down = [...FONT_SIZE_LADDER].reverse().find((s) => s < current);
  return down ?? Math.max(1, Math.round(current) - 1);
}
