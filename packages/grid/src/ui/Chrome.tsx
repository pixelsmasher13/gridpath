/**
 * The spreadsheet's built-in chrome: the formatting toolbar, a File menu
 * (open, new, recent workbooks, version history, download), the save status
 * ("Saving…", "Saved", "Unsaved changes"), and the version history panel.
 */
import React, { useEffect, useRef, useState } from "react";
import styled from "styled-components";
import { Check, ChevronDown, Download, FilePlus, FolderOpen, History, RotateCcw, Save, Sparkles, X } from "lucide-react";
import {
  FormatToolbar,
  IconBtn,
  Menu,
  MenuHint,
  MenuItem,
  MenuSep,
  PopoverHost,
  TextBtn,
  type FormatToolbarProps,
  type MoreControl,
  type ToolbarControl,
} from "../../../../src/screens/SpreadsheetScreen/components/FormatToolbar";
import type { StoredWorkbook, WorkbookVersion } from "./store";
import type { AgentBatch } from "../ai/toolkit";
import { cellSpan, describeBatch } from "./review";

/** Where the workbook stands relative to its last save. */
export type SaveStatus =
  | { state: "saved"; savedAt: string }
  | { state: "unsaved" }
  | { state: "saving" }
  | { state: "error"; message: string }
  /** Agent changes await review; saving waits for accept or reject. */
  | { state: "review"; pending: number }
  /** Nothing to report (fresh workbook, or no saving configured). */
  | { state: "idle" };

/** The package's grid has no panel (or save path yet) for these. */
const HIDDEN: ToolbarControl[] = ["painter", "merge", "borders", "conditionalFormatting", "dataValidation", "findReplace", "outline"];
/** Less-used controls under "More", so the bar fits narrow embeds; more of
 * them as the spreadsheet gets narrower. */
const MORE: MoreControl[] = ["strike", "fontSizeStep", "verticalAlign"];
const MORE_NARROW: MoreControl[] = [...MORE, "align"];
const NARROW = 760;
const NARROWEST = 600;

type Handlers = Omit<FormatToolbarProps, "disabled" | "hidden" | "canUndo" | "canRedo" | "trailing" | "leading" | "painterActive" | "outline">;

export interface ChromeProps {
  actions: Handlers;
  disabled: boolean;
  canUndo: boolean;
  canRedo: boolean;
  status: SaveStatus;
  /** A store or `onSave` is configured: show Save and the status. */
  canSave: boolean;
  onSave: () => void;
  onDownload: () => void;
  /** The File menu's open / new / recent entries (the host may turn them off). */
  fileMenu: boolean;
  onOpenFile: (file: File) => void;
  onNew: () => void;
  listRecent: (() => Promise<StoredWorkbook[]>) | null;
  onOpenRecent: (id: string) => void;
  currentId: string;
  onShowHistory: (() => void) | null;
  /** Open the review panel (null when the host turned it off). */
  onReview: (() => void) | null;
  notice: string | null;
}

function when(iso: string): string {
  const d = new Date(iso);
  const today = new Date();
  const yesterday = new Date(Date.now() - 86_400_000);
  const time = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  if (d.toDateString() === today.toDateString()) return `Today ${time}`;
  if (d.toDateString() === yesterday.toDateString()) return `Yesterday ${time}`;
  return `${d.toLocaleDateString([], { month: "short", day: "numeric" })} ${time}`;
}

function size(bytes: number): string {
  return bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function StatusText({ status, onRetry, onReview, short }: { status: SaveStatus; onRetry: () => void; onReview: (() => void) | null; short?: boolean }) {
  switch (status.state) {
    case "saving":
      return <Status>Saving…</Status>;
    case "saved":
      return <Status title={`Saved ${when(status.savedAt)}`}>Saved</Status>;
    case "unsaved":
      return <Status $warn title="Unsaved changes">{short ? "Unsaved" : "Unsaved changes"}</Status>;
    case "review": {
      const full = `Review ${status.pending} AI change${status.pending === 1 ? "" : "s"} to save`;
      const text = short ? `Review ${status.pending} change${status.pending === 1 ? "" : "s"}` : full;
      return onReview ? (
        <Status $warn>
          <LinkBtn title={short ? `${full}: accept or reject the AI's changes` : "Accept or reject the AI's changes"} onClick={onReview}>{text}</LinkBtn>
        </Status>
      ) : (
        <Status $warn title="Accept or reject the AI's changes to save">{text}</Status>
      );
    }
    case "error":
      return (
        <Status $warn title={status.message}>
          Save failed · <LinkBtn onClick={onRetry}>Retry</LinkBtn>
        </Status>
      );
    default:
      return null;
  }
}

function FileMenu(props: ChromeProps) {
  const [open, setOpen] = useState(false);
  const [recent, setRecent] = useState<StoredWorkbook[] | null>(null);
  const input = useRef<HTMLInputElement | null>(null);
  const host = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (!open) return;
    props.listRecent?.().then(setRecent, () => setRecent([]));
    const onDown = (e: MouseEvent) => {
      if (host.current && !host.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    return () => document.removeEventListener("mousedown", onDown);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);
  const pick = (fn: () => void) => () => {
    setOpen(false);
    fn();
  };
  const others = (recent ?? []).filter((r) => r.id !== props.currentId).slice(0, 8);
  return (
    <PopoverHost ref={host}>
      <TextBtn title="File" disabled={props.disabled} onClick={() => setOpen((o) => !o)}>
        File <ChevronDown size={11} style={{ opacity: 0.7 }} />
      </TextBtn>
      <input
        ref={input}
        type="file"
        accept=".xlsx,.xlsm"
        hidden
        onChange={(e) => {
          const f = e.target.files?.[0];
          e.target.value = "";
          if (f) props.onOpenFile(f);
        }}
      />
      {open && (
        <Menu style={{ minWidth: 220 }}>
          {props.fileMenu && (
            <>
              <MenuItem onClick={pick(() => input.current?.click())}>
                <Row><FolderOpen size={14} /> Open…</Row>
              </MenuItem>
              <MenuItem onClick={pick(props.onNew)}>
                <Row><FilePlus size={14} /> New workbook</Row>
              </MenuItem>
            </>
          )}
          {props.fileMenu && props.listRecent && others.length > 0 && (
            <>
              <MenuSep />
              <MenuHint style={{ padding: "2px 8px" }}>Recent</MenuHint>
              {others.map((r) => (
                <MenuItem key={r.id} onClick={pick(() => props.onOpenRecent(r.id))} title={r.id}>
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 150 }}>{r.name}</span>
                  <MenuHint>{when(r.savedAt)}</MenuHint>
                </MenuItem>
              ))}
            </>
          )}
          {(props.fileMenu || props.onShowHistory) && <MenuSep />}
          {props.onShowHistory && (
            <MenuItem onClick={pick(props.onShowHistory)}>
              <Row><History size={14} /> Version history</Row>
            </MenuItem>
          )}
          <MenuItem onClick={pick(props.onDownload)}>
            <Row><Download size={14} /> Download .xlsx</Row>
          </MenuItem>
        </Menu>
      )}
    </PopoverHost>
  );
}

export function Chrome(props: ChromeProps) {
  const wrap = useRef<HTMLDivElement | null>(null);
  const [width, setWidth] = useState(Infinity);
  useEffect(() => {
    const el = wrap.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(([e]) => setWidth(e.contentRect.width));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  return (
    <ChromeWrap ref={wrap}>
      <FormatToolbar
        {...props.actions}
        disabled={props.disabled}
        hidden={width < NARROWEST ? [...HIDDEN, "font"] : HIDDEN}
        more={width < NARROW ? MORE_NARROW : MORE}
        compact={width < NARROW}
        canUndo={props.canUndo}
        canRedo={props.canRedo}
        painterActive={false}
        outline={null}
        leading={<FileMenu {...props} />}
        trailing={
          <>
            <StatusText status={props.status} onRetry={props.onSave} onReview={props.onReview} short={width < NARROW} />
            {props.canSave && (
              <IconBtn title="Save (⌘S)" disabled={props.disabled || props.status.state === "saving"} onClick={props.onSave}>
                <Save size={14} />
              </IconBtn>
            )}
            {props.onShowHistory && (
              <IconBtn title="Version history" disabled={props.disabled} onClick={props.onShowHistory}>
                <History size={14} />
              </IconBtn>
            )}
            <IconBtn title="Download .xlsx" disabled={props.disabled} onClick={props.onDownload}>
              <Download size={14} />
            </IconBtn>
          </>
        }
      />
      {props.notice && <Notice role="status">{props.notice}</Notice>}
    </ChromeWrap>
  );
}

export function HistoryPanel(props: {
  name: string;
  versions: WorkbookVersion[] | null;
  onRestore: (versionId: string) => void;
  onDownload: (versionId: string) => void;
  onClose: () => void;
}) {
  return (
    <Panel role="dialog" aria-label="Version history">
      <PanelHead>
        <strong>Version history</strong>
        <IconBtn title="Close" onClick={props.onClose}>
          <X size={14} />
        </IconBtn>
      </PanelHead>
      <PanelHint>{props.name} · saved in this browser</PanelHint>
      {props.versions === null && <PanelHint>Loading…</PanelHint>}
      {props.versions?.length === 0 && <PanelHint>No saved versions yet.</PanelHint>}
      {props.versions?.map((v, i) => (
        <VersionRow key={v.id}>
          <div>
            <div>{when(v.savedAt)}{i === 0 ? " · current" : ""}</div>
            <MenuHint>{size(v.size)}</MenuHint>
          </div>
          <div style={{ display: "flex", gap: 4 }}>
            {i > 0 && (
              <TextBtn title="Restore this version" onClick={() => props.onRestore(v.id)}>
                <RotateCcw size={12} /> Restore
              </TextBtn>
            )}
            <IconBtn title="Download this version" onClick={() => props.onDownload(v.id)}>
              <Download size={13} />
            </IconBtn>
          </div>
        </VersionRow>
      ))}
    </Panel>
  );
}

function when2(iso: string): string {
  return new Date(iso).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
}

/** The agent's pending changes: what each did, where, accept or reject. */
export function ReviewPanel(props: {
  batches: AgentBatch[];
  onAccept: (id?: string) => void;
  onReject: (id?: string) => void;
  onJump: (cell: { sheet: string; row: number; col: number }) => void;
  onClose: () => void;
}) {
  const pending = props.batches.filter((b) => b.status === "pending");
  const accepted = props.batches.length - pending.length;
  return (
    <Panel role="dialog" aria-label="AI changes">
      <PanelHead>
        <strong style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
          <Sparkles size={14} /> AI changes{pending.length ? ` (${pending.length})` : ""}
        </strong>
        <IconBtn title="Close" onClick={props.onClose}>
          <X size={14} />
        </IconBtn>
      </PanelHead>
      {pending.length > 0 ? (
        <>
          <PanelHint>Highlighted in the sheet. Nothing is saved until you accept.</PanelHint>
          <BarRow>
            <TextBtn title="Keep every pending change" onClick={() => props.onAccept()}>
              <Check size={12} /> Accept all
            </TextBtn>
            <TextBtn title="Undo every pending change" onClick={() => props.onReject()}>
              <X size={12} /> Reject all
            </TextBtn>
          </BarRow>
        </>
      ) : (
        <PanelHint>Nothing waiting for review.{accepted ? ` ${accepted} change${accepted === 1 ? "" : "s"} accepted.` : ""}</PanelHint>
      )}
      {pending.map((b) => {
        const d = describeBatch(b);
        const first = cellSpan(b.cells).first;
        // Rejecting undoes this change and every one made after it.
        const later = props.batches.length - 1 - props.batches.indexOf(b);
        return (
          <Change key={b.id}>
            <ChangeBody
              role="button"
              tabIndex={0}
              title={first ? "Show in the sheet" : undefined}
              onClick={() => first && props.onJump(first)}
              onKeyDown={(e) => e.key === "Enter" && first && props.onJump(first)}
            >
              <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
                <strong>{d.title}</strong>
                <MenuHint>{when2(b.created_at)}</MenuHint>
              </div>
              {d.where && <MenuHint>{d.where}</MenuHint>}
              {d.lines.map((l, i) => (
                <Detail key={i}>{l}</Detail>
              ))}
              {d.more > 0 && <MenuHint>and {d.more} more</MenuHint>}
            </ChangeBody>
            <BarRow>
              <TextBtn title="Keep this change" onClick={() => props.onAccept(b.id)}>
                <Check size={12} /> Accept
              </TextBtn>
              <TextBtn
                title={later ? `Undo this change and the ${later} made after it` : "Undo this change"}
                onClick={() => props.onReject(b.id)}
              >
                <X size={12} /> Reject{later ? ` (+${later} after)` : ""}
              </TextBtn>
            </BarRow>
          </Change>
        );
      })}
    </Panel>
  );
}

const BarRow = styled.div`
  display: flex;
  gap: 6px;
  padding: 0 12px 8px;
`;

const Change = styled.div`
  border-top: 1px solid #eef0f2;
  padding-top: 8px;
`;

const ChangeBody = styled.div`
  padding: 0 12px 6px;
  cursor: pointer;
  display: flex;
  flex-direction: column;
  gap: 2px;
  &:hover { background: #f7f8fa; }
`;

const Detail = styled.div`
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 11px;
  color: #444;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
`;

const ChromeWrap = styled.div`
  position: relative;
  flex: none;
`;

const Row = styled.span`
  display: inline-flex;
  align-items: center;
  gap: 8px;
`;

const Status = styled.span<{ $warn?: boolean }>`
  font-size: 12px;
  color: ${(p) => (p.$warn ? "#8a5a00" : "#5f6368")};
  white-space: nowrap;
  padding: 0 4px;
`;

const LinkBtn = styled.button`
  background: none;
  border: 0;
  padding: 0;
  color: #2563b0;
  cursor: pointer;
  font-size: 12px;
  text-decoration: underline;
`;

const Notice = styled.div`
  position: absolute;
  right: 12px;
  top: calc(100% + 6px);
  z-index: 60;
  background: #333537;
  color: #fff;
  font-size: 12px;
  padding: 6px 10px;
  border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
  pointer-events: none;
`;

const Panel = styled.div`
  position: absolute;
  top: 0;
  right: 0;
  bottom: 0;
  width: 280px;
  z-index: 40;
  background: #fff;
  border-left: 1px solid #d2d6db;
  box-shadow: -6px 0 18px rgba(0, 0, 0, 0.08);
  overflow-y: auto;
  font-size: 12px;
  color: #333537;
`;

const PanelHead = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 10px 4px 12px;
  font-size: 13px;
`;

const PanelHint = styled.div`
  color: #9aa0a6;
  font-size: 11px;
  padding: 0 12px 8px;
`;

const VersionRow = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 8px 12px;
  border-top: 1px solid #eef0f2;
`;
