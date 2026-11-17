import React, { useEffect, useRef, useState } from "react";
import styled from "styled-components";
import { Layers, FilePlus, Plus, X, Unlink, ChevronDown, Check } from "lucide-react";
import { workbookName, type WorkspaceMember } from "../workspaceApi";

/**
 * The workspace: the workbooks that are worked on together, and what reads
 * what. A workbook here does not need a tab: one that another workbook
 * reads is loaded in the background, and a click gives it a tab.
 *
 * There can be several workspaces. The header names the one shown and
 * opens the list of the others (switch, new, rename, delete). A workspace
 * is saved as it changes; there is no save step.
 *
 * Files and folders dropped anywhere on the window are added to the
 * workspace shown (the screen owns the drop; `dropping` is true while
 * something is held over the window).
 */

const Panel = styled.div<{ $dropping: boolean }>`
  width: 100%;
  background: #1a1a1a;
  border-right: 1px solid #2a2a2a;
  border-bottom: 1px solid #2a2a2a;
  display: flex;
  flex-direction: column;
  color: #d4d4d4;
  min-height: 0;
  max-height: 45%;
  flex-shrink: 0;
  position: relative;
  outline: ${(p) => (p.$dropping ? "1px dashed #3363AD" : "none")};
  outline-offset: -3px;
`;

const Header = styled.div`
  padding: 6px 8px 6px 10px;
  border-bottom: 1px solid #2a2a2a;
  font-size: 11px;
  color: #b3b3b3;
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
`;

const Switcher = styled.button`
  flex: 1;
  min-width: 0;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: transparent;
  border: 1px solid transparent;
  border-radius: 5px;
  color: #d4d4d4;
  padding: 4px 6px;
  cursor: pointer;
  text-align: left;
  &:hover { background: #232323; border-color: #333; }
`;

const SwitcherName = styled.span`
  flex: 1;
  min-width: 0;
  font-size: 12px;
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const NameInput = styled.input`
  flex: 1;
  min-width: 0;
  background: #111;
  border: 1px solid #3363AD;
  border-radius: 5px;
  color: #e4e4e4;
  font-size: 12px;
  font-weight: 600;
  padding: 4px 6px;
  outline: none;
`;

const IconBtnGroup = styled.div`
  flex-shrink: 0;
  display: inline-flex;
  gap: 4px;
`;

const HeaderIconBtn = styled.button`
  background: transparent;
  border: 1px solid #333;
  color: #d4d4d4;
  border-radius: 5px;
  padding: 4px 6px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  &:hover { background: #2a2a2a; border-color: #3363AD; color: #fff; }
`;

const Menu = styled.div`
  position: absolute;
  top: 34px;
  left: 8px;
  right: 8px;
  z-index: 30;
  background: #222;
  border: 1px solid #383838;
  border-radius: 6px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.45);
  padding: 4px;
  max-height: 320px;
  overflow-y: auto;
`;

const MenuRow = styled.button<{ $danger?: boolean }>`
  width: 100%;
  display: flex;
  align-items: center;
  gap: 8px;
  background: transparent;
  border: 0;
  border-radius: 4px;
  color: ${(p) => (p.$danger ? "#e07a7a" : "#d4d4d4")};
  font-size: 12px;
  padding: 6px 8px;
  cursor: pointer;
  text-align: left;
  &:hover { background: #2e2e2e; }
`;

const MenuName = styled.span`
  flex: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const MenuCount = styled.span`
  flex-shrink: 0;
  color: #7c7c7c;
  font-size: 10px;
`;

const MenuSep = styled.div`
  height: 1px;
  background: #333;
  margin: 4px 0;
`;

const MenuNote = styled.div`
  color: #9b9b9b;
  font-size: 11px;
  line-height: 1.4;
  padding: 6px 8px 2px;
`;

const List = styled.div`
  overflow-y: auto;
  padding: 6px;

  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
`;

const Item = styled.div<{ $active: boolean; $dim: boolean }>`
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding: 7px 10px;
  border-radius: 5px;
  cursor: pointer;
  margin-bottom: 2px;
  background: ${(p) => (p.$active ? "#252526" : "transparent")};
  border-left: 2px solid ${(p) => (p.$active ? "#3363AD" : "transparent")};
  opacity: ${(p) => (p.$dim ? 0.6 : 1)};
  &:hover { background: #232323; }
  position: relative;
`;

const NameRow = styled.div`
  display: flex;
  align-items: center;
  gap: 6px;
  min-width: 0;
  padding-right: 16px;
`;

const Dot = styled.span<{ $color: string }>`
  flex-shrink: 0;
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: ${(p) => p.$color};
`;

const Name = styled.div`
  font-size: 12px;
  font-weight: 500;
  color: #e4e4e4;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Fineprint = styled.div<{ $tone?: "warn" | "bad" }>`
  font-size: 10px;
  color: ${(p) => (p.$tone === "bad" ? "#e07a7a" : p.$tone === "warn" ? "#d9a441" : "#7c7c7c")};
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Empty = styled.div`
  color: #6f6f6f;
  font-size: 12px;
  line-height: 1.45;
  padding: 10px 12px 12px;
`;

const RemoveBtn = styled.button`
  position: absolute;
  right: 6px;
  top: 6px;
  background: transparent;
  border: 0;
  color: #7c7c7c;
  padding: 2px;
  border-radius: 3px;
  cursor: pointer;
  visibility: hidden;
  &:hover { color: #d4d4d4; background: rgba(255,255,255,0.06); }

  ${Item}:hover & { visibility: visible; }
`;

const STATE_DOT: Record<WorkspaceMember["state"], string> = {
  open: "#4caf7d",
  loaded: "#3363AD",
  idle: "#4a4a4a",
  missing: "#e07a7a",
  failed: "#e07a7a",
};

const STATE_TEXT: Record<WorkspaceMember["state"], string> = {
  open: "open",
  loaded: "loaded, no tab",
  idle: "not loaded",
  missing: "file not found",
  failed: "could not be loaded",
};

const names = (paths: string[]) => paths.map(workbookName).join(", ");
const models = (n: number) => `${n} model${n === 1 ? "" : "s"}`;

/** One saved workspace, as the switcher lists it. */
export interface WorkspaceEntry {
  id: string;
  name: string;
  /** Models on it. */
  count: number;
}

interface Props {
  /** Every saved workspace, and the one shown. */
  workspaces: WorkspaceEntry[];
  activeId: string | null;
  onSwitch: (id: string) => void;
  /** Makes a new, empty workspace and shows it; returns its name. */
  onCreate: () => string;
  /** Renames the workspace shown. */
  onRename: (name: string) => void;
  /** Deletes the workspace shown (the list only; no file is touched). */
  onDelete: () => void;

  members: WorkspaceMember[];
  activePath: string | null;
  /** Something is being dragged over the window. */
  dropping: boolean;
  onOpen: (path: string) => void;
  onAdd: () => void;
  onNew: () => void;
  onRemove: (path: string) => void;
}

export const WorkspacePanel: React.FC<Props> = ({
  workspaces, activeId, onSwitch, onCreate, onRename, onDelete,
  members, activePath, dropping, onOpen, onAdd, onNew, onRemove,
}) => {
  const [menuOpen, setMenuOpen] = useState(false);
  const [confirmDelete, setConfirmDelete] = useState(false);
  /** The name being typed, while the workspace shown is being renamed. */
  const [draft, setDraft] = useState<string | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const active = workspaces.find((w) => w.id === activeId) ?? workspaces[0] ?? null;

  // A click outside the menu, or Escape, closes it.
  useEffect(() => {
    if (!menuOpen) return;
    const onDown = (ev: MouseEvent) => {
      if (!(ev.target as HTMLElement | null)?.closest?.("[data-workspace-menu]")) setMenuOpen(false);
    };
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === "Escape") setMenuOpen(false);
    };
    document.addEventListener("mousedown", onDown, true);
    document.addEventListener("keydown", onKey, true);
    return () => {
      document.removeEventListener("mousedown", onDown, true);
      document.removeEventListener("keydown", onKey, true);
    };
  }, [menuOpen]);
  useEffect(() => {
    if (!menuOpen) setConfirmDelete(false);
  }, [menuOpen]);

  const commitRename = () => {
    if (draft !== null && draft.trim()) onRename(draft);
    setDraft(null);
  };

  return (
    <Panel $dropping={dropping} data-workspace-panel="1" ref={rootRef}>
      <Header>
        <Layers size={12} style={{ flexShrink: 0 }} />
        {draft !== null ? (
          <NameInput
            autoFocus
            value={draft}
            placeholder="Workspace name"
            onFocus={(e) => e.currentTarget.select()}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") commitRename();
              else if (e.key === "Escape") setDraft(null);
              e.stopPropagation();
            }}
            onBlur={commitRename}
          />
        ) : (
          <Switcher
            data-workspace-menu="1"
            onClick={() => setMenuOpen((v) => !v)}
            title="Workspaces: switch to another, make a new one, rename or delete this one"
          >
            <SwitcherName>{active?.name ?? "Workspace"}</SwitcherName>
            <ChevronDown size={12} style={{ flexShrink: 0, color: "#7c7c7c" }} />
          </Switcher>
        )}
        <IconBtnGroup>
          <HeaderIconBtn onClick={onNew} title="New workbook in this workspace (saved to a file, so it can read the others)">
            <FilePlus size={12} />
          </HeaderIconBtn>
          <HeaderIconBtn onClick={onAdd} title="Add models to this workspace">
            <Plus size={12} />
          </HeaderIconBtn>
        </IconBtnGroup>
      </Header>
      {menuOpen && (
        <Menu data-workspace-menu="1">
          {workspaces.map((w) => (
            <MenuRow
              key={w.id}
              onClick={() => {
                setMenuOpen(false);
                if (w.id !== active?.id) onSwitch(w.id);
              }}
            >
              <span style={{ width: 12, flexShrink: 0, display: "inline-flex" }}>{w.id === active?.id && <Check size={12} />}</span>
              <MenuName>{w.name}</MenuName>
              <MenuCount>{models(w.count)}</MenuCount>
            </MenuRow>
          ))}
          <MenuSep />
          <MenuRow
            onClick={() => {
              setMenuOpen(false);
              setDraft(onCreate());
            }}
          >
            New workspace
          </MenuRow>
          {active && (
            <MenuRow
              onClick={() => {
                setMenuOpen(false);
                setDraft(active.name);
              }}
            >
              Rename “{active.name}”
            </MenuRow>
          )}
          {active && !confirmDelete && (
            <MenuRow $danger onClick={() => setConfirmDelete(true)}>
              Delete “{active.name}”…
            </MenuRow>
          )}
          {active && confirmDelete && (
            <>
              <MenuNote>
                Delete the workspace “{active.name}”? Its {models(active.count)} stay on disk; only the list goes.
              </MenuNote>
              <MenuRow
                $danger
                onClick={() => {
                  setMenuOpen(false);
                  onDelete();
                }}
              >
                Delete workspace
              </MenuRow>
              <MenuRow onClick={() => setConfirmDelete(false)}>Cancel</MenuRow>
            </>
          )}
        </Menu>
      )}
      {members.length === 0 ? (
        <Empty>{dropping ? "Drop to add these models." : "Drop models or a folder here to work on them together."}</Empty>
      ) : (
        <List>
          {members.map((m) => {
            const lines: Array<{ text: string; tone?: "warn" | "bad" }> = [];
            if (m.state === "failed" && m.error) lines.push({ text: m.error, tone: "bad" });
            if (m.reads.length) lines.push({ text: `reads ${names(m.reads)}` });
            if (m.readBy.length) lines.push({ text: `read by ${names(m.readBy)}` });
            if (m.broken.length) lines.push({ text: `broken: ${m.broken.join(", ")}`, tone: "bad" });
            if (m.needsSaving) lines.push({ text: "links moved, needs saving", tone: "warn" });
            const title = [
              m.path,
              STATE_TEXT[m.state],
              m.member ? null : "Not in this workspace: loaded because an open workbook reads it.",
              ...lines.map((l) => l.text),
            ]
              .filter(Boolean)
              .join("\n");
            return (
              <Item
                key={m.path}
                $active={m.path === activePath}
                $dim={!m.member}
                onClick={() => (m.state === "missing" ? undefined : onOpen(m.path))}
                title={title}
              >
                <NameRow>
                  <Dot $color={STATE_DOT[m.state]} />
                  <Name>{m.name}</Name>
                  {m.broken.length > 0 && <Unlink size={11} color="#e07a7a" style={{ flexShrink: 0 }} />}
                </NameRow>
                <Fineprint $tone={m.state === "missing" || m.state === "failed" ? "bad" : undefined}>
                  {STATE_TEXT[m.state]}
                  {m.member ? "" : " · linked"}
                </Fineprint>
                {lines.map((l, i) => (
                  <Fineprint key={i} $tone={l.tone}>
                    {l.text}
                  </Fineprint>
                ))}
                {m.member && (
                  <RemoveBtn
                    onClick={(e) => {
                      e.stopPropagation();
                      onRemove(m.path);
                    }}
                    title="Take out of this workspace (the file is not touched)"
                  >
                    <X size={12} />
                  </RemoveBtn>
                )}
              </Item>
            );
          })}
        </List>
      )}
    </Panel>
  );
};
