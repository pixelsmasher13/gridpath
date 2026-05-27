import React from "react";
import styled from "styled-components";
import { Clock, Archive, Trash2, FilePlus, FolderOpen } from "lucide-react";
import type { SessionRow } from "../sessionDb";

const Panel = styled.div`
  width: 100%;
  height: 100%;
  background: #1a1a1a;
  border-right: 1px solid #2a2a2a;
  display: flex;
  flex-direction: column;
  color: #d4d4d4;
  overflow: hidden;
`;

const Header = styled.div`
  padding: 8px 8px 8px 12px;
  border-bottom: 1px solid #2a2a2a;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: #b3b3b3;
  display: flex;
  align-items: center;
  gap: 6px;
`;

const IconBtnGroup = styled.div`
  margin-left: auto;
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

const List = styled.div`
  flex: 1;
  overflow-y: auto;
  padding: 6px;

  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
`;

const Item = styled.div<{ $active: boolean }>`
  display: flex;
  flex-direction: column;
  gap: 2px;
  padding: 8px 10px;
  border-radius: 5px;
  cursor: pointer;
  margin-bottom: 2px;
  background: ${(p) => (p.$active ? "#252526" : "transparent")};
  border-left: 2px solid ${(p) => (p.$active ? "#3363AD" : "transparent")};
  &:hover { background: #232323; }
  position: relative;
`;

const Name = styled.div<{ $placeholder?: boolean }>`
  font-size: 12px;
  font-weight: 500;
  color: ${(p) => (p.$placeholder ? "#6f6f6f" : "#e4e4e4")};
  font-style: ${(p) => (p.$placeholder ? "italic" : "normal")};
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Fineprint = styled.div`
  font-size: 10px;
  color: #7c7c7c;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Empty = styled.div`
  color: #6f6f6f;
  font-size: 12px;
  padding: 12px;
  text-align: center;
`;

const IconBtn = styled.button`
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

const Actions = styled.div`
  position: absolute;
  right: 6px;
  top: 6px;
  display: flex;
  gap: 2px;
`;

function formatAgo(ts: string): string {
  // ts is the SQLite datetime('now') string, e.g. "2026-05-19 21:00:00".
  // We treat it as UTC and compute a relative time from now.
  const d = new Date(ts.replace(" ", "T") + "Z");
  if (isNaN(d.getTime())) return "";
  const s = Math.round((Date.now() - d.getTime()) / 1000);
  if (s < 60) return "just now";
  const m = Math.round(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.round(m / 60);
  if (h < 48) return `${h}h ago`;
  return `${Math.round(h / 24)}d ago`;
}

interface Props {
  sessions: SessionRow[];
  activePath: string | null;
  onOpen: (s: SessionRow) => void;
  onArchive: (id: string) => void;
  onDelete: (id: string) => void;
  onNewBlank: () => void;
  onOpenExisting: () => void;
}

export const SessionSidebar: React.FC<Props> = ({
  sessions, activePath, onOpen, onArchive, onDelete, onNewBlank, onOpenExisting,
}) => {
  return (
    <Panel>
      <Header>
        <Clock size={11} />
        Recent sessions
        <IconBtnGroup>
          <HeaderIconBtn onClick={onNewBlank} title="New blank workbook (⌘N)">
            <FilePlus size={12} />
          </HeaderIconBtn>
          <HeaderIconBtn onClick={onOpenExisting} title="Open existing xlsx (⌘T)">
            <FolderOpen size={12} />
          </HeaderIconBtn>
        </IconBtnGroup>
      </Header>
      <List>
        {sessions.length === 0 && (
          <Empty>No recent sessions yet. Open an xlsx to start.</Empty>
        )}
        {sessions.map((s) => {
          const isUntitled = s.workbook_path.startsWith("untitled-");
          const displayName = s.name || (isUntitled ? "Untitled draft" : "Untitled session");
          const fineprintLeft = isUntitled ? "Unsaved" : s.workbook_path.split("/").pop();
          return (
          <Item
            key={s.id}
            $active={s.workbook_path === activePath}
            onClick={() => onOpen(s)}
            title={`${displayName}\n${s.workbook_path}`}
          >
            <Name $placeholder={!s.name}>{displayName}</Name>
            <Fineprint>
              {fineprintLeft} · {formatAgo(s.updated_at)}
            </Fineprint>
            <Actions>
              <IconBtn
                onClick={(e) => { e.stopPropagation(); onArchive(s.id); }}
                title="Archive"
              >
                <Archive size={12} />
              </IconBtn>
              <IconBtn
                onClick={(e) => { e.stopPropagation(); onDelete(s.id); }}
                title="Delete"
              >
                <Trash2 size={12} />
              </IconBtn>
            </Actions>
          </Item>
          );
        })}
      </List>
    </Panel>
  );
};
