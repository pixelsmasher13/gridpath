import React, { useState } from "react";
import styled from "styled-components";
import { X, Loader2 } from "lucide-react";
import type { WorkbookTab } from "../state/tabs";

const Bar = styled.div`
  display: flex;
  align-items: stretch;
  border-bottom: 1px solid #2a2a2a;
  background: #1e1e1e;
  height: 52px;
  overflow-x: auto;
  overflow-y: hidden;
  &::-webkit-scrollbar { height: 0; }
`;

const Tab = styled.div<{ $active: boolean }>`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 0 14px;
  min-width: 200px;
  max-width: 280px;
  background: ${(p) => (p.$active ? "#252526" : "#1e1e1e")};
  border-right: 1px solid #2a2a2a;
  border-top: 2px solid ${(p) => (p.$active ? "#3363AD" : "transparent")};
  color: ${(p) => (p.$active ? "#e4e4e4" : "#9b9b9b")};
  cursor: pointer;
  user-select: none;
  position: relative;
  flex-shrink: 0;

  &:hover {
    background: ${(p) => (p.$active ? "#252526" : "#222")};
    color: #d4d4d4;
  }
`;

const TabLabel = styled.div`
  display: flex;
  flex-direction: column;
  gap: 1px;
  flex: 1;
  min-width: 0;
`;

const Name = styled.span<{ $placeholder?: boolean }>`
  font-size: 12px;
  font-weight: 500;
  color: ${(p) => (p.$placeholder ? "#6f6f6f" : "inherit")};
  font-style: ${(p) => (p.$placeholder ? "italic" : "normal")};
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Fineprint = styled.span`
  font-size: 10px;
  color: #7c7c7c;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const NameInput = styled.input`
  background: #1a1a1a;
  border: 1px solid #444;
  color: #e4e4e4;
  font-size: 12px;
  border-radius: 3px;
  padding: 2px 4px;
  width: 100%;
  outline: none;
  &:focus { border-color: #3363AD; }
`;

const Trailing = styled.div`
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 2px;
  flex-shrink: 0;
`;

const DirtyDot = styled.span`
  width: 6px;
  height: 6px;
  border-radius: 50%;
  background: #d4d4d4;
`;

const CloseBtn = styled.button`
  background: transparent;
  border: 0;
  padding: 2px;
  border-radius: 3px;
  color: inherit;
  cursor: pointer;
  display: flex;
  align-items: center;
  visibility: hidden;
  &:hover { background: rgba(255,255,255,0.1); }

  ${Tab}:hover & { visibility: visible; }
`;

const Spin = styled(Loader2)`
  animation: spin 0.9s linear infinite;
  color: #3363AD;
  @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
`;

interface Props {
  tabs: WorkbookTab[];
  activeTabId: string | null;
  onActivate: (tabId: string) => void;
  onClose: (tabId: string) => void;
  onRename: (tabId: string, name: string) => void;
}

export const TabBar: React.FC<Props> = ({ tabs, activeTabId, onActivate, onClose, onRename }) => {
  const [renamingId, setRenamingId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState("");

  const startRename = (t: WorkbookTab) => {
    setRenamingId(t.id);
    setRenameDraft(t.name || "");
  };

  const commitRename = () => {
    if (renamingId) {
      const trimmed = renameDraft.trim();
      if (trimmed) onRename(renamingId, trimmed);
    }
    setRenamingId(null);
    setRenameDraft("");
  };

  return (
    <Bar>
      {tabs.map((t) => {
        const placeholder = !t.name;
        const isRenaming = renamingId === t.id;
        return (
          <Tab
            key={t.id}
            $active={t.id === activeTabId}
            onClick={() => !isRenaming && onActivate(t.id)}
            onDoubleClick={(e) => {
              e.stopPropagation();
              startRename(t);
            }}
            title={`${t.name || "Untitled session"} · ${t.path}`}
          >
            <TabLabel>
              {isRenaming ? (
                <NameInput
                  autoFocus
                  value={renameDraft}
                  onChange={(e) => setRenameDraft(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      commitRename();
                    } else if (e.key === "Escape") {
                      setRenamingId(null);
                      setRenameDraft("");
                    }
                    e.stopPropagation();
                  }}
                  onClick={(e) => e.stopPropagation()}
                />
              ) : (
                <Name $placeholder={placeholder}>
                  {t.name || "Untitled session"}
                </Name>
              )}
              <Fineprint>{t.filename}</Fineprint>
            </TabLabel>
            <Trailing>
              {t.agentRunning ? <Spin size={12} /> : t.dirty ? <DirtyDot /> : <span style={{ width: 6, height: 6 }} />}
              <CloseBtn
                onClick={(e) => {
                  e.stopPropagation();
                  onClose(t.id);
                }}
                aria-label="Close session"
              >
                <X size={12} />
              </CloseBtn>
            </Trailing>
          </Tab>
        );
      })}
    </Bar>
  );
};
