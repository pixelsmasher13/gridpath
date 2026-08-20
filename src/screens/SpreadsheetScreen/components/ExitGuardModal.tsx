import React, { useEffect } from "react";
import styled from "styled-components";
import { AlertTriangle, X } from "lucide-react";
import type { WorkbookTab } from "../state/tabs";

const Backdrop = styled.div`
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.55);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 200;
`;

const Modal = styled.div`
  width: 440px;
  max-width: calc(100vw - 40px);
  background: #1e1e1e;
  border: 1px solid #333;
  border-radius: 10px;
  color: #d4d4d4;
  box-shadow: 0 16px 56px rgba(0, 0, 0, 0.55);
  display: flex;
  flex-direction: column;
  overflow: hidden;
`;

const Header = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 14px 16px;
  border-bottom: 1px solid #2a2a2a;
  font-weight: 600;
  font-size: 13px;
`;

const HeaderIcon = styled.span`
  display: inline-flex;
  color: #f59e0b;
`;

const CloseBtn = styled.button`
  margin-left: auto;
  background: transparent;
  border: 0;
  color: #9b9b9b;
  cursor: pointer;
  padding: 4px;
  border-radius: 4px;
  display: flex;
  &:hover:not(:disabled) { background: rgba(255, 255, 255, 0.06); color: #e4e4e4; }
  &:disabled { opacity: 0.4; cursor: not-allowed; }
`;

const Body = styled.div`
  padding: 16px 16px 12px;
  display: flex;
  flex-direction: column;
  gap: 10px;
`;

const Message = styled.div`
  font-size: 13px;
  color: #d4d4d4;
  line-height: 1.5;
`;

const FileList = styled.ul`
  margin: 0;
  padding: 0 0 0 16px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  font-size: 12px;
  color: #b3b3b3;
  list-style: disc;
`;

const Footer = styled.div`
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 12px 14px;
  border-top: 1px solid #2a2a2a;
`;

const Btn = styled.button<{ $variant?: "primary" | "danger" | "ghost" }>`
  background: ${(p) =>
    p.$variant === "primary"
      ? "#3363AD"
      : p.$variant === "danger"
      ? "transparent"
      : "transparent"};
  color: ${(p) =>
    p.$variant === "primary"
      ? "#fff"
      : p.$variant === "danger"
      ? "#fca5a5"
      : "#d4d4d4"};
  border: 1px solid ${(p) =>
    p.$variant === "primary"
      ? "#3363AD"
      : p.$variant === "danger"
      ? "#7f1d1d"
      : "#333"};
  border-radius: 5px;
  padding: 7px 14px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:hover:not(:disabled) {
    background: ${(p) =>
      p.$variant === "primary"
        ? "#4275c4"
        : p.$variant === "danger"
        ? "rgba(127, 29, 29, 0.18)"
        : "#2a2a2a"};
  }
  &:disabled { opacity: 0.5; cursor: not-allowed; }
`;

interface Props {
  open: boolean;
  dirtyTabs: WorkbookTab[];
  saving: boolean;
  onSave: () => void;
  onDiscard: () => void;
  onCancel: () => void;
}

export const ExitGuardModal: React.FC<Props> = ({
  open, dirtyTabs, saving, onSave, onDiscard, onCancel,
}) => {
  // Esc cancels (matches the cancel button — safest default).
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape" && !saving) onCancel();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, saving, onCancel]);

  if (!open) return null;
  const n = dirtyTabs.length;

  return (
    <Backdrop onClick={() => { if (!saving) onCancel(); }}>
      <Modal onClick={(e) => e.stopPropagation()}>
        <Header>
          <HeaderIcon><AlertTriangle size={14} /></HeaderIcon>
          Unsaved changes
          <CloseBtn onClick={onCancel} disabled={saving} title="Dismiss (Esc)">
            <X size={14} />
          </CloseBtn>
        </Header>
        <Body>
          <Message>
            {n === 1
              ? "There's one workbook with unsaved changes:"
              : `There are ${n} workbooks with unsaved changes:`}
          </Message>
          <FileList>
            {dirtyTabs.map((t) => (
              <li key={t.id}>{t.filename}</li>
            ))}
          </FileList>
          <Message style={{ color: "#9b9b9b", fontSize: 12 }}>
            Save before quitting, or discard to lose them.
          </Message>
        </Body>
        <Footer>
          <Btn onClick={onCancel} disabled={saving}>Cancel</Btn>
          <Btn $variant="danger" onClick={onDiscard} disabled={saving}>Discard</Btn>
          <Btn $variant="primary" onClick={onSave} disabled={saving}>
            {saving ? "Saving…" : n === 1 ? "Save" : "Save All"}
          </Btn>
        </Footer>
      </Modal>
    </Backdrop>
  );
};
