import React, { useEffect } from "react";
import styled from "styled-components";
import { AlertTriangle, X } from "lucide-react";

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
  width: 460px;
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
  &:hover { background: rgba(255, 255, 255, 0.06); color: #e4e4e4; }
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

const LossBox = styled.div`
  background: rgba(245, 158, 11, 0.08);
  border: 1px solid rgba(245, 158, 11, 0.35);
  border-radius: 6px;
  padding: 10px 12px;
  font-size: 12.5px;
  line-height: 1.5;
  color: #e8c88a;
`;

const Fine = styled.div`
  font-size: 12px;
  color: #9b9b9b;
  line-height: 1.5;
`;

const Footer = styled.div`
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 12px 14px;
  border-top: 1px solid #2a2a2a;
`;

const Btn = styled.button<{ $variant?: "primary" | "danger" }>`
  background: ${(p) => (p.$variant === "primary" ? "#3363AD" : "transparent")};
  color: ${(p) =>
    p.$variant === "primary" ? "#fff" : p.$variant === "danger" ? "#fca5a5" : "#d4d4d4"};
  border: 1px solid ${(p) =>
    p.$variant === "primary" ? "#3363AD" : p.$variant === "danger" ? "#7f1d1d" : "#333"};
  border-radius: 5px;
  padding: 7px 14px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:hover {
    background: ${(p) =>
      p.$variant === "primary"
        ? "#4275c4"
        : p.$variant === "danger"
        ? "rgba(127, 29, 29, 0.18)"
        : "#2a2a2a"};
  }
`;

interface Props {
  open: boolean;
  filename: string;
  /** Why the surgical path wasn't available — shown verbatim. */
  note: string | null;
  /** What this copy will be missing ("14 comments, 5 custom XML parts"). */
  losses: string;
  onConfirm: () => void;
  onCancel: () => void;
}

/**
 * Explicit consent gate for reduced-fidelity Save As. A copy written by the
 * fallback exporter drops content the surgical saver would have preserved —
 * that trade-off must be a decision the user makes, not a toast they might
 * miss.
 */
export const FidelitySaveModal: React.FC<Props> = ({
  open, filename, note, losses, onConfirm, onCancel,
}) => {
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onCancel();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onCancel]);

  if (!open) return null;

  return (
    <Backdrop onClick={onCancel}>
      <Modal onClick={(e) => e.stopPropagation()}>
        <Header>
          <HeaderIcon><AlertTriangle size={14} /></HeaderIcon>
          Save a reduced-fidelity copy?
          <CloseBtn onClick={onCancel} title="Cancel (Esc)">
            <X size={14} />
          </CloseBtn>
        </Header>
        <Body>
          <Message>
            {note ?? "The high-fidelity save path is unavailable for this edit"}.
            Writing <strong>{filename}</strong> through the fallback exporter
            will drop content from the copy:
          </Message>
          <LossBox>
            {losses
              ? `This copy will be missing: ${losses}.`
              : "No specific losses were detected, but the file will be fully re-serialized."}
          </LossBox>
          <Fine>
            The original file is not modified either way. You can also cancel
            and remove the unsupported edit to keep the save surgical.
          </Fine>
        </Body>
        <Footer>
          <Btn onClick={onCancel}>Cancel</Btn>
          <Btn $variant="danger" onClick={onConfirm}>Save reduced copy</Btn>
        </Footer>
      </Modal>
    </Backdrop>
  );
};
