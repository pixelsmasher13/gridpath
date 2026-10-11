import React, { useMemo, useState } from "react";
import styled from "styled-components";
import { History, ChevronDown, ChevronRight } from "lucide-react";
import type { WorkbookJournal } from "../journal";
import { describeJournal } from "../journal";
import type { ChangeBatch, UniverMutation } from "../../../../packages/gridpath/src/core/types";

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
  width: 520px;
  max-width: calc(100vw - 40px);
  max-height: calc(100vh - 80px);
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
const Body = styled.div`
  padding: 14px 16px 10px;
  display: flex;
  flex-direction: column;
  gap: 10px;
  overflow-y: auto;
  min-height: 0;
`;
const Message = styled.div`
  font-size: 13px;
  line-height: 1.5;
`;
const Muted = styled.div`
  font-size: 12px;
  color: #9b9b9b;
  line-height: 1.5;
`;
const Toggle = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  background: transparent;
  border: 0;
  color: #9ec1f5;
  font-size: 12px;
  cursor: pointer;
  padding: 0;
  &:hover { text-decoration: underline; }
`;
const BatchList = styled.ol`
  margin: 0;
  padding: 0 0 0 18px;
  display: flex;
  flex-direction: column;
  gap: 8px;
  font-size: 12px;
`;
const BatchTitle = styled.div`
  color: #e4e4e4;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
`;
const BatchMeta = styled.div`
  color: #9b9b9b;
`;
const Footer = styled.div`
  display: flex;
  justify-content: flex-end;
  align-items: center;
  gap: 8px;
  padding: 12px 14px;
  border-top: 1px solid #2a2a2a;
`;
const Btn = styled.button<{ $variant?: "primary" | "danger" | "ghost" }>`
  background: ${(p) => (p.$variant === "primary" ? "#3363AD" : "transparent")};
  color: ${(p) => (p.$variant === "primary" ? "#fff" : p.$variant === "danger" ? "#fca5a5" : "#d4d4d4")};
  border: 1px solid ${(p) => (p.$variant === "primary" ? "#3363AD" : p.$variant === "danger" ? "#7f1d1d" : "#333")};
  border-radius: 5px;
  padding: 7px 14px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:hover:not(:disabled) {
    background: ${(p) => (p.$variant === "primary" ? "#4275c4" : p.$variant === "danger" ? "rgba(127, 29, 29, 0.18)" : "#2a2a2a")};
  }
`;

/** What one batch did, in a line: counts by kind plus the sheets it touched. */
function summarizeBatch(b: ChangeBatch): { counts: string; sheets: string[] } {
  const n = { cells: 0, formats: 0, sizes: 0, structure: 0, sheets: 0, other: 0 };
  const touched = new Set<string>();
  for (const m of b.mutations as UniverMutation[]) {
    const t = (m as { type: string }).type;
    const anyM = m as unknown as { sheet?: string; address?: { sheet?: string }; name?: string; old_name?: string };
    const s = anyM.sheet ?? anyM.address?.sheet ?? (t === "create_sheet" || t === "delete_sheet" ? anyM.name : t === "rename_sheet" ? anyM.old_name : undefined);
    if (s) touched.add(s);
    if (t === "set_cell" || t === "set_range" || t === "clear_range") n.cells++;
    else if (t === "set_format") n.formats += ((m as unknown as { cells?: unknown[] }).cells?.length ?? 1);
    else if (t === "set_column_width" || t === "set_row_height") n.sizes++;
    else if (t === "create_sheet" || t === "delete_sheet" || t === "rename_sheet") n.sheets++;
    else if (/^(insert|delete)_(rows|columns)$|^(un)?merge_cells$|^(un)?freeze_panes$|^(hide|show)_(rows|columns)$/.test(t)) n.structure++;
    else n.other++;
  }
  const parts: string[] = [];
  if (n.cells) parts.push(`${n.cells.toLocaleString()} cell write${n.cells === 1 ? "" : "s"}`);
  if (n.formats) parts.push(`${n.formats.toLocaleString()} formatted cell${n.formats === 1 ? "" : "s"}`);
  if (n.sizes) parts.push(`${n.sizes} width/height change${n.sizes === 1 ? "" : "s"}`);
  if (n.structure) parts.push(`${n.structure} row/column/merge/freeze change${n.structure === 1 ? "" : "s"}`);
  if (n.sheets) parts.push(`${n.sheets} sheet change${n.sheets === 1 ? "" : "s"}`);
  if (n.other) parts.push(`${n.other} other`);
  return { counts: parts.join(", ") || "no edits", sheets: [...touched] };
}

interface Props {
  open: boolean;
  filename: string;
  journal: WorkbookJournal | null;
  onRestore: () => void;
  onDiscard: () => void;
}

/**
 * Offer to bring back the unsaved work a previous session left for this
 * file. Restore puts it on the grid as reviewable batches; Discard deletes
 * the record (asked twice, since there is no undo). No close button: the
 * record stays until one is chosen.
 */
export const RecoverModal: React.FC<Props> = ({ open, filename, journal, onRestore, onDiscard }) => {
  const [details, setDetails] = useState(false);
  const [confirmDiscard, setConfirmDiscard] = useState(false);
  const batches = useMemo(() => (journal?.batches ?? []).map((b) => ({ b, ...summarizeBatch(b) })), [journal]);
  if (!open || !journal) return null;
  const when = new Date(journal.updatedAt).toLocaleString();
  const shown = batches.slice(0, 8);
  return (
    <Backdrop>
      <Modal onClick={(e) => e.stopPropagation()}>
        <Header>
          <History size={14} color="#f59e0b" />
          Unsaved changes from a previous session
        </Header>
        <Body>
          <Message>
            <b>{filename}</b> has unsaved work from {when}: {describeJournal(journal)}.
          </Message>
          <Muted>
            Restore puts the changes back on the grid, where each agent batch can still be reviewed and rejected before you save.
            Discard deletes the record for good; the file on disk is not touched either way.
          </Muted>
          {batches.length > 0 && (
            <Toggle onClick={() => setDetails((d) => !d)}>
              {details ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              {details ? "Hide details" : `Show details (${batches.length} agent batch${batches.length === 1 ? "" : "es"})`}
            </Toggle>
          )}
          {details && (
            <BatchList>
              {shown.map(({ b, counts, sheets }) => (
                <li key={b.id}>
                  <BatchTitle title={b.prompt}>{b.prompt || "(no prompt)"}</BatchTitle>
                  <BatchMeta>
                    {counts}
                    {sheets.length > 0 && ` · ${sheets.slice(0, 4).join(", ")}${sheets.length > 4 ? ` +${sheets.length - 4}` : ""}`}
                    {b.created_at && ` · ${new Date(b.created_at).toLocaleTimeString()}`}
                  </BatchMeta>
                </li>
              ))}
              {batches.length > shown.length && <Muted>… {batches.length - shown.length} more</Muted>}
            </BatchList>
          )}
        </Body>
        <Footer>
          {confirmDiscard ? (
            <>
              <Muted style={{ marginRight: "auto" }}>Delete these changes permanently?</Muted>
              <Btn onClick={() => setConfirmDiscard(false)}>Keep</Btn>
              <Btn $variant="danger" onClick={onDiscard}>Discard for good</Btn>
            </>
          ) : (
            <>
              <Btn $variant="danger" onClick={() => setConfirmDiscard(true)}>Discard</Btn>
              <Btn $variant="primary" onClick={onRestore} autoFocus>Restore</Btn>
            </>
          )}
        </Footer>
      </Modal>
    </Backdrop>
  );
};
