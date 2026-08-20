import React, { useEffect, useMemo, useState } from "react";
import styled from "styled-components";
import { X, Check, ArrowRight, Crosshair, Layers, Trash2 } from "lucide-react";
import type { ChangeBatch } from "../types";
import { groupBatchForReview, type ReviewCell, type ReviewGroups } from "../agent/reviewGroups";

/**
 * Full-screen review of an agent batch: regions grouped by sheet + row band,
 * old → new per cell, jump-to-cell, and per-cell / per-region reject while
 * the batch is still pending.
 */

const Backdrop = styled.div`
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.6);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 180;
`;

const Surface = styled.div`
  width: min(1060px, calc(100vw - 48px));
  height: calc(100vh - 72px);
  background: #1c1c1d;
  border: 1px solid #333;
  border-radius: 10px;
  color: #d4d4d4;
  box-shadow: 0 18px 64px rgba(0, 0, 0, 0.6);
  display: flex;
  flex-direction: column;
  overflow: hidden;
`;

const Header = styled.div`
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 12px 16px;
  border-bottom: 1px solid #2a2a2a;
  flex-shrink: 0;
`;

const Title = styled.div`
  font-size: 13px;
  font-weight: 600;
  color: #ececec;
`;

const Counts = styled.div`
  font-size: 12px;
  color: #8c8c8c;
`;

const HeaderActions = styled.div`
  margin-left: auto;
  display: flex;
  align-items: center;
  gap: 8px;
`;

const AcceptAllBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 5px 12px;
  font-size: 12px;
  border-radius: 5px;
  border: 1px solid #16a34a;
  background: #16a34a;
  color: #fff;
  font-weight: 500;
  cursor: pointer;
  &:hover { background: #15803d; }
`;

const RejectAllBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  padding: 5px 12px;
  font-size: 12px;
  border-radius: 5px;
  border: 1px solid #444;
  background: transparent;
  color: #b3b3b3;
  cursor: pointer;
  &:hover { background: #2a2a2a; color: #fca5a5; border-color: #7f1d1d; }
`;

const CloseBtn = styled.button`
  background: transparent;
  border: 0;
  color: #9b9b9b;
  cursor: pointer;
  padding: 5px;
  border-radius: 4px;
  display: flex;
  &:hover { background: rgba(255, 255, 255, 0.06); color: #e4e4e4; }
`;

const BodyRow = styled.div`
  flex: 1;
  display: flex;
  min-height: 0;
`;

const Rail = styled.div`
  width: 190px;
  flex-shrink: 0;
  border-right: 1px solid #2a2a2a;
  padding: 10px 8px;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 2px;
  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
`;

const RailItem = styled.button<{ $active: boolean }>`
  display: flex;
  align-items: center;
  gap: 6px;
  width: 100%;
  flex-shrink: 0; /* same clipping trap as RegionCard — Rail is a flex column */
  text-align: left;
  padding: 6px 9px;
  border-radius: 5px;
  border: 0;
  cursor: pointer;
  font-size: 12px;
  background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.22)" : "transparent")};
  color: ${(p) => (p.$active ? "#cfe1f8" : "#b3b3b3")};
  &:hover { background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.28)" : "rgba(255,255,255,0.05)")}; }
`;

const RailCount = styled.span`
  margin-left: auto;
  font-size: 11px;
  color: #7c7c7c;
`;

const Content = styled.div`
  flex: 1;
  overflow-y: auto;
  padding: 12px 16px 20px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  &::-webkit-scrollbar { width: 8px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 4px; }
`;

const RegionCard = styled.div`
  border: 1px solid #2d2d2d;
  border-radius: 8px;
  background: #202021;
  overflow: hidden;
  /* Content is a flex column; without this, cards SHRINK to fit the
     viewport instead of overflowing it, and overflow:hidden silently
     clips the cell rows — a big batch showed ~25 rows with no scrollbar. */
  flex-shrink: 0;
`;

const RegionHeader = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  background: #242426;
  border-bottom: 1px solid #2d2d2d;
  position: sticky;
  top: 0;
  /* Rows are later siblings — without a stacking context they'd paint
     over the stuck header while scrolling underneath it. */
  z-index: 1;
`;

const RegionRange = styled.span`
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 12px;
  color: #93c5fd;
`;

const RegionSheet = styled.span`
  font-size: 11px;
  color: #8c8c8c;
`;

const RegionLabel = styled.span`
  font-size: 12px;
  color: #d9d9d9;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const RegionRejectBtn = styled.button`
  margin-left: auto;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 3px 8px;
  font-size: 11px;
  border-radius: 4px;
  border: 1px solid #3a3a3a;
  background: transparent;
  color: #9b9b9b;
  cursor: pointer;
  flex-shrink: 0;
  &:hover { background: rgba(127, 29, 29, 0.15); color: #fca5a5; border-color: #7f1d1d; }
`;

const CellRow = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 12px;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 12px;
  line-height: 1.5;
  cursor: pointer;
  border-bottom: 1px solid #262627;
  &:last-child { border-bottom: none; }
  &:hover { background: #26262a; }
  &:hover .cell-reject { opacity: 1; }
`;

const Addr = styled.span`
  width: 58px;
  flex-shrink: 0;
  color: #93c5fd;
`;

/**
 * Per-cell context (column header / row label) — only what VARIES within
 * the region; shared context lives in the region header instead.
 */
const Ctx = styled.span`
  font-family: -apple-system, "SF Pro Text", "Inter", system-ui, sans-serif;
  font-size: 11px;
  color: #8a97a6;
  max-width: 200px;
  flex-shrink: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

/** Context shared by every cell in the region (e.g. one column's header). */
const SharedCtx = styled.span`
  font-size: 11px;
  color: #8a97a6;
  padding: 1px 7px;
  border-radius: 9px;
  background: rgba(255, 255, 255, 0.05);
  border: 1px solid #333;
  flex-shrink: 0;
  max-width: 200px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const Old = styled.span`
  color: #f0a5a5;
  text-decoration: line-through;
  opacity: 0.85;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 34%;
`;

const New = styled.span`
  color: #9fe3b4;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  flex: 1;
  min-width: 0;
`;

const JumpHint = styled.span`
  display: inline-flex;
  color: #5c6b7d;
  flex-shrink: 0;
`;

const CellRejectBtn = styled.button`
  opacity: 0;
  display: inline-flex;
  align-items: center;
  padding: 2px 5px;
  border-radius: 4px;
  border: 1px solid transparent;
  background: transparent;
  color: #8c8c8c;
  cursor: pointer;
  flex-shrink: 0;
  transition: opacity 0.1s;
  &:hover { background: rgba(127, 29, 29, 0.2); color: #fca5a5; }
`;

const OthersCard = styled(RegionCard)``;

const OtherRow = styled.div`
  padding: 5px 12px;
  font-size: 12px;
  color: #c9c9c9;
  border-bottom: 1px solid #262627;
  &:last-child { border-bottom: none; }
`;

const EmptyNote = styled.div`
  color: #6f6f6f;
  font-size: 13px;
  text-align: center;
  padding: 40px 0;
`;

export interface ReviewModalProps {
  batch: ChangeBatch;
  filename: string;
  /** Reads a human label for a row (leftmost text cell). */
  labelOf: (sheet: string, row: number) => string | null;
  /** Reads a column header (topmost text/date cell of the column). */
  headerOf: (sheet: string, col: number) => string | null;
  onClose: () => void;
  /** Accept the whole batch (only offered while pending). */
  onAcceptAll: () => void;
  /** Reject the whole batch (only offered while pending). */
  onRejectAll: () => void;
  /** Minimize the modal and bring the cell into view in the grid. */
  onJump: (sheet: string, row: number, col: number) => void;
  /** Reject a subset of cells (per-cell / per-region). Pending batches only. */
  onRejectCells: (cells: Array<{ sheet: string; row: number; col: number }>) => void;
}

export const ReviewModal: React.FC<ReviewModalProps> = ({
  batch, filename, labelOf, headerOf, onClose, onAcceptAll, onRejectAll, onJump, onRejectCells,
}) => {
  const [sheetFilter, setSheetFilter] = useState<string | null>(null);

  useEffect(() => {
    const onEsc = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onEsc);
    return () => window.removeEventListener("keydown", onEsc);
  }, [onClose]);

  const groups: ReviewGroups = useMemo(
    () => groupBatchForReview(batch.mutations, labelOf, headerOf),
    [batch.mutations, labelOf, headerOf],
  );

  const sheets = Object.keys(groups.countsBySheet);
  const pending = batch.status === "pending";
  const visibleRegions = sheetFilter
    ? groups.regions.filter((r) => r.sheet === sheetFilter)
    : groups.regions;
  const showOthers = !sheetFilter && groups.others.length > 0;

  return (
    <Backdrop onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <Surface>
        <Header>
          <Title>Review changes</Title>
          <Counts>
            {groups.totalCells} cell edit{groups.totalCells === 1 ? "" : "s"}
            {groups.others.length > 0 && ` · ${groups.others.length} other`}
            {" · "}{filename}
          </Counts>
          <HeaderActions>
            {pending && (
              <>
                <AcceptAllBtn onClick={onAcceptAll}><Check size={12} /> Accept all</AcceptAllBtn>
                <RejectAllBtn onClick={onRejectAll}><X size={12} /> Reject all</RejectAllBtn>
              </>
            )}
            <CloseBtn onClick={onClose} title="Close (Esc)"><X size={15} /></CloseBtn>
          </HeaderActions>
        </Header>
        <BodyRow>
          <Rail>
            <RailItem $active={sheetFilter === null} onClick={() => setSheetFilter(null)}>
              <Layers size={12} />
              All sheets
              <RailCount>{groups.totalCells}</RailCount>
            </RailItem>
            {sheets.map((s) => (
              <RailItem key={s} $active={sheetFilter === s} onClick={() => setSheetFilter(s)}>
                {s}
                <RailCount>{groups.countsBySheet[s]}</RailCount>
              </RailItem>
            ))}
          </Rail>
          <Content>
            {visibleRegions.length === 0 && !showOthers && (
              <EmptyNote>No cell edits{sheetFilter ? ` on ${sheetFilter}` : ""}.</EmptyNote>
            )}
            {visibleRegions.map((region) => {
              // Show per-cell only the context that VARIES within the region:
              // a column run shares one header (goes in the region header, row
              // labels per cell); a row fill shares one row label (already the
              // region label, headers per cell).
              const headers = new Set(region.cells.map((c) => c.header).filter(Boolean));
              const rowLabels = new Set(region.cells.map((c) => c.rowLabel).filter(Boolean));
              const sharedHeader = headers.size === 1 ? [...headers][0] : null;
              const perCellHeader = headers.size > 1;
              const perCellRowLabel = rowLabels.size > 1;
              return (
                <RegionCard key={`${region.sheet}:${region.range}`}>
                  <RegionHeader>
                    <RegionRange>{region.range}</RegionRange>
                    <RegionSheet>{region.sheet}</RegionSheet>
                    {region.label && <RegionLabel>— {region.label}</RegionLabel>}
                    {sharedHeader && <SharedCtx title={`Column header shared by every cell in this region`}>{sharedHeader}</SharedCtx>}
                    <span style={{ marginLeft: "auto", fontSize: 11, color: "#7c7c7c", flexShrink: 0 }}>
                      {region.cells.length} cell{region.cells.length === 1 ? "" : "s"}
                    </span>
                    {pending && (
                      <RegionRejectBtn
                        onClick={() => onRejectCells(region.cells.map(({ sheet, row, col }) => ({ sheet, row, col })))}
                        title="Revert every cell in this region"
                      >
                        <Trash2 size={11} /> Reject region
                      </RegionRejectBtn>
                    )}
                  </RegionHeader>
                  {region.cells.map((c) => (
                    <ReviewCellRow
                      key={`${c.row},${c.col}`}
                      cell={c}
                      ctx={[perCellHeader ? c.header : null, perCellRowLabel ? c.rowLabel : null]
                        .filter(Boolean)
                        .join(" · ") || null}
                      pending={pending}
                      onJump={onJump}
                      onRejectCells={onRejectCells}
                    />
                  ))}
                </RegionCard>
              );
            })}
            {showOthers && (
              <OthersCard>
                <RegionHeader>
                  <RegionLabel>Structure &amp; formatting</RegionLabel>
                  <span style={{ marginLeft: "auto", fontSize: 11, color: "#7c7c7c" }}>
                    {groups.others.length}
                  </span>
                </RegionHeader>
                {groups.others.map((o, i) => (
                  <OtherRow key={i}>{o.desc}</OtherRow>
                ))}
              </OthersCard>
            )}
          </Content>
        </BodyRow>
      </Surface>
    </Backdrop>
  );
};

const ReviewCellRow: React.FC<{
  cell: ReviewCell;
  /** Pre-joined varying context for this cell ("FY2026E · ARPU growth"), or null. */
  ctx: string | null;
  pending: boolean;
  onJump: (sheet: string, row: number, col: number) => void;
  onRejectCells: (cells: Array<{ sheet: string; row: number; col: number }>) => void;
}> = ({ cell, ctx, pending, onJump, onRejectCells }) => (
  <CellRow
    onClick={() => onJump(cell.sheet, cell.row, cell.col)}
    title={`Jump to ${cell.sheet}!${cell.addr}`}
  >
    <Addr>{cell.addr}</Addr>
    {ctx && <Ctx title={ctx}>{ctx}</Ctx>}
    <Old>{cell.oldValue === "" ? "(empty)" : cell.oldValue}</Old>
    <ArrowRight size={10} style={{ color: "#555", flexShrink: 0 }} />
    <New>{cell.newValue === "" ? "(cleared)" : cell.newValue}</New>
    <JumpHint><Crosshair size={11} /></JumpHint>
    {pending && (
      <CellRejectBtn
        className="cell-reject"
        title="Revert this cell only"
        onClick={(e) => {
          e.stopPropagation();
          onRejectCells([{ sheet: cell.sheet, row: cell.row, col: cell.col }]);
        }}
      >
        <X size={11} />
      </CellRejectBtn>
    )}
  </CellRow>
);
