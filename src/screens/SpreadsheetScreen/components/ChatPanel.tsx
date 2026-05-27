import React, { useEffect, useRef, useState } from "react";
import styled from "styled-components";
import { Send, Square, Loader2, ChevronDown, ChevronRight, Check, X, ArrowRight, User, Globe, Undo2, Redo2, MousePointerSquareDashed, Zap } from "lucide-react";
import type { ChangeBatch, UniverMutation } from "../types";
import type { WorkbookTab } from "../state/tabs";

const Panel = styled.div`
  display: flex;
  flex-direction: column;
  background: #252526;
  border-left: 1px solid #2a2a2a;
  height: 100%;
  overflow: hidden;
  color: #d4d4d4;
`;

const Header = styled.div`
  padding: 10px 14px;
  border-bottom: 1px solid #2a2a2a;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: #b3b3b3;
  display: flex;
  align-items: center;
  gap: 6px;
`;

const Thread = styled.div`
  flex: 1;
  overflow-y: auto;
  padding: 12px 12px 4px 12px;
  display: flex;
  flex-direction: column;
  gap: 12px;

  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
`;

const Empty = styled.div`
  color: #6f6f6f;
  font-size: 13px;
  padding: 16px;
  text-align: center;
  line-height: 1.5;
`;

const Bubble = styled.div<{ $role: "user" | "agent" }>`
  display: flex;
  gap: 8px;
  ${(p) => p.$role === "user" ? "align-self: flex-end; max-width: 86%;" : "align-self: stretch;"}
`;

const Avatar = styled.div<{ $role: "user" | "agent" }>`
  width: 22px;
  height: 22px;
  border-radius: 50%;
  flex-shrink: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  background: ${(p) => (p.$role === "user" ? "#3a3a3a" : "rgba(51, 99, 173, 0.18)")};
  color: ${(p) => (p.$role === "user" ? "#d4d4d4" : "#93c5fd")};
  margin-top: 2px;
  font-size: 10px;
  font-weight: 600;
  letter-spacing: 0.3px;
`;

const UserBox = styled.div`
  background: #2d2d2d;
  border: 1px solid #353535;
  padding: 8px 11px;
  border-radius: 8px;
  font-size: 13px;
  color: #e4e4e4;
  white-space: pre-wrap;
  line-height: 1.4;
`;

const AgentBox = styled.div`
  flex: 1;
  min-width: 0;
  display: flex;
  flex-direction: column;
  gap: 8px;
`;

const AgentProse = styled.div`
  font-size: 13px;
  color: #d4d4d4;
  line-height: 1.5;
  white-space: pre-wrap;
  word-wrap: break-word;
`;

const FetchedChips = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
`;

const FetchChip = styled.div`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 2px 7px;
  border-radius: 10px;
  background: #1d2c3a;
  color: #93c5fd;
  font-size: 11px;
  border: 1px solid #2a3f5a;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const BatchCard = styled.div`
  background: #1f1f20;
  border: 1px solid #333;
  border-radius: 8px;
  padding: 10px 11px;
  display: flex;
  flex-direction: column;
  gap: 8px;
`;

const CardHeader = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
  color: #b3b3b3;
  cursor: pointer;
  user-select: none;
`;

const CardActions = styled.div`
  display: flex;
  gap: 6px;
  flex-wrap: wrap;
`;

const StatusBadge = styled.span<{ $status: string }>`
  font-size: 10px;
  padding: 2px 6px;
  border-radius: 3px;
  text-transform: uppercase;
  letter-spacing: 0.3px;
  background: ${(p) =>
    p.$status === "accepted" ? "#1f3a2a"
    : p.$status === "rejected" ? "#3a1f1f"
    : p.$status === "streaming" ? "#1f2c3a"
    : "#3a311f"};
  color: ${(p) =>
    p.$status === "accepted" ? "#86efac"
    : p.$status === "rejected" ? "#fca5a5"
    : p.$status === "streaming" ? "#93c5fd"
    : "#fde68a"};
`;

const AcceptBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 10px;
  font-size: 11px;
  border-radius: 5px;
  border: 1px solid #16a34a;
  background: #16a34a;
  color: #fff;
  font-weight: 500;
  cursor: pointer;
  &:hover { background: #15803d; border-color: #15803d; }
`;

const RejectBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 10px;
  font-size: 11px;
  border-radius: 5px;
  border: 1px solid #444;
  background: transparent;
  color: #b3b3b3;
  cursor: pointer;
  &:hover { background: #2a2a2a; color: #fca5a5; border-color: #7f1d1d; }
`;

const GhostBtn = styled.button`
  padding: 3px 8px;
  font-size: 11px;
  border-radius: 4px;
  border: 1px solid #333;
  background: transparent;
  color: #9b9b9b;
  cursor: pointer;
  &:hover { background: #2a2a2a; color: #d4d4d4; }
`;

const DiffList = styled.div`
  display: flex;
  flex-direction: column;
  gap: 2px;
  max-height: 220px;
  overflow-y: auto;
  border-top: 1px solid #2a2a2a;
  padding-top: 8px;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;

  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
`;

const DiffRow = styled.div`
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 11px;
  line-height: 1.4;
`;

const CellAddr = styled.span`
  color: #93c5fd;
  width: 50px;
  flex-shrink: 0;
`;

const OldVal = styled.span`
  color: #fca5a5;
  text-decoration: line-through;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 90px;
`;

const NewVal = styled.span`
  color: #86efac;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  flex: 1;
  min-width: 0;
`;

const Composer = styled.div`
  border-top: 1px solid #2a2a2a;
  padding: 10px 12px;
  background: #1e1e1e;
  display: flex;
  flex-direction: column;
  gap: 8px;
`;

/**
 * Compact pill that lives ABOVE the composer (not inside the input row —
 * keeps the input as wide as possible). Toggles auto-apply mode.
 */
const AutoToggle = styled.button<{ $active: boolean }>`
  align-self: flex-start;
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 3px 8px;
  border-radius: 10px;
  font-size: 10px;
  font-weight: 500;
  letter-spacing: 0.2px;
  cursor: pointer;
  background: ${(p) => (p.$active ? "rgba(34, 197, 94, 0.15)" : "rgba(255,255,255,0.04)")};
  border: 1px solid ${(p) => (p.$active ? "rgba(34, 197, 94, 0.4)" : "#333")};
  color: ${(p) => (p.$active ? "#86efac" : "#9b9b9b")};
  text-transform: uppercase;
  &:hover {
    background: ${(p) => (p.$active ? "rgba(34, 197, 94, 0.2)" : "rgba(255,255,255,0.08)")};
    color: ${(p) => (p.$active ? "#86efac" : "#d4d4d4")};
  }
`;

const ComposerTopRow = styled.div`
  display: flex;
  gap: 8px;
  align-items: center;
  flex-wrap: wrap;
`;

const ModeMenuWrap = styled.div`
  position: relative;
`;

const ModeMenu = styled.div`
  position: absolute;
  bottom: calc(100% + 6px);
  left: 0;
  z-index: 30;
  background: #1a1a1a;
  border: 1px solid #333;
  border-radius: 6px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.45);
  min-width: 280px;
  overflow: hidden;
  padding: 4px;
`;

const ModeItem = styled.button<{ $active: boolean }>`
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: 2px;
  padding: 8px 10px;
  border-radius: 4px;
  width: 100%;
  text-align: left;
  cursor: pointer;
  background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.2)" : "transparent")};
  border: 0;
  color: #d4d4d4;
  &:hover { background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.25)" : "rgba(255,255,255,0.05)")}; }
`;

const ModeItemTitle = styled.div`
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  font-weight: 500;
`;

const ModeItemDesc = styled.div`
  font-size: 11px;
  color: #8c8c8c;
  line-height: 1.4;
`;

const SelectionChip = styled.div`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  align-self: flex-start;
  padding: 3px 4px 3px 8px;
  border-radius: 10px;
  background: rgba(51, 99, 173, 0.15);
  color: #93c5fd;
  font-size: 11px;
  border: 1px solid rgba(51, 99, 173, 0.35);
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const SelectionChipDismiss = styled.button`
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 14px;
  height: 14px;
  margin-left: 2px;
  border-radius: 50%;
  background: transparent;
  border: 0;
  color: #93c5fd;
  cursor: pointer;
  opacity: 0.7;
  &:hover { background: rgba(51, 99, 173, 0.3); opacity: 1; }
`;

const ComposerRow = styled.div`
  display: flex;
  gap: 8px;
  align-items: stretch;
`;

const Input = styled.textarea`
  flex: 1;
  background: #1a1a1a;
  border: 1px solid #333;
  border-radius: 6px;
  padding: 8px 12px;
  font-size: 13px;
  color: #d4d4d4;
  outline: none;
  font-family: inherit;
  resize: none;
  min-height: 36px;
  max-height: 160px;
  line-height: 1.4;
  &:focus { border-color: #3363AD; box-shadow: 0 0 0 1px rgba(51, 99, 173, 0.4); }
  &:disabled { opacity: 0.55; cursor: not-allowed; }
  &::placeholder { color: #6f6f6f; }
`;

const SendButton = styled.button`
  background: #3363AD;
  color: #fff;
  border: none;
  border-radius: 6px;
  width: 36px;
  height: 36px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  align-self: flex-end;
  flex-shrink: 0;
  &:hover:not(:disabled) { background: #4275c4; }
  &:disabled { opacity: 0.45; cursor: not-allowed; }
`;

const StopBtn = styled.button`
  background: #3a1f1f;
  color: #fca5a5;
  border: 1px solid #7f1d1d;
  border-radius: 6px;
  width: 36px;
  height: 36px;
  align-self: flex-end;
  flex-shrink: 0;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
  &:hover { background: #4d2424; color: #fecaca; }
`;

const InflightProse = styled.div`
  font-size: 13px;
  color: #b3b3b3;
  line-height: 1.5;
  white-space: pre-wrap;
  word-wrap: break-word;
`;

const Spin = styled(Loader2)`
  animation: spin 0.9s linear infinite;
  @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
`;

interface Props {
  tab: WorkbookTab | null;
  prompt: string;
  onPromptChange: (v: string) => void;
  onSubmit: () => void;
  onStop: () => void;
  onAccept: (batchId: string) => void;
  onReject: (batchId: string) => void;
  onUndo: (batchId: string) => void;
  onRedo: (batchId: string) => void;
  /** Currently-selected range on the active tab (e.g. "Sheet1!A1:C5 (15 cells)"). */
  selectionLabel: string | null;
  /**
   * Click handler on the chip's × button — dismisses the selection so it
   * won't be shipped to the agent as a focus block. The dismissal is
   * scoped to the current selection: changing the selection in the grid
   * re-arms the chip automatically (parent owns that logic).
   */
  onDismissSelection: () => void;
  autoApply: boolean;
  onSetAutoApply: (next: boolean) => void;
}

export const ChatPanel: React.FC<Props> = ({
  tab, prompt, onPromptChange, onSubmit, onStop, onAccept, onReject, onUndo, onRedo, selectionLabel, onDismissSelection, autoApply, onSetAutoApply,
}) => {
  const [modeMenuOpen, setModeMenuOpen] = useState(false);
  const modeMenuWrapRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!modeMenuOpen) return;
    const onClick = (e: MouseEvent) => {
      if (modeMenuWrapRef.current && !modeMenuWrapRef.current.contains(e.target as Node)) {
        setModeMenuOpen(false);
      }
    };
    const onEsc = (e: KeyboardEvent) => {
      if (e.key === "Escape") setModeMenuOpen(false);
    };
    window.addEventListener("mousedown", onClick);
    window.addEventListener("keydown", onEsc);
    return () => {
      window.removeEventListener("mousedown", onClick);
      window.removeEventListener("keydown", onEsc);
    };
  }, [modeMenuOpen]);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const threadRef = useRef<HTMLDivElement>(null);
  const [expandedBatches, setExpandedBatches] = useState<Record<string, boolean>>({});

  // ⌘K focuses the composer
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toLowerCase().includes("mac");
      const mod = isMac ? e.metaKey : e.ctrlKey;
      if (mod && e.key.toLowerCase() === "k") {
        e.preventDefault();
        inputRef.current?.focus();
        inputRef.current?.select();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Autoscroll to bottom on new messages / streaming updates.
  const lastBatchKey = `${tab?.batches.length ?? 0}:${tab?.streamingText?.length ?? 0}:${tab?.batches.find((b) => b.status === "streaming")?.mutations.length ?? 0}`;
  useEffect(() => {
    if (threadRef.current) {
      threadRef.current.scrollTop = threadRef.current.scrollHeight;
    }
  }, [lastBatchKey, tab?.id]);

  // Re-fit the textarea height whenever `prompt` changes from anywhere
  // (parent clearing after send, programmatic set, paste, etc.) — onInput
  // alone only catches direct user typing.
  useEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = `${el.scrollHeight}px`;
  }, [prompt]);

  const batches = tab?.batches ?? [];
  const running = !!tab?.agentRunning;

  return (
    <Panel>
      <Header>
        Conversation
        {tab && <span style={{ marginLeft: "auto", color: "#7c7c7c", textTransform: "none", letterSpacing: 0 }}>{tab.filename}</span>}
      </Header>

      <Thread ref={threadRef}>
        {!tab && (
          <Empty>
            Open a workbook with the <strong>+</strong> tab button.<br />
            Then ask the agent to edit it.
          </Empty>
        )}
        {tab && batches.length === 0 && !running && (
          <Empty>
            Ready when you are.<br />
            Try: <em>"Build a simple income model for 2022–2025"</em>
          </Empty>
        )}
        {batches.map((b) => (
          <ChatTurn
            key={b.id}
            batch={b}
            // Pass the live streaming prose only to the currently-streaming
            // batch. After `done` fires, the text is persisted into
            // batch.agent_text and we don't need the live ref anymore.
            liveText={b.status === "streaming" ? (tab?.streamingText ?? "") : ""}
            expanded={!!expandedBatches[b.id] || b.status === "streaming"}
            onToggleExpand={() => setExpandedBatches((s) => ({ ...s, [b.id]: !s[b.id] }))}
            onAccept={onAccept}
            onReject={onReject}
            onUndo={onUndo}
            onRedo={onRedo}
          />
        ))}
        {/* If the agent is "running" but no batch is in flight yet, show a thinking bubble. */}
        {running && !batches.find((b) => b.status === "streaming") && (
          <Bubble $role="agent">
            <Avatar $role="agent">A</Avatar>
            <AgentBox>
              <InflightProse>
                <Spin size={11} /> {tab?.statusMessage || "Thinking…"}
              </InflightProse>
            </AgentBox>
          </Bubble>
        )}
      </Thread>

      <Composer>
        <ComposerTopRow>
          <ModeMenuWrap ref={modeMenuWrapRef}>
            <AutoToggle
              $active={autoApply}
              onClick={() => setModeMenuOpen((v) => !v)}
              title="Click to switch between auto and manual modes"
            >
              <Zap size={9} fill={autoApply ? "currentColor" : "none"} />
              {autoApply ? "Auto" : "Manual"}
              <ChevronDown size={9} style={{ opacity: 0.7 }} />
            </AutoToggle>
            {modeMenuOpen && (
              <ModeMenu>
                <ModeItem
                  $active={!autoApply}
                  onClick={() => { onSetAutoApply(false); setModeMenuOpen(false); }}
                >
                  <ModeItemTitle>
                    {!autoApply && <Check size={11} color="#93c5fd" />}
                    Manual
                  </ModeItemTitle>
                  <ModeItemDesc>
                    You review every batch the agent produces before it's applied to the sheet. Use Accept / Reject in the chat. Safer for unfamiliar prompts.
                  </ModeItemDesc>
                </ModeItem>
                <ModeItem
                  $active={autoApply}
                  onClick={() => { onSetAutoApply(true); setModeMenuOpen(false); }}
                >
                  <ModeItemTitle>
                    {autoApply && <Check size={11} color="#93c5fd" />}
                    Auto
                  </ModeItemTitle>
                  <ModeItemDesc>
                    Edits apply immediately as the agent finishes a turn. ⌘Z still undoes anything. Best when you trust the prompt and want flow.
                  </ModeItemDesc>
                </ModeItem>
              </ModeMenu>
            )}
          </ModeMenuWrap>
          {selectionLabel && (
            <SelectionChip title="The agent will treat this range as the primary edit target. Click × to skip sending it.">
              <MousePointerSquareDashed size={11} />
              Selected: {selectionLabel}
              <SelectionChipDismiss
                onClick={onDismissSelection}
                title="Don't pass this selection to the agent"
                aria-label="Dismiss selection"
              >
                <X size={10} />
              </SelectionChipDismiss>
            </SelectionChip>
          )}
        </ComposerTopRow>
        <ComposerRow>
          <Input
            ref={inputRef}
            placeholder={
              !tab
                ? "Open a workbook first…"
                : running
                ? "Agent is working…"
                : "Ask the agent to edit this spreadsheet…  (⌘K)"
            }
            value={prompt}
            onChange={(e) => onPromptChange(e.target.value)}
            onInput={(e) => {
              // Auto-expand: reset to min then grow to content height,
              // clamped via CSS max-height. Resetting first lets the
              // textarea SHRINK when the user deletes lines, not just grow.
              const el = e.currentTarget;
              el.style.height = "auto";
              el.style.height = `${el.scrollHeight}px`;
            }}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey && !running && prompt.trim()) {
                e.preventDefault();
                onSubmit();
              }
            }}
            disabled={!tab}
            rows={1}
          />
          {running ? (
            <StopBtn onClick={onStop} title="Stop (⌘.)">
              <Square size={12} fill="currentColor" />
            </StopBtn>
          ) : (
            <SendButton onClick={onSubmit} disabled={!tab || !prompt.trim()} title="Send (Enter)">
              <Send size={14} />
            </SendButton>
          )}
        </ComposerRow>
      </Composer>
    </Panel>
  );
};

const ChatTurn: React.FC<{
  batch: ChangeBatch;
  /**
   * Live streaming prose for the currently-active batch — populated while
   * the agent is mid-turn (before `done` lands its text into batch.agent_text).
   * Empty for completed batches.
   */
  liveText: string;
  expanded: boolean;
  onToggleExpand: () => void;
  onAccept: (id: string) => void;
  onReject: (id: string) => void;
  onUndo: (id: string) => void;
  onRedo: (id: string) => void;
}> = ({ batch, liveText, expanded, onToggleExpand, onAccept, onReject, onUndo, onRedo }) => {
  const isStreaming = batch.status === "streaming";
  const cellCount = batch.mutations.length;
  const sheets = Array.from(
    new Set(batch.mutations.map((m) => mutationSheet(m)).filter(Boolean)),
  );
  return (
    <>
      {/* User bubble — the prompt that started this turn. */}
      <Bubble $role="user">
        <UserBox>{batch.prompt || "(empty prompt)"}</UserBox>
        <Avatar $role="user"><User size={12} /></Avatar>
      </Bubble>
      {/* Agent bubble — prose + batch card. */}
      <Bubble $role="agent">
        <Avatar $role="agent">A</Avatar>
        <AgentBox>
          {/* Thinking block — the agent's prose as it works. During
              streaming this is the live buffer; once `done` lands, it's
              the persisted agent_text. Renders above the batch card so
              the narrative flows: "here's what I'm doing → here are the
              changes → here's why" (justification block below). */}
          {(batch.agent_text || (isStreaming && liveText)) && (
            <AgentProse>{batch.agent_text || liveText}</AgentProse>
          )}
          {batch.fetched_urls && batch.fetched_urls.length > 0 && (
            <FetchedChips>
              {batch.fetched_urls.map((url, i) => (
                <FetchChip key={i} title={url}>
                  <Globe size={10} />
                  {shortenUrl(url)}
                </FetchChip>
              ))}
            </FetchedChips>
          )}
          {(cellCount > 0 || isStreaming) && (
            <BatchCard>
              <CardHeader onClick={onToggleExpand}>
                {expanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                <span style={{ flex: 1, color: "#e4e4e4", fontWeight: 500 }}>
                  {isStreaming ? "Writing…" : "Proposed changes"}
                </span>
                <span style={{ color: "#7c7c7c" }}>
                  {cellCount} edit{cellCount === 1 ? "" : "s"}
                  {sheets.length > 0 && ` · ${sheets.join(", ")}`}
                </span>
                <StatusBadge $status={batch.status}>{batch.status}</StatusBadge>
              </CardHeader>
              {expanded && cellCount > 0 && (
                <DiffList>
                  {batch.mutations.slice(0, 80).map((m, i) => (
                    <DiffRow key={i}>
                      <CellAddr>{addrOf(m)}</CellAddr>
                      <OldVal>{fmtVal(oldOf(m))}</OldVal>
                      <ArrowRight size={10} style={{ color: "#555", flexShrink: 0 }} />
                      <NewVal>{fmtVal(newOf(m))}</NewVal>
                    </DiffRow>
                  ))}
                  {batch.mutations.length > 80 && (
                    <DiffRow>
                      <span style={{ color: "#666", fontStyle: "italic" }}>
                        … {batch.mutations.length - 80} more
                      </span>
                    </DiffRow>
                  )}
                </DiffList>
              )}
              <CardActions>
                {batch.status === "pending" && (
                  <>
                    <AcceptBtn onClick={() => onAccept(batch.id)}><Check size={11} /> Accept</AcceptBtn>
                    <RejectBtn onClick={() => onReject(batch.id)}><X size={11} /> Reject</RejectBtn>
                  </>
                )}
                {batch.status === "accepted" && (
                  <GhostBtn onClick={() => onUndo(batch.id)} title="Revert this change (⌘Z)">
                    <Undo2 size={11} style={{ marginRight: 4 }} /> Undo
                  </GhostBtn>
                )}
                {batch.status === "rejected" && (
                  <GhostBtn onClick={() => onRedo(batch.id)} title="Re-apply this change (⌘⇧Z)">
                    <Redo2 size={11} style={{ marginRight: 4 }} /> Redo
                  </GhostBtn>
                )}
              </CardActions>
            </BatchCard>
          )}
          {isStreaming && cellCount === 0 && (
            <InflightProse>
              <Spin size={11} /> Waiting for the first edit…
            </InflightProse>
          )}
          {/* Justification block — the agent's final summary after `done`.
              Closes the turn below the batch card so the user sees the
              changes first, then the rationale. Same flat typography as
              the thinking block above. */}
          {batch.justification && <AgentProse>{batch.justification}</AgentProse>}
        </AgentBox>
      </Bubble>
    </>
  );
};

// --- diff display helpers (mirror ChangesPanel) ---

function colLetters(col: number): string {
  let n = col;
  let s = "";
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}

function mutationSheet(m: UniverMutation): string {
  if (m.type === "set_cell") return m.address.sheet;
  if (m.type === "create_sheet" || m.type === "delete_sheet") return m.name;
  if (m.type === "rename_sheet") return m.new_name;
  if ("sheet" in m) return (m as any).sheet;
  return "";
}

function addrOf(m: UniverMutation): string {
  if (m.type === "set_cell") return `${colLetters(m.address.col)}${m.address.row + 1}`;
  if (m.type === "set_range") return `${colLetters(m.start_col)}${m.start_row + 1}+`;
  if (m.type === "set_format" || m.type === "merge_cells" || m.type === "unmerge_cells" || m.type === "clear_range") return m.range;
  if (m.type === "set_column_width") return "cols " + m.columns.map((c) => colLetters(c)).join(",");
  if (m.type === "set_row_height") return "rows " + m.rows.map((r) => r + 1).join(",");
  if (m.type === "create_sheet") return `+ ${m.name}`;
  if (m.type === "delete_sheet") return `− ${m.name}`;
  if (m.type === "rename_sheet") return `${m.old_name} →`;
  if (m.type === "insert_rows") return `row ${m.before + 1}`;
  if (m.type === "delete_rows") return `row ${m.start + 1}`;
  if (m.type === "insert_columns") return `col ${colLetters(m.before)}`;
  if (m.type === "delete_columns") return `col ${colLetters(m.start)}`;
  if (m.type === "freeze_panes" || m.type === "unfreeze_panes") return `panes`;
  if (m.type === "hide_rows" || m.type === "show_rows") return "rows " + m.rows.map((r) => r + 1).join(",");
  if (m.type === "hide_columns" || m.type === "show_columns") return "cols " + m.columns.map((c) => colLetters(c)).join(",");
  return "?";
}

function oldOf(m: UniverMutation): any {
  if (m.type === "set_cell") return m.old_formula ?? m.old_value;
  if (m.type === "set_format") return "fmt";
  if (m.type === "set_column_width") return "width";
  if (m.type === "set_row_height") return "height";
  if (m.type === "merge_cells") return "split";
  if (m.type === "unmerge_cells") return "merged";
  if (m.type === "clear_range") return "value";
  if (m.type === "create_sheet") return "—";
  if (m.type === "delete_sheet") return "sheet";
  if (m.type === "rename_sheet") return m.old_name;
  if (m.type === "insert_rows" || m.type === "insert_columns") return "—";
  if (m.type === "delete_rows" || m.type === "delete_columns") return "exists";
  if (m.type === "freeze_panes" || m.type === "unfreeze_panes") return "frozen?";
  if (m.type === "hide_rows" || m.type === "hide_columns") return "shown";
  if (m.type === "show_rows" || m.type === "show_columns") return "hidden";
  return null;
}

function newOf(m: UniverMutation): any {
  if (m.type === "set_cell") return m.new_formula ?? m.new_value;
  if (m.type === "set_format") return formatSummary(m.new_format);
  if (m.type === "set_column_width") return `${m.new_width}px`;
  if (m.type === "set_row_height") return `${m.new_height}px`;
  if (m.type === "merge_cells") return "merged";
  if (m.type === "unmerge_cells") return "split";
  if (m.type === "clear_range") return "(cleared)";
  if (m.type === "create_sheet") return `new sheet${m.tab_color ? ` · ${m.tab_color}` : ""}`;
  if (m.type === "delete_sheet") return "(deleted)";
  if (m.type === "rename_sheet") return m.new_name;
  if (m.type === "insert_rows") return `+${m.count} rows`;
  if (m.type === "delete_rows") return `−${m.count} rows`;
  if (m.type === "insert_columns") return `+${m.count} cols`;
  if (m.type === "delete_columns") return `−${m.count} cols`;
  if (m.type === "freeze_panes") return `freeze ${m.freeze_rows}r × ${m.freeze_cols}c`;
  if (m.type === "unfreeze_panes") return "unfrozen";
  if (m.type === "hide_rows" || m.type === "hide_columns") return "hidden";
  if (m.type === "show_rows" || m.type === "show_columns") return "shown";
  return "(range)";
}

function formatSummary(f: any): string {
  const parts: string[] = [];
  if (f?.bold) parts.push("bold");
  if (f?.italic) parts.push("italic");
  if (f?.underline) parts.push("underline");
  if (f?.strike) parts.push("strike");
  if (f?.font_family) parts.push(f.font_family);
  if (f?.font_color) parts.push(`color ${f.font_color}`);
  if (f?.font_size) parts.push(`${f.font_size}pt`);
  if (f?.horizontal_align) parts.push(`align ${f.horizontal_align}`);
  if (f?.number_format) parts.push(`fmt ${f.number_format}`);
  return parts.length ? parts.join(" + ") : "(empty format)";
}

function shortenUrl(url: string): string {
  try {
    const u = new URL(url);
    const path = u.pathname.length > 24 ? u.pathname.slice(0, 23) + "…" : u.pathname;
    return `${u.hostname}${path === "/" ? "" : path}`;
  } catch {
    return url.length > 40 ? url.slice(0, 39) + "…" : url;
  }
}

function fmtVal(v: any): string {
  if (v === null || v === undefined) return "(empty)";
  if (typeof v === "string") return v.length > 24 ? v.slice(0, 23) + "…" : v;
  return String(v);
}
