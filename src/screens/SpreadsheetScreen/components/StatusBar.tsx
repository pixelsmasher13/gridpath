import React, { useEffect, useRef, useState } from "react";
import styled from "styled-components";
import { Check, ChevronUp } from "lucide-react";
import type { StatusPhase } from "../state/tabs";
import {
  getSettingValue,
  getModel,
  setModel,
  getEffort,
  setEffort,
  SETTING_KEYS,
  MODEL_PRESETS,
  EFFORT_PRESETS,
  humanizeModel,
  type Provider,
} from "../settingsApi";

const Bar = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 4px 12px;
  border-top: 1px solid #2a2a2a;
  font-size: 11px;
  color: #9b9b9b;
  background: #1e1e1e;
  min-height: 24px;
`;

const Dot = styled.span<{ $phase: StatusPhase }>`
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: ${(p) => {
    switch (p.$phase) {
      case "thinking": return "#3363AD";
      case "writing":  return "#22c55e";
      case "error":    return "#ef4444";
      case "done":     return "#7c7c7c";
      default:         return "#3a3a3a";
    }
  }};
`;

export type SelectionStats = {
  /** Non-empty cells in the selection. */
  count: number;
  /** Cells that parsed as numbers. */
  numericCount: number;
  sum: number;
  avg: number;
  min: number;
  max: number;
};

interface Props {
  workbookPath: string | null;
  dirty: boolean;
  phase: StatusPhase;
  message: string;
  lastSavedAt: number | null;
  inputTokens?: number;
  outputTokens?: number;
  /** Excel-style live stats for the current multi-cell selection. */
  selectionStats?: SelectionStats | null;
  /** True while an agent turn is in flight — model switches are disabled. */
  agentRunning?: boolean;
  /** Opens Settings so the user can type a custom model id. */
  onOpenSettings?: () => void;
}

export const StatusBar: React.FC<Props> = ({
  workbookPath, dirty, phase, message, lastSavedAt, inputTokens, outputTokens,
  selectionStats, agentRunning, onOpenSettings,
}) => {
  // Re-render every 5 seconds so the "saved Ns ago" stays roughly fresh.
  const [, setTick] = useState(0);
  useEffect(() => {
    if (!lastSavedAt) return;
    const id = setInterval(() => setTick((n) => n + 1), 5000);
    return () => clearInterval(id);
  }, [lastSavedAt]);

  const [apiChoice, setApiChoice] = useState("");
  const [modelId, setModelId] = useState("");
  const [effort, setEffortLocal] = useState("");
  const [menuOpen, setMenuOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const menuWrapRef = useRef<HTMLDivElement>(null);

  const refresh = React.useCallback(async () => {
    const choice = await getSettingValue(SETTING_KEYS.apiChoice);
    let provider: Provider | null = null;
    if (choice === "openai-codex") provider = "openai-codex";
    else if (choice === "claude" || choice === "claude-subscription") provider = "claude";
    setApiChoice(choice);
    if (!provider) {
      setModelId("");
      setEffortLocal("");
      return;
    }
    const [raw, eff] = await Promise.all([getModel(provider), getEffort(provider)]);
    setModelId(raw);
    setEffortLocal(eff);
  }, []);

  useEffect(() => {
    let cancelled = false;
    const run = async () => {
      if (cancelled) return;
      await refresh();
    };
    run();
    const id = setInterval(run, 1500);
    const onFocus = () => { run(); };
    window.addEventListener("focus", onFocus);
    return () => {
      cancelled = true;
      clearInterval(id);
      window.removeEventListener("focus", onFocus);
    };
  }, [refresh]);

  // Close the menu on outside click / Escape.
  useEffect(() => {
    if (!menuOpen) return;
    const onDown = (e: MouseEvent) => {
      if (!menuWrapRef.current?.contains(e.target as Node)) setMenuOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setMenuOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [menuOpen]);

  const provider: Provider | null =
    apiChoice === "openai-codex" ? "openai-codex"
    : apiChoice === "claude" || apiChoice === "claude-subscription" ? "claude"
    : null;

  const modelLabel = humanizeModel(modelId, apiChoice);
  const presets = provider ? MODEL_PRESETS[provider] : [];
  const knownIds = new Set(presets.map((p) => p.id));
  const isCustom = !!modelId && !knownIds.has(modelId);
  const effortPresets = provider ? EFFORT_PRESETS[provider] : [];
  // Label for the button: preset label ("Thorough"), or the raw value for
  // hand-set ones ("medium"), or nothing while unknown.
  const effortLabel = effortPresets.find((p) => p.id === effort)?.label ?? effort;

  const pickModel = async (id: string) => {
    if (!provider || agentRunning || busy) return;
    setBusy(true);
    try {
      await setModel(provider, id);
      setModelId(id);
      setMenuOpen(false);
    } catch (e) {
      console.warn("[statusbar] set model failed:", e);
    } finally {
      setBusy(false);
    }
  };

  // Effort clicks keep the menu open — it's a toggle, and users often set
  // it on the way to picking a model or closing by clicking away.
  const pickEffort = async (id: string) => {
    if (!provider || agentRunning || busy) return;
    setBusy(true);
    try {
      await setEffort(provider, id);
      setEffortLocal(id);
    } catch (e) {
      console.warn("[statusbar] set effort failed:", e);
    } finally {
      setBusy(false);
    }
  };

  return (
    <Bar>
      <Dot $phase={phase} />
      <span>{message || (phase === "idle" ? "Ready" : phase)}</span>
      <RightCluster>
        {selectionStats && (
          <span style={{ color: "#b3b3b3" }} title="Selection statistics">
            {selectionStats.numericCount > 0 && (
              <>
                Average: {fmtStat(selectionStats.avg)}
                {"\u2003"}Sum: {fmtStat(selectionStats.sum)}
                {"\u2003"}
              </>
            )}
            Count: {selectionStats.count}
          </span>
        )}
        <span style={{ color: "#7c7c7c" }}>
          {workbookPath ? (
            <>
              {workbookPath.split("/").pop()}
              {dirty ? " •" : ""}
            </>
          ) : (
            "No workbook"
          )}
        </span>
        {lastSavedAt && (
          <span style={{ color: "#666" }}>
            · saved {formatSince(lastSavedAt)}
          </span>
        )}
        {(inputTokens != null || outputTokens != null) && (
          <span style={{ color: "#666" }}>
            · in {inputTokens ?? 0} / out {outputTokens ?? 0}
          </span>
        )}
        {provider && modelLabel && (
          <ModelMenuWrap ref={menuWrapRef}>
            <ModelButton
              type="button"
              $open={menuOpen}
              disabled={!!agentRunning || busy}
              title={
                agentRunning
                  ? "Model switches apply on the next turn — wait for the agent to finish"
                  : "Change model"
              }
              onClick={() => setMenuOpen((v) => !v)}
            >
              {modelLabel}
              {effortLabel ? ` · ${effortLabel}` : ""}
              <ChevronUp size={10} style={{ opacity: 0.7 }} />
            </ModelButton>
            {menuOpen && (
              <ModelMenu role="menu">
                <MenuHeader>
                  {provider === "openai-codex" ? "ChatGPT models" : "Claude models"}
                </MenuHeader>
                {presets.map((p) => {
                  const selected = modelId === p.id;
                  return (
                    <MenuItem
                      key={p.id}
                      role="menuitemradio"
                      aria-checked={selected}
                      $selected={selected}
                      onClick={() => pickModel(p.id)}
                    >
                      <span style={{ flex: 1, minWidth: 0 }}>
                        <MenuItemTitle>{p.label}</MenuItemTitle>
                        {p.recommended && <MenuItemHint>Recommended</MenuItemHint>}
                      </span>
                      {selected && <Check size={12} color="#86efac" />}
                    </MenuItem>
                  );
                })}
                {isCustom && (
                  <MenuItem
                    role="menuitemradio"
                    aria-checked
                    $selected
                    onClick={() => setMenuOpen(false)}
                  >
                    <span style={{ flex: 1, minWidth: 0 }}>
                      <MenuItemTitle>{humanizeModel(modelId)}</MenuItemTitle>
                      <MenuItemHint>Custom</MenuItemHint>
                    </span>
                    <Check size={12} color="#86efac" />
                  </MenuItem>
                )}
                {effortPresets.length > 0 && (
                  <>
                    <MenuDivider />
                    <MenuHeader>Effort</MenuHeader>
                    <EffortRow>
                      {effortPresets.map((p) => (
                        <EffortSegment
                          key={p.id}
                          type="button"
                          $active={effort === p.id}
                          title={p.hint}
                          onClick={() => pickEffort(p.id)}
                        >
                          {p.label}
                        </EffortSegment>
                      ))}
                    </EffortRow>
                  </>
                )}
                {onOpenSettings && (
                  <>
                    <MenuDivider />
                    <MenuItem
                      role="menuitem"
                      $selected={false}
                      onClick={() => {
                        setMenuOpen(false);
                        onOpenSettings();
                      }}
                    >
                      <span style={{ flex: 1 }}>
                        <MenuItemTitle>Custom model ID…</MenuItemTitle>
                        <MenuItemHint>Open Settings</MenuItemHint>
                      </span>
                    </MenuItem>
                  </>
                )}
              </ModelMenu>
            )}
          </ModelMenuWrap>
        )}
      </RightCluster>
    </Bar>
  );
};

const RightCluster = styled.div`
  margin-left: auto;
  display: flex;
  align-items: center;
  gap: 8px;
`;

const ModelMenuWrap = styled.div`
  position: relative;
  display: inline-flex;
`;

const ModelButton = styled.button<{ $open: boolean }>`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  margin: 0;
  padding: 1px 6px;
  border-radius: 4px;
  border: 1px solid ${(p) => (p.$open ? "#3a4a5c" : "transparent")};
  background: ${(p) => (p.$open ? "#252a30" : "transparent")};
  color: #9b9b9b;
  font-size: 11px;
  line-height: 1.2;
  cursor: pointer;
  &:hover:not(:disabled) {
    color: #d4d4d4;
    background: #252a30;
    border-color: #333;
  }
  &:disabled {
    opacity: 0.55;
    cursor: not-allowed;
  }
`;

const ModelMenu = styled.div`
  position: absolute;
  right: 0;
  bottom: calc(100% + 6px);
  min-width: 200px;
  max-width: 260px;
  padding: 6px;
  background: #252526;
  border: 1px solid #333;
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.45);
  z-index: 40;
`;

const MenuHeader = styled.div`
  padding: 4px 8px 6px;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  color: #7c7c7c;
`;

const MenuItem = styled.button<{ $selected: boolean }>`
  display: flex;
  align-items: center;
  gap: 8px;
  width: 100%;
  text-align: left;
  margin: 0;
  padding: 7px 8px;
  border: 0;
  border-radius: 5px;
  background: ${(p) => (p.$selected ? "rgba(51, 99, 173, 0.18)" : "transparent")};
  color: #d4d4d4;
  cursor: pointer;
  &:hover {
    background: ${(p) => (p.$selected ? "rgba(51, 99, 173, 0.28)" : "#2a2a2a")};
  }
`;

const MenuItemTitle = styled.div`
  font-size: 12px;
  font-weight: 500;
  color: #e8e8e8;
`;

const MenuItemHint = styled.div`
  font-size: 10px;
  color: #7c7c7c;
  margin-top: 1px;
`;

const MenuDivider = styled.div`
  height: 1px;
  background: #333;
  margin: 4px 2px;
`;

const EffortRow = styled.div`
  display: flex;
  padding: 0 2px 4px;
`;

const EffortSegment = styled.button<{ $active: boolean }>`
  flex: 1;
  padding: 5px 4px;
  border: 1px solid ${(p) => (p.$active ? "rgba(51, 99, 173, 0.7)" : "#333")};
  background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.18)" : "transparent")};
  color: ${(p) => (p.$active ? "#cfe1f8" : "#9a9a9a")};
  font-size: 11px;
  font-weight: 500;
  cursor: pointer;
  &:first-child { border-radius: 5px 0 0 5px; }
  &:last-child { border-radius: 0 5px 5px 0; }
  &:not(:first-child) { margin-left: -1px; }
  &:hover { color: #cfe1f8; }
`;

/** Compact number formatting for the status-bar stats (thousands + ≤2 dp). */
function fmtStat(n: number): string {
  if (!Number.isFinite(n)) return "—";
  const rounded = Math.round(n * 100) / 100;
  return rounded.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function formatSince(ts: number): string {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 5) return "just now";
  if (s < 60) return `${s}s ago`;
  const m = Math.round(s / 60);
  if (m < 60) return `${m}m ago`;
  const h = Math.round(m / 60);
  return `${h}h ago`;
}
