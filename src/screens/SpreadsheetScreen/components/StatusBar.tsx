import React, { useEffect, useState } from "react";
import styled from "styled-components";
import type { StatusPhase } from "../state/tabs";
import { getSettingValue, getModel, SETTING_KEYS, type Provider } from "../settingsApi";

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

interface Props {
  workbookPath: string | null;
  dirty: boolean;
  phase: StatusPhase;
  message: string;
  lastSavedAt: number | null;
  inputTokens?: number;
  outputTokens?: number;
}

export const StatusBar: React.FC<Props> = ({
  workbookPath, dirty, phase, message, lastSavedAt, inputTokens, outputTokens,
}) => {
  // Re-render every 5 seconds so the "saved Ns ago" stays roughly fresh.
  const [, setTick] = useState(0);
  useEffect(() => {
    if (!lastSavedAt) return;
    const id = setInterval(() => setTick((n) => n + 1), 5000);
    return () => clearInterval(id);
  }, [lastSavedAt]);

  // Current model indicator. Re-polls every 1.5s + refreshes whenever the
  // window regains focus, so a provider switch in SettingsModal shows up
  // promptly without wiring an explicit setting-changed event.
  const [modelLabel, setModelLabel] = useState<string>("");
  useEffect(() => {
    let cancelled = false;
    const refresh = async () => {
      const apiChoice = await getSettingValue(SETTING_KEYS.apiChoice);
      let provider: Provider | null = null;
      if (apiChoice === "openai-codex") provider = "openai-codex";
      else if (apiChoice === "claude" || apiChoice === "claude-subscription") provider = "claude";
      if (!provider) { if (!cancelled) setModelLabel(""); return; }
      const raw = await getModel(provider);
      if (!cancelled) setModelLabel(humanizeModel(raw, apiChoice));
    };
    refresh();
    const id = setInterval(refresh, 1500);
    const onFocus = () => { refresh(); };
    window.addEventListener("focus", onFocus);
    return () => {
      cancelled = true;
      clearInterval(id);
      window.removeEventListener("focus", onFocus);
    };
  }, []);

  return (
    <Bar>
      <Dot $phase={phase} />
      <span>{message || (phase === "idle" ? "Ready" : phase)}</span>
      <span style={{ marginLeft: "auto", color: "#7c7c7c" }}>
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
      {modelLabel && (
        <span style={{ color: "#666" }} title="Current LLM model (change in Settings)">
          · {modelLabel}
        </span>
      )}
    </Bar>
  );
};

/**
 * Compress a model string into a status-bar-friendly label.
 *   claude-sonnet-4-5-20250929  → "Sonnet 4.5"
 *   claude-opus-4-7             → "Opus 4.7"
 *   claude-haiku-4-5-20251001   → "Haiku 4.5"
 *   gpt-5-codex                 → "GPT-5 Codex"
 *   gpt-5.5                     → "GPT-5.5"
 * Falls back to the raw string if no pattern matches so unknown models
 * still surface (rather than silently hiding when Anthropic ships
 * something new).
 */
function humanizeModel(raw: string, apiChoice: string): string {
  if (!raw) {
    // No model override set — show the provider name as a coarse hint.
    if (apiChoice === "openai-codex") return "ChatGPT";
    if (apiChoice === "claude-subscription") return "Claude Pro";
    if (apiChoice === "claude") return "Claude";
    return "";
  }
  const claude = raw.match(/^claude-(sonnet|opus|haiku)-(\d+(?:[-.]\d+)?)/i);
  if (claude) {
    const tier = claude[1][0].toUpperCase() + claude[1].slice(1);
    const ver = claude[2].replace(/-/g, ".");
    return `${tier} ${ver}`;
  }
  const gpt = raw.match(/^gpt-([\d.]+)(?:-(.+))?$/i);
  if (gpt) {
    const ver = gpt[1];
    const suffix = gpt[2] ? ` ${gpt[2][0].toUpperCase() + gpt[2].slice(1)}` : "";
    return `GPT-${ver}${suffix}`;
  }
  return raw;
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
