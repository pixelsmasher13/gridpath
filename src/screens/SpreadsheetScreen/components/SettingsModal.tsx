import React, { useEffect, useState } from "react";
import styled from "styled-components";
import { X, Eye, EyeOff, Check, Info, LogOut, User as UserIcon } from "lucide-react";
import { getSettingValue, setSettingValue, SETTING_KEYS, getModel, setModel, MODEL_PRESETS, getEffort, setEffort, EFFORT_PRESETS, type Provider } from "../settingsApi";
import { codexLogin, codexLogout, codexStatus, type CodexStatus } from "../codexAuth";
import { useAuth } from "../../../Providers/AuthProvider";
import type { WorkbookTab } from "../state/tabs";
import { listSessions, type SessionRow } from "../sessionDb";

const Backdrop = styled.div`
  position: fixed;
  inset: 0;
  background: rgba(0, 0, 0, 0.55);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 100;
`;

const Modal = styled.div`
  width: 540px;
  max-width: calc(100vw - 40px);
  max-height: calc(100vh - 80px);
  background: #1e1e1e;
  border: 1px solid #333;
  border-radius: 10px;
  color: #d4d4d4;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  box-shadow: 0 16px 56px rgba(0, 0, 0, 0.55);
`;

const Header = styled.div`
  display: flex;
  align-items: center;
  padding: 12px 14px;
  border-bottom: 1px solid #2a2a2a;
  font-weight: 600;
  font-size: 13px;
`;

const TabStrip = styled.div`
  display: flex;
  gap: 2px;
  padding: 6px 10px 0;
  border-bottom: 1px solid #2a2a2a;
  background: #1a1a1a;
`;

const TabBtn = styled.button<{ $active: boolean }>`
  background: transparent;
  border: 0;
  color: ${(p) => (p.$active ? "#e4e4e4" : "#8a8a8a")};
  font-size: 12px;
  font-weight: 500;
  padding: 8px 12px;
  cursor: pointer;
  border-bottom: 2px solid ${(p) => (p.$active ? "#3363AD" : "transparent")};
  margin-bottom: -1px;
  &:hover { color: #e4e4e4; }
`;

const UsageTable = styled.div`
  display: flex;
  flex-direction: column;
  background: #161616;
  border: 1px solid #2a2a2a;
  border-radius: 6px;
  overflow: hidden;
`;

const UsageRow = styled.div<{ $head?: boolean }>`
  display: grid;
  grid-template-columns: 1fr 70px 70px 80px 90px;
  gap: 8px;
  padding: ${(p) => (p.$head ? "8px 12px" : "10px 12px")};
  font-size: ${(p) => (p.$head ? "10px" : "12px")};
  font-weight: ${(p) => (p.$head ? "500" : "400")};
  text-transform: ${(p) => (p.$head ? "uppercase" : "none")};
  letter-spacing: ${(p) => (p.$head ? "0.4px" : "normal")};
  color: ${(p) => (p.$head ? "#7c7c7c" : "#d4d4d4")};
  border-bottom: ${(p) => (p.$head ? "1px solid #2a2a2a" : "1px solid #1f1f1f")};
  background: ${(p) => (p.$head ? "#1a1a1a" : "transparent")};
  &:last-child { border-bottom: 0; }
`;

const UsageCell = styled.span<{ $mono?: boolean; $dim?: boolean; $right?: boolean }>`
  font-family: ${(p) => (p.$mono ? "ui-monospace, SFMono-Regular, Menlo, monospace" : "inherit")};
  color: ${(p) => (p.$dim ? "#7c7c7c" : "inherit")};
  text-align: ${(p) => (p.$right ? "right" : "left")};
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const UsageSummary = styled.div`
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 10px;
`;

const UsageStat = styled.div`
  background: #161616;
  border: 1px solid #2a2a2a;
  border-radius: 6px;
  padding: 10px 12px;
  display: flex;
  flex-direction: column;
  gap: 2px;
`;

const UsageStatLabel = styled.span`
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  color: #7c7c7c;
`;

const UsageStatValue = styled.span`
  font-size: 18px;
  font-weight: 600;
  color: #e4e4e4;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
`;

const UsageStatHint = styled.span`
  font-size: 10px;
  color: #7c7c7c;
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
  padding: 16px;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  gap: 16px;

  /* Slim, always-visible scrollbar. Low-contrast 10% white thumb so
     it doesn't compete with content, bumps to 20% when the thumb
     itself is hovered for a clear grab affordance. */
  &::-webkit-scrollbar { width: 6px; }
  &::-webkit-scrollbar-track { background: transparent; }
  &::-webkit-scrollbar-thumb {
    background: rgba(255, 255, 255, 0.10);
    border-radius: 3px;
    transition: background 180ms ease;
  }
  &::-webkit-scrollbar-thumb:hover { background: rgba(255, 255, 255, 0.20); }
`;

const SectionLabel = styled.div`
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: #b3b3b3;
`;

const FieldGroup = styled.div`
  display: flex;
  flex-direction: column;
  gap: 6px;
`;

const FieldLabel = styled.label`
  font-size: 12px;
  color: #e4e4e4;
  font-weight: 500;
  display: flex;
  align-items: center;
  gap: 6px;
`;

const FieldRow = styled.div`
  display: flex;
  gap: 6px;
  align-items: stretch;
`;

const Input = styled.input`
  flex: 1;
  background: #161616;
  border: 1px solid #333;
  border-radius: 5px;
  padding: 8px 10px;
  color: #e4e4e4;
  font-size: 12px;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  outline: none;
  &:focus { border-color: #3363AD; }
  &::placeholder { color: #555; }
`;

const RevealBtn = styled.button`
  background: transparent;
  border: 1px solid #333;
  border-radius: 5px;
  color: #9b9b9b;
  cursor: pointer;
  padding: 0 10px;
  display: flex;
  align-items: center;
  &:hover { background: #2a2a2a; color: #e4e4e4; }
`;

const Hint = styled.div`
  font-size: 11px;
  color: #7c7c7c;
  line-height: 1.45;
  display: flex;
  align-items: flex-start;
  gap: 6px;
`;

const Active = styled.div<{ $ok: boolean }>`
  background: ${(p) => (p.$ok ? "rgba(34, 197, 94, 0.08)" : "rgba(239, 68, 68, 0.08)")};
  border: 1px solid ${(p) => (p.$ok ? "rgba(34, 197, 94, 0.3)" : "rgba(239, 68, 68, 0.3)")};
  color: ${(p) => (p.$ok ? "#86efac" : "#fca5a5")};
  padding: 8px 10px;
  border-radius: 6px;
  font-size: 12px;
  display: flex;
  align-items: center;
  gap: 6px;
`;

const AccountCard = styled.div`
  background: #1a1a1a;
  border: 1px solid #2a2a2a;
  border-radius: 6px;
  padding: 12px;
  display: flex;
  flex-direction: column;
  gap: 6px;
`;

const AccountHeader = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 4px;
`;

const Avatar = styled.div`
  width: 28px;
  height: 28px;
  border-radius: 50%;
  background: rgba(51, 99, 173, 0.18);
  color: #93c5fd;
  display: flex;
  align-items: center;
  justify-content: center;
  font-weight: 600;
  font-size: 12px;
`;

const AccountField = styled.div`
  display: grid;
  grid-template-columns: 110px 1fr;
  gap: 8px;
  font-size: 12px;
  line-height: 1.5;
`;

const AccountFieldLabel = styled.span`
  color: #7c7c7c;
`;

const AccountFieldValue = styled.span`
  color: #d4d4d4;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
`;

const LogoutBtn = styled.button`
  align-self: flex-start;
  background: transparent;
  border: 1px solid #333;
  color: #d4d4d4;
  border-radius: 5px;
  padding: 4px 10px;
  font-size: 11px;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 5px;
  margin-top: 4px;
  &:hover { background: #2a2a2a; color: #fca5a5; border-color: #7f1d1d; }
`;

const ProviderGroup = styled.div`
  display: flex;
  flex-direction: column;
  gap: 4px;
  background: #161616;
  border: 1px solid #2a2a2a;
  border-radius: 6px;
  padding: 6px;
`;

const ProviderOption = styled.label<{ $active: boolean }>`
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 10px;
  border-radius: 4px;
  cursor: pointer;
  background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.18)" : "transparent")};
  border: 1px solid ${(p) => (p.$active ? "rgba(51, 99, 173, 0.5)" : "transparent")};
  &:hover { background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.22)" : "#1f1f1f")}; }
`;

const Radio = styled.span<{ $active: boolean }>`
  width: 14px;
  height: 14px;
  border-radius: 50%;
  border: 1.5px solid ${(p) => (p.$active ? "#3363AD" : "#444")};
  display: inline-flex;
  align-items: center;
  justify-content: center;
  flex-shrink: 0;
  &:after {
    content: '';
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: ${(p) => (p.$active ? "#3363AD" : "transparent")};
  }
`;

// Small marker shown ONLY for providers that aren't configured. Configured
// ones get nothing — the radio + the "Connected" banner above already
// communicate their state. Three matching badges in a row read as
// decoration rather than information.
const NotConfiguredHint = styled.span`
  margin-left: auto;
  font-size: 10px;
  color: #9b6b6b;
  display: inline-flex;
  align-items: center;
  gap: 4px;
`;

const NotConfiguredDot = styled.span`
  width: 5px;
  height: 5px;
  border-radius: 50%;
  background: #c46060;
`;

const Footer = styled.div`
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 12px 14px;
  border-top: 1px solid #2a2a2a;
`;

const Btn = styled.button<{ $primary?: boolean }>`
  background: ${(p) => (p.$primary ? "#3363AD" : "transparent")};
  color: ${(p) => (p.$primary ? "#fff" : "#d4d4d4")};
  border: 1px solid ${(p) => (p.$primary ? "#3363AD" : "#333")};
  border-radius: 5px;
  padding: 7px 14px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:hover:not(:disabled) {
    background: ${(p) => (p.$primary ? "#4275c4" : "#2a2a2a")};
  }
  &:disabled { opacity: 0.5; cursor: not-allowed; }
`;

interface Props {
  open: boolean;
  onClose: () => void;
  /**
   * Open workspace tabs — used by the Usage tab to show per-session token
   * + prompt-cache totals. The data is in-memory only (we don't persist
   * usage stats yet), so closing a tab clears its row.
   */
  tabs: WorkbookTab[];
}

const EffortSegment = styled.button<{ $active: boolean }>`
  flex: 1;
  padding: 6px 8px;
  border: 1px solid ${(p) => (p.$active ? "rgba(51, 99, 173, 0.7)" : "#2a2a2a")};
  background: ${(p) => (p.$active ? "rgba(51, 99, 173, 0.18)" : "#161616")};
  color: ${(p) => (p.$active ? "#cfe1f8" : "#9a9a9a")};
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:first-child { border-radius: 6px 0 0 6px; }
  &:last-child { border-radius: 0 6px 6px 0; }
  &:not(:first-child) { margin-left: -1px; }
  &:hover { color: #cfe1f8; }
`;

/**
 * Three-way effort control (Fast / Standard / Thorough) mapped to the
 * provider's native reasoning-effort scale. Middle option = the provider
 * default, so untouched settings behave exactly as before.
 */
const EffortPicker: React.FC<{
  provider: Provider;
  value: string;
  onChange: (v: string) => void;
}> = ({ provider, value, onChange }) => {
  const presets = EFFORT_PRESETS[provider];
  // Unset or unrecognized (e.g. hand-set "medium" on Claude) → highlight
  // nothing rather than lying about what's active.
  const active = presets.find((p) => p.id === value);
  return (
    <FieldGroup>
      <FieldLabel>Effort</FieldLabel>
      <FieldRow>
        {presets.map((p) => (
          <EffortSegment
            key={p.id}
            $active={active?.id === p.id}
            onClick={() => onChange(p.id)}
            title={p.hint}
            type="button"
          >
            {p.label}
          </EffortSegment>
        ))}
      </FieldRow>
      <Hint>
        <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
        {active ? active.hint : "How hard the model thinks per request. Fast is snappier and cheaper; Thorough reasons longest before acting."}
        {" "}Applies from the next message.
      </Hint>
    </FieldGroup>
  );
};

const ModelPicker: React.FC<{
  provider: Provider;
  value: string;
  onChange: (v: string) => void;
}> = ({ provider, value, onChange }) => {
  const presets = MODEL_PRESETS[provider];
  const recommended = presets.find((p) => p.recommended)?.id ?? presets[0]?.id ?? "";
  return (
    <FieldGroup>
      <FieldLabel>Model</FieldLabel>
      <FieldRow>
        <Input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={recommended}
          list={`models-${provider}`}
          style={{ fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace" }}
        />
        <datalist id={`models-${provider}`}>
          {presets.map((p) => <option key={p.id} value={p.id} label={p.label} />)}
        </datalist>
      </FieldRow>
      <Hint>
        <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
        Leave blank to use the app's built-in default ({recommended}). Or pick from the status-bar model menu. Free-form — type any model id the provider supports.
      </Hint>
    </FieldGroup>
  );
};

export const SettingsModal: React.FC<Props> = ({ open, onClose, tabs }) => {
  const { user, firebaseUser, logout } = useAuth();
  type TabView = "general" | "usage";
  const [view, setView] = useState<TabView>("general");
  const [apiKey, setApiKey] = useState("");
  const [oauth, setOauth] = useState("");
  const [revealKey, setRevealKey] = useState(false);
  const [revealOauth, setRevealOauth] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [codex, setCodex] = useState<CodexStatus | null>(null);
  const [codexBusy, setCodexBusy] = useState(false);
  // Model picker state, per provider. Empty string means "use built-in default".
  const [modelClaude, setModelClaude] = useState("");
  const [modelCodex, setModelCodex] = useState("");
  // Effort per provider (raw provider values; Rust returns the effective
  // value, defaults included, so these are never empty after load).
  const [effortClaude, setEffortClaude] = useState("");
  const [effortCodex, setEffortCodex] = useState("");
  // Auto-apply mode toggle. Same setting the composer toggle writes to.
  const [autoApply, setAutoApplyLocal] = useState(false);
  // Explicit active-provider choice. Drives the api_choice setting, which
  // the agent loop reads to dispatch (Claude vs Codex; API key vs OAuth).
  type ActiveProvider = "claude" | "claude-subscription" | "openai-codex";
  const [activeProvider, setActiveProvider] = useState<ActiveProvider>("claude-subscription");

  const refreshCodex = async () => {
    try { setCodex(await codexStatus()); } catch { /* ignore */ }
  };
  useEffect(() => { if (open) refreshCodex(); }, [open]);

  const handleLogout = async () => {
    if (!confirm("Sign out of GridPath? You'll be sent to the login screen.")) return;
    try {
      await logout();
      onClose();
    } catch (e) {
      console.error("[settings] logout failed:", e);
    }
  };

  const initials = (firebaseUser?.displayName || firebaseUser?.email || user?.email || "?")
    .split(/\s+|@/)
    .filter(Boolean)
    .slice(0, 2)
    .map((s) => s[0]?.toUpperCase())
    .join("");

  // Re-load values whenever the modal opens — the legacy automation screen
  // writes to the same DB, so values could've changed elsewhere.
  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setLoading(true);
    (async () => {
      const [k, o, mc, mco, choice, aa, ec, eco] = await Promise.all([
        getSettingValue(SETTING_KEYS.apiKeyClaude),
        getSettingValue(SETTING_KEYS.apiKeyClaudeOauth),
        getModel("claude"),
        getModel("openai-codex"),
        getSettingValue(SETTING_KEYS.apiChoice),
        getSettingValue(SETTING_KEYS.autoApply),
        getEffort("claude"),
        getEffort("openai-codex"),
      ]);
      if (cancelled) return;
      setApiKey(k);
      setOauth(o);
      setModelClaude(mc);
      setModelCodex(mco);
      setEffortClaude(ec);
      setEffortCodex(eco);
      setAutoApplyLocal(aa === "1");
      // Normalize the stored api_choice to one of our three known values.
      // Fall back to whichever credential the user actually has set.
      const known = ["claude", "claude-subscription", "openai-codex"];
      if (known.includes(choice)) {
        setActiveProvider(choice as ActiveProvider);
      } else if (o) {
        setActiveProvider("claude-subscription");
      } else if (k) {
        setActiveProvider("claude");
      } else {
        setActiveProvider("claude-subscription");
      }
      setDirty(false);
      setLoading(false);
    })();
    return () => { cancelled = true; };
  }, [open]);

  // Esc closes
  useEffect(() => {
    if (!open) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, onClose]);

  if (!open) return null;

  // Connection status reflects the EXPLICITLY selected provider — does the
  // credential for that selection exist?
  const providerLabel: Record<ActiveProvider, string> = {
    "claude": "Anthropic API key",
    "claude-subscription": "Claude subscription",
    "openai-codex": "ChatGPT subscription",
  };
  const credentialPresent: Record<ActiveProvider, boolean> = {
    "claude": !!apiKey.trim(),
    "claude-subscription": !!oauth.trim(),
    "openai-codex": !!codex?.logged_in,
  };
  const hasCredential = credentialPresent[activeProvider];
  const activeMode = providerLabel[activeProvider];

  const onSave = async () => {
    setSaving(true);
    try {
      await setSettingValue(SETTING_KEYS.apiKeyClaude, apiKey.trim());
      await setSettingValue(SETTING_KEYS.apiKeyClaudeOauth, oauth.trim());
      // Persist the explicit provider choice the user made via the radio.
      await setSettingValue(SETTING_KEYS.apiChoice, activeProvider);
      // Model overrides — empty string clears, falling back to built-in default.
      await setModel("claude", modelClaude.trim());
      await setModel("openai-codex", modelCodex.trim());
      await setEffort("claude", effortClaude.trim());
      await setEffort("openai-codex", effortCodex.trim());
      // Auto-apply persisted to the same key the composer toggle writes to.
      await setSettingValue(SETTING_KEYS.autoApply, autoApply ? "1" : "0");
      setDirty(false);
      onClose();
    } catch (e) {
      console.error("[settings] save failed:", e);
      alert(`Save failed: ${e}`);
    } finally {
      setSaving(false);
    }
  };

  return (
    <Backdrop onClick={onClose}>
      <Modal onClick={(e) => e.stopPropagation()}>
        <Header>
          Settings
          <CloseBtn onClick={onClose}><X size={14} /></CloseBtn>
        </Header>
        <TabStrip>
          <TabBtn $active={view === "general"} onClick={() => setView("general")}>General</TabBtn>
          <TabBtn $active={view === "usage"} onClick={() => setView("usage")}>Usage</TabBtn>
        </TabStrip>
        {view === "usage" ? (
          <UsageView tabs={tabs} />
        ) : (
        <Body>
          {(firebaseUser || user) && (
            <>
              <SectionLabel>Account</SectionLabel>
              <AccountCard>
                <AccountHeader>
                  <Avatar>{initials || <UserIcon size={12} />}</Avatar>
                  <div style={{ display: "flex", flexDirection: "column", gap: 1, minWidth: 0 }}>
                    <span style={{ fontSize: 13, fontWeight: 500, color: "#e4e4e4" }}>
                      {firebaseUser?.displayName || "GridPath user"}
                    </span>
                    <span style={{ fontSize: 11, color: "#7c7c7c" }}>
                      Google Account
                    </span>
                  </div>
                </AccountHeader>
                {firebaseUser?.displayName && (
                  <AccountField>
                    <AccountFieldLabel>Name</AccountFieldLabel>
                    <AccountFieldValue>{firebaseUser.displayName}</AccountFieldValue>
                  </AccountField>
                )}
                <AccountField>
                  <AccountFieldLabel>Email</AccountFieldLabel>
                  <AccountFieldValue>{firebaseUser?.email || user?.email || "—"}</AccountFieldValue>
                </AccountField>
                {firebaseUser?.uid && (
                  <AccountField>
                    <AccountFieldLabel>Account ID</AccountFieldLabel>
                    <AccountFieldValue title={firebaseUser.uid}>{firebaseUser.uid}</AccountFieldValue>
                  </AccountField>
                )}
                {user?.userId && (
                  <AccountField>
                    <AccountFieldLabel>GridPath User ID</AccountFieldLabel>
                    <AccountFieldValue title={user.userId}>{user.userId}</AccountFieldValue>
                  </AccountField>
                )}
                <LogoutBtn onClick={handleLogout}>
                  <LogOut size={11} /> Sign out
                </LogoutBtn>
              </AccountCard>
            </>
          )}

          <SectionLabel>Active provider</SectionLabel>
          <ProviderGroup>
            {(["claude-subscription", "claude", "openai-codex"] as ActiveProvider[]).map((p) => (
              <ProviderOption
                key={p}
                $active={activeProvider === p}
                onClick={() => { setActiveProvider(p); setDirty(true); }}
              >
                <Radio $active={activeProvider === p} />
                <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
                  <span style={{ fontSize: 12, color: "#e4e4e4", fontWeight: 500 }}>{providerLabel[p]}</span>
                  <span style={{ fontSize: 10, color: "#7c7c7c" }}>
                    {p === "claude-subscription"
                      ? "Anthropic OAuth · use your Claude Pro/Max plan"
                      : p === "claude"
                      ? "Anthropic API · pay per token"
                      : "OpenAI Codex · use your ChatGPT Plus/Pro plan"}
                  </span>
                </div>
                {!credentialPresent[p] && (
                  <NotConfiguredHint>
                    <NotConfiguredDot />
                    Not set
                  </NotConfiguredHint>
                )}
              </ProviderOption>
            ))}
          </ProviderGroup>

          <Active $ok={hasCredential}>
            {hasCredential ? (
              <>
                <Check size={12} /> Connected · using {activeMode}
              </>
            ) : (
              <>
                <Info size={12} /> {activeMode} selected but its credential isn't configured below. Add one or pick a different provider.
              </>
            )}
          </Active>

          <SectionLabel>Behavior</SectionLabel>
          <ProviderOption
            $active={autoApply}
            onClick={() => { setAutoApplyLocal(!autoApply); setDirty(true); }}
            style={{ background: autoApply ? "rgba(34, 197, 94, 0.08)" : "transparent", borderColor: autoApply ? "rgba(34, 197, 94, 0.3)" : "transparent" }}
          >
            <Radio $active={autoApply} />
            <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
              <span style={{ fontSize: 12, color: "#e4e4e4", fontWeight: 500 }}>Auto-apply agent edits</span>
              <span style={{ fontSize: 10, color: "#7c7c7c" }}>
                Skip the manual Accept / Reject step — the agent's batch lands directly. You can still ⌘Z to undo.
              </span>
            </div>
          </ProviderOption>

          <SectionLabel>Claude</SectionLabel>

          <FieldGroup>
            <FieldLabel>Anthropic API key</FieldLabel>
            <FieldRow>
              <Input
                type={revealKey ? "text" : "password"}
                value={apiKey}
                onChange={(e) => { setApiKey(e.target.value); setDirty(true); }}
                placeholder="sk-ant-api03-…"
                disabled={loading}
              />
              <RevealBtn onClick={() => setRevealKey((v) => !v)} title={revealKey ? "Hide" : "Show"}>
                {revealKey ? <EyeOff size={12} /> : <Eye size={12} />}
              </RevealBtn>
            </FieldRow>
            <Hint>
              <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
              Get one at console.anthropic.com — pay-as-you-go, billed per token.
            </Hint>
          </FieldGroup>

          <FieldGroup>
            <FieldLabel>Claude subscription OAuth token</FieldLabel>
            <FieldRow>
              <Input
                type={revealOauth ? "text" : "password"}
                value={oauth}
                onChange={(e) => { setOauth(e.target.value); setDirty(true); }}
                placeholder="sk-ant-oat01-…"
                disabled={loading}
              />
              <RevealBtn onClick={() => setRevealOauth((v) => !v)} title={revealOauth ? "Hide" : "Show"}>
                {revealOauth ? <EyeOff size={12} /> : <Eye size={12} />}
              </RevealBtn>
            </FieldRow>
            <Hint>
              <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
              Generated by running <code style={{ background: "#161616", padding: "1px 4px", borderRadius: 3 }}>claude setup-token</code> in your terminal (requires a Claude Pro/Max subscription). When set, the agent uses your subscription instead of pay-per-token billing.
            </Hint>
          </FieldGroup>

          <ModelPicker
            provider="claude"
            value={modelClaude}
            onChange={(v) => { setModelClaude(v); setDirty(true); }}
          />

          <EffortPicker
            provider="claude"
            value={effortClaude}
            onChange={(v) => { setEffortClaude(v); setDirty(true); }}
          />

          <Hint style={{ color: "#7c7c7c" }}>
            <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
            If both are set, the OAuth token is preferred. Credentials are stored locally in this app's SQLite database and never leave your machine except to talk to Anthropic.
          </Hint>

          <SectionLabel>ChatGPT (subscription)</SectionLabel>
          <FieldGroup>
            <Hint>
              <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
              Sign in with your ChatGPT Plus or Pro account to use the spreadsheet agent on your existing subscription instead of paying per API token. Opens a browser window to sign in.
            </Hint>
            {codex?.logged_in ? (
              <div style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: 9,
                background: "#161616",
                border: "1px solid #2a2a2a",
                borderRadius: 6,
              }}>
                <Check size={14} color="#86efac" />
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 12, color: "#e4e4e4" }}>Signed in</div>
                  {codex.account_id && (
                    <div style={{ fontSize: 11, color: "#7c7c7c", fontFamily: "ui-monospace, SFMono-Regular, Menlo, monospace", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      Account {codex.account_id}
                    </div>
                  )}
                </div>
                <Btn
                  disabled={codexBusy}
                  onClick={async () => {
                    setCodexBusy(true);
                    try { await codexLogout(); await refreshCodex(); } finally { setCodexBusy(false); }
                  }}
                >
                  Sign out
                </Btn>
              </div>
            ) : (
              <Btn
                $primary
                disabled={codexBusy}
                onClick={async () => {
                  setCodexBusy(true);
                  try {
                    await codexLogin();
                    // After login, set api_choice so the agent uses Codex.
                    await setSettingValue(SETTING_KEYS.apiChoice, "openai-codex");
                    await refreshCodex();
                  } catch (e) {
                    alert(`ChatGPT sign-in failed: ${e}`);
                  } finally {
                    setCodexBusy(false);
                  }
                }}
                style={{ alignSelf: "flex-start" }}
              >
                {codexBusy ? "Opening browser…" : "Sign in with ChatGPT"}
              </Btn>
            )}
          </FieldGroup>

          <ModelPicker
            provider="openai-codex"
            value={modelCodex}
            onChange={(v) => { setModelCodex(v); setDirty(true); }}
          />

          <EffortPicker
            provider="openai-codex"
            value={effortCodex}
            onChange={(v) => { setEffortCodex(v); setDirty(true); }}
          />
        </Body>
        )}
        <Footer>
          <Btn onClick={onClose} disabled={saving}>
            {view === "usage" ? "Close" : "Cancel"}
          </Btn>
          {view === "general" && (
            <Btn $primary onClick={onSave} disabled={!dirty || saving}>
              {saving ? "Saving…" : "Save"}
            </Btn>
          )}
        </Footer>
      </Modal>
    </Backdrop>
  );
};

/**
 * Usage view. Stats are summed from the persistent SpreadsheetSession DB
 * — `spreadsheet_session_add_tokens` already accumulates per-session
 * totals on every agent turn, so this view is just the readback. Open
 * tabs prefer their live in-memory counts (which update mid-turn before
 * the DB write lands); closed/archived sessions show their persisted
 * totals from the DB. The two combine into a single all-time view.
 */
const UsageView: React.FC<{ tabs: WorkbookTab[] }> = ({ tabs }) => {
  const [sessions, setSessions] = useState<SessionRow[]>([]);
  useEffect(() => {
    let cancelled = false;
    const refresh = async () => {
      try {
        const rows = await listSessions(200);
        if (!cancelled) setSessions(rows);
      } catch (e) {
        console.warn("[usage] failed to load sessions:", e);
      }
    };
    refresh();
    // Refresh every 3s while the modal is open so freshly written tokens
    // appear without needing to reopen.
    const id = setInterval(refresh, 3000);
    return () => { cancelled = true; clearInterval(id); };
  }, []);

  // One row per WORKBOOK (the section is "By workbook"): a workbook
  // accumulates a new session row every time it's opened fresh (session id
  // = tab id), so its lifetime spend is the SUM over all its sessions.
  // Per session, prefer the live tab's counts when that session is the one
  // currently open (live counts lead the DB write by a few seconds
  // mid-turn) — matched by session ID, not path. The old path-based match
  // made EVERY historical session of an open workbook mirror the live
  // tab's counts: duplicate rows with identical numbers, prior sessions'
  // real usage hidden, and the live counts double-counted into the
  // lifetime totals.
  type Row = {
    id: string;
    name: string;
    path: string;
    inputTokens: number;
    outputTokens: number;
    cacheReadTokens: number;
    cacheCreationTokens: number;
    open: boolean;
  };
  const openById = new Map(tabs.map((t) => [t.id, t]));
  const byPath = new Map<string, Row>();
  for (const s of sessions) {
    const open = openById.get(s.id);
    const input = open ? open.inputTokens ?? 0 : s.total_input_tokens ?? 0;
    const output = open ? open.outputTokens ?? 0 : s.total_output_tokens ?? 0;
    const cacheRead = open ? open.cacheReadTokens ?? 0 : s.total_cache_read_tokens ?? 0;
    const cacheCreation = open
      ? open.cacheCreationTokens ?? 0
      : s.total_cache_creation_tokens ?? 0;
    const existing = byPath.get(s.workbook_path);
    if (existing) {
      // Sessions arrive newest-first, so the first one seen already named
      // the row; older sessions just add their totals.
      existing.inputTokens += input;
      existing.outputTokens += output;
      existing.cacheReadTokens += cacheRead;
      existing.cacheCreationTokens += cacheCreation;
      existing.open = existing.open || !!open;
    } else {
      byPath.set(s.workbook_path, {
        id: s.id,
        name: (open ? open.name || s.name || open.filename : s.name) || "Untitled",
        path: s.workbook_path,
        inputTokens: input,
        outputTokens: output,
        cacheReadTokens: cacheRead,
        cacheCreationTokens: cacheCreation,
        open: !!open,
      });
    }
  }
  const rows: Row[] = [...byPath.values()];

  const totalIn = rows.reduce((s, r) => s + r.inputTokens, 0);
  const totalOut = rows.reduce((s, r) => s + r.outputTokens, 0);
  const totalCacheRead = rows.reduce((s, r) => s + r.cacheReadTokens, 0);
  const totalCacheCreation = rows.reduce((s, r) => s + r.cacheCreationTokens, 0);
  // Cache-hit ratio against total prompt tokens billed at full rate
  // (input_tokens already excludes cache reads in Anthropic's accounting).
  const cacheDenominator = totalIn + totalCacheRead + totalCacheCreation;
  const cacheHitPct = cacheDenominator > 0
    ? Math.round((totalCacheRead / cacheDenominator) * 100)
    : 0;
  const fmt = (n: number) => n.toLocaleString();

  return (
    <Body>
      <SectionLabel>Lifetime usage</SectionLabel>
      <UsageSummary>
        <UsageStat>
          <UsageStatLabel>Input tokens</UsageStatLabel>
          <UsageStatValue>{fmt(totalIn)}</UsageStatValue>
          <UsageStatHint>Sent to the model across all sessions</UsageStatHint>
        </UsageStat>
        <UsageStat>
          <UsageStatLabel>Output tokens</UsageStatLabel>
          <UsageStatValue>{fmt(totalOut)}</UsageStatValue>
          <UsageStatHint>Generated by the agent</UsageStatHint>
        </UsageStat>
        <UsageStat>
          <UsageStatLabel>Cache read</UsageStatLabel>
          <UsageStatValue>{fmt(totalCacheRead)}</UsageStatValue>
          <UsageStatHint>~90% cheaper than fresh input</UsageStatHint>
        </UsageStat>
        <UsageStat>
          <UsageStatLabel>Cache writes</UsageStatLabel>
          <UsageStatValue>{fmt(totalCacheCreation)}</UsageStatValue>
          <UsageStatHint>One-time per ~5min TTL refresh</UsageStatHint>
        </UsageStat>
      </UsageSummary>

      <Hint>
        <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
        Prompt caching is automatic on the Claude provider — the system
        prompt + tools schema are served from Anthropic's cache on every turn
        after the first. Lifetime cache hit rate: <strong style={{ color: "#86efac" }}>{cacheHitPct}%</strong>.
        ChatGPT (Codex) doesn't expose cache stats.
      </Hint>

      <SectionLabel>By workbook</SectionLabel>
      {rows.length === 0 ? (
        <Hint>
          <Info size={11} style={{ marginTop: 1, flexShrink: 0 }} />
          No sessions yet. Token usage is tracked per workbook session and
          persisted across app restarts.
        </Hint>
      ) : (
        <UsageTable>
          <UsageRow $head>
            <UsageCell>Workbook</UsageCell>
            <UsageCell $right>Input</UsageCell>
            <UsageCell $right>Output</UsageCell>
            <UsageCell $right>Cache read</UsageCell>
            <UsageCell $right>Cache write</UsageCell>
          </UsageRow>
          {rows.map((r) => (
            <UsageRow key={r.id}>
              <UsageCell title={r.path}>
                {r.name}{r.open ? " · open" : ""}
              </UsageCell>
              <UsageCell $mono $right $dim={!r.inputTokens}>{fmt(r.inputTokens)}</UsageCell>
              <UsageCell $mono $right $dim={!r.outputTokens}>{fmt(r.outputTokens)}</UsageCell>
              <UsageCell $mono $right $dim={!r.cacheReadTokens}>{fmt(r.cacheReadTokens)}</UsageCell>
              <UsageCell $mono $right $dim={!r.cacheCreationTokens}>{fmt(r.cacheCreationTokens)}</UsageCell>
            </UsageRow>
          ))}
        </UsageTable>
      )}
    </Body>
  );
};
