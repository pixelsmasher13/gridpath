import React, { useEffect, useState } from "react";
import styled from "styled-components";
import { Eye, EyeOff, ArrowRight, Sheet, Check, Copy, CheckCheck } from "lucide-react";
import { invoke } from "@tauri-apps/api/core";
import { getSettingValue, setSettingValue, SETTING_KEYS } from "../settingsApi";
import { codexLogin, codexStatus, type CodexStatus } from "../codexAuth";

const Page = styled.div`
  min-height: 100vh;
  background: #1e1e1e;
  color: #d4d4d4;
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: -apple-system, "SF Pro Text", "Inter", system-ui, sans-serif;
  padding: 24px;
`;

const Card = styled.div`
  width: 100%;
  max-width: 520px;
  display: flex;
  flex-direction: column;
  gap: 24px;
`;

const Brand = styled.div`
  display: flex;
  align-items: center;
  gap: 10px;
  color: #e4e4e4;
  font-weight: 600;
  font-size: 14px;
  letter-spacing: 0.2px;
`;

const BrandMark = styled.div`
  width: 28px;
  height: 28px;
  border-radius: 6px;
  background: linear-gradient(135deg, #3363AD, #5687cf);
  display: flex;
  align-items: center;
  justify-content: center;
  color: #fff;
`;

const Title = styled.h1`
  font-size: 28px;
  font-weight: 600;
  color: #f4f4f4;
  margin: 0;
  line-height: 1.25;
`;

const Subtitle = styled.div`
  font-size: 14px;
  color: #9b9b9b;
  line-height: 1.5;
`;

const Tabs = styled.div`
  display: inline-flex;
  background: #161616;
  border: 1px solid #2a2a2a;
  border-radius: 6px;
  padding: 3px;
  align-self: flex-start;
`;

const Tab = styled.button<{ $active: boolean }>`
  background: ${(p) => (p.$active ? "#2a2a2a" : "transparent")};
  color: ${(p) => (p.$active ? "#e4e4e4" : "#9b9b9b")};
  border: 0;
  border-radius: 4px;
  padding: 5px 12px;
  font-size: 12px;
  font-weight: 500;
  cursor: pointer;
  &:hover { color: #e4e4e4; }
`;

const FieldLabel = styled.label`
  font-size: 12px;
  color: #b3b3b3;
  font-weight: 500;
`;

const FieldRow = styled.div`
  display: flex;
  gap: 6px;
`;

const Input = styled.input`
  flex: 1;
  background: #161616;
  border: 1px solid #333;
  border-radius: 5px;
  padding: 9px 11px;
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

const HelpText = styled.div`
  font-size: 12px;
  color: #7c7c7c;
  line-height: 1.5;
  code {
    background: #161616;
    padding: 1px 5px;
    border-radius: 3px;
    color: #c4c4c4;
  }
`;

const CommandRow = styled.div`
  display: flex;
  align-items: center;
  gap: 6px;
  background: #161616;
  border: 1px solid #2a2a2a;
  border-radius: 5px;
  padding: 7px 8px 7px 11px;
`;

const CommandCode = styled.code`
  flex: 1;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  font-size: 12px;
  color: #e4e4e4;
  background: transparent !important;
  padding: 0 !important;
`;

const CopyCmdBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  background: #2a2a2a;
  border: 1px solid #333;
  border-radius: 4px;
  color: #d4d4d4;
  font-size: 11px;
  font-weight: 500;
  padding: 5px 9px;
  cursor: pointer;
  white-space: nowrap;
  &:hover { background: #333; color: #fff; }
`;

const ErrorBanner = styled.div`
  font-size: 12px;
  line-height: 1.45;
  color: #fecaca;
  background: rgba(239, 68, 68, 0.1);
  border: 1px solid rgba(239, 68, 68, 0.35);
  border-radius: 6px;
  padding: 9px 11px;
`;

const Actions = styled.div`
  display: flex;
  align-items: center;
  gap: 12px;
  margin-top: 8px;
`;

const PrimaryBtn = styled.button`
  background: #3363AD;
  color: #fff;
  border: 0;
  border-radius: 5px;
  padding: 9px 18px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  &:hover:not(:disabled) { background: #4275c4; }
  &:disabled { opacity: 0.45; cursor: not-allowed; }
`;

const SkipBtn = styled.button`
  background: transparent;
  color: #7c7c7c;
  border: 0;
  font-size: 12px;
  cursor: pointer;
  padding: 6px 8px;
  &:hover { color: #c4c4c4; }
`;

const SETUP_TOKEN_CMD = "claude setup-token";

interface Props {
  onComplete: () => void;
}

export const SpreadsheetOnboarding: React.FC<Props> = ({ onComplete }) => {
  // Pre-load any existing credentials so the user sees what's there.
  const [mode, setMode] = useState<"subscription" | "chatgpt" | "api">("subscription");
  const [apiKey, setApiKey] = useState("");
  const [oauth, setOauth] = useState("");
  const [reveal, setReveal] = useState(false);
  const [saving, setSaving] = useState(false);
  const [codex, setCodex] = useState<CodexStatus | null>(null);
  const [copied, setCopied] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    (async () => {
      const [k, o, c] = await Promise.all([
        getSettingValue(SETTING_KEYS.apiKeyClaude),
        getSettingValue(SETTING_KEYS.apiKeyClaudeOauth),
        codexStatus().catch(() => null),
      ]);
      setApiKey(k);
      setOauth(o);
      setCodex(c);
      // Pick the tab that already has a value — subscription wins ties,
      // then ChatGPT, then API key.
      if (o) setMode("subscription");
      else if (c?.logged_in) setMode("chatgpt");
      else if (k) setMode("api");
    })();
  }, []);

  // Clear stale errors when the user switches tabs or edits the credential.
  useEffect(() => {
    setError(null);
  }, [mode, apiKey, oauth]);

  const canContinue =
    (mode === "subscription" && !!oauth.trim()) ||
    (mode === "api" && !!apiKey.trim()) ||
    (mode === "chatgpt" && !!codex?.logged_in);

  const copySetupToken = async () => {
    try {
      await navigator.clipboard.writeText(SETUP_TOKEN_CMD);
      setCopied(true);
      window.setTimeout(() => setCopied(false), 1800);
    } catch (e) {
      console.warn("[onboarding] clipboard write failed:", e);
      setError("Couldn't copy — select the command and copy manually.");
    }
  };

  const onContinue = async () => {
    if (!canContinue) return;
    setSaving(true);
    setError(null);
    try {
      // Prove the credential works before marking onboarding done — otherwise
      // users land in the app with a dead agent and an opaque first-prompt error.
      if (mode === "api") {
        await invoke("validate_llm_credential", {
          kind: "claude",
          credential: apiKey.trim(),
        });
        await setSettingValue(SETTING_KEYS.apiKeyClaude, apiKey.trim());
        await setSettingValue(SETTING_KEYS.apiChoice, "claude");
      } else if (mode === "subscription") {
        await invoke("validate_llm_credential", {
          kind: "claude-subscription",
          credential: oauth.trim(),
        });
        await setSettingValue(SETTING_KEYS.apiKeyClaudeOauth, oauth.trim());
        await setSettingValue(SETTING_KEYS.apiChoice, "claude-subscription");
      } else if (mode === "chatgpt") {
        await invoke("validate_llm_credential", {
          kind: "openai-codex",
          credential: null,
        });
        await setSettingValue(SETTING_KEYS.apiChoice, "openai-codex");
      }
      onComplete();
    } catch (e) {
      console.error("[onboarding] validate/save failed:", e);
      setError(String(e));
    } finally {
      setSaving(false);
    }
  };

  const onChatGPTLogin = async () => {
    setSaving(true);
    setError(null);
    try {
      await codexLogin();
      const c = await codexStatus().catch(() => null);
      setCodex(c);
      // Persist the provider choice on successful sign-in (not just on
      // Continue) so the credential is actually used even if the user
      // closes the window before pressing Continue — matches Settings.
      if (c?.logged_in) {
        await setSettingValue(SETTING_KEYS.apiChoice, "openai-codex");
      }
    } catch (e) {
      console.error("[onboarding] ChatGPT login failed:", e);
      setError(`ChatGPT sign-in failed: ${e}`);
    } finally {
      setSaving(false);
    }
  };

  // Copy adapts to the selected provider so we never tell a ChatGPT user
  // we're "connecting Claude" or sending their token to Anthropic.
  const providerLabel = mode === "chatgpt" ? "ChatGPT" : "Claude";
  const providerVendor = mode === "chatgpt" ? "OpenAI" : "Anthropic";

  const onSkip = () => {
    const ok = window.confirm(
      "Skip setup for now?\n\nGridPath can't edit spreadsheets until you connect Claude or ChatGPT. " +
        "You can add a credential anytime from Settings."
    );
    if (ok) onComplete();
  };

  return (
    <Page>
      <Card>
        <Brand>
          <BrandMark><Sheet size={16} /></BrandMark>
          GridPath
        </Brand>

        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          <Title>Connect {providerLabel}</Title>
          <Subtitle>
            GridPath uses {providerLabel} to edit your spreadsheets. Paste a credential to get
            started — you can change it anytime in Settings.
          </Subtitle>
        </div>

        <Tabs>
          <Tab $active={mode === "subscription"} onClick={() => setMode("subscription")}>
            Claude Pro / Max
          </Tab>
          <Tab $active={mode === "chatgpt"} onClick={() => setMode("chatgpt")}>
            ChatGPT Plus / Pro
          </Tab>
          <Tab $active={mode === "api"} onClick={() => setMode("api")}>
            Anthropic API key
          </Tab>
        </Tabs>

        {mode === "chatgpt" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <HelpText>
              Sign in with your ChatGPT Plus or Pro account. GridPath uses
              the OpenAI Codex Responses API on your existing subscription — no
              extra billing per token. The sign-in opens a browser window.
            </HelpText>
            {codex?.logged_in ? (
              <div style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                padding: "10px 12px",
                background: "rgba(34, 197, 94, 0.08)",
                border: "1px solid rgba(34, 197, 94, 0.3)",
                color: "#86efac",
                borderRadius: 6,
                fontSize: 12,
              }}>
                <Check size={12} />
                <span>Signed in</span>
                {codex.account_id && (
                  <span style={{ marginLeft: "auto", fontFamily: "ui-monospace, monospace", fontSize: 11, color: "#86efac", opacity: 0.7 }}>
                    Account {codex.account_id.slice(0, 12)}…
                  </span>
                )}
              </div>
            ) : (
              <button
                onClick={onChatGPTLogin}
                disabled={saving}
                style={{
                  alignSelf: "flex-start",
                  background: "#3363AD",
                  color: "#fff",
                  border: 0,
                  borderRadius: 5,
                  padding: "9px 16px",
                  fontSize: 13,
                  fontWeight: 500,
                  cursor: "pointer",
                  opacity: saving ? 0.45 : 1,
                }}
              >
                {saving ? "Opening browser…" : "Sign in with ChatGPT"}
              </button>
            )}
          </div>
        ) : mode === "subscription" ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <HelpText>
              In a terminal, run this command (requires the Claude Code CLI and a
              Claude Pro or Max subscription), then paste the token below.
            </HelpText>
            <CommandRow>
              <CommandCode>{SETUP_TOKEN_CMD}</CommandCode>
              <CopyCmdBtn type="button" onClick={copySetupToken} title="Copy command">
                {copied ? <CheckCheck size={12} /> : <Copy size={12} />}
                {copied ? "Copied" : "Copy"}
              </CopyCmdBtn>
            </CommandRow>
            <FieldLabel>OAuth token</FieldLabel>
            <FieldRow>
              <Input
                type={reveal ? "text" : "password"}
                placeholder="sk-ant-oat01-…"
                value={oauth}
                onChange={(e) => setOauth(e.target.value)}
                autoFocus
              />
              <RevealBtn onClick={() => setReveal((v) => !v)} title={reveal ? "Hide" : "Show"}>
                {reveal ? <EyeOff size={12} /> : <Eye size={12} />}
              </RevealBtn>
            </FieldRow>
            <HelpText>
              No extra billing — you use your existing Claude Pro or Max subscription.
              We'll verify the token before continuing.
            </HelpText>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            <FieldLabel>API key</FieldLabel>
            <FieldRow>
              <Input
                type={reveal ? "text" : "password"}
                placeholder="sk-ant-api03-…"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                autoFocus
              />
              <RevealBtn onClick={() => setReveal((v) => !v)} title={reveal ? "Hide" : "Show"}>
                {reveal ? <EyeOff size={12} /> : <Eye size={12} />}
              </RevealBtn>
            </FieldRow>
            <HelpText>
              Get one at <code>console.anthropic.com</code> — pay-as-you-go, billed per token.
              We'll verify the key before continuing.
            </HelpText>
          </div>
        )}

        {error && <ErrorBanner>{error}</ErrorBanner>}

        <Actions>
          <PrimaryBtn onClick={onContinue} disabled={!canContinue || saving}>
            {saving ? "Verifying…" : "Continue"}
            <ArrowRight size={13} />
          </PrimaryBtn>
          <SkipBtn onClick={onSkip} disabled={saving}>Skip for now</SkipBtn>
        </Actions>

        <HelpText style={{ marginTop: 8 }}>
          Credentials are stored locally on this machine and never leave except to talk to {providerVendor}.
        </HelpText>
      </Card>
    </Page>
  );
};
