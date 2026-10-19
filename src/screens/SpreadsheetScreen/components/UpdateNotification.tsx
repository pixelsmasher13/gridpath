import React, { useEffect, useState } from "react";
import styled, { keyframes } from "styled-components";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { X, Download, RotateCw } from "lucide-react";

/**
 * Update notification toast — Cursor-style bottom-right card.
 *
 * Wire-up:
 *   1. Rust's `check_for_update_on_startup` (main.rs) polls latest.json
 *      and emits `update-available` with { version, date, body }.
 *   2. This component mounts globally, listens for that event,
 *      and shows the toast.
 *   3. "Install & Restart" invokes `install_update` (downloads + applies
 *      the .app.tar.gz in place), then `relaunch_app` (Rust restart).
 *   4. "Later" hides the toast for this session; user gets re-prompted
 *      on next launch if the update is still available.
 */
interface UpdatePayload {
  version: string;
  date?: string | null;
  body?: string | null;
}

type Phase = "idle" | "available" | "installing" | "ready_to_restart" | "error";

export const UpdateNotification: React.FC = () => {
  const [phase, setPhase] = useState<Phase>("idle");
  const [info, setInfo] = useState<UpdatePayload | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let unlistenFn: (() => void) | undefined;
    (async () => {
      try {
        unlistenFn = await listen<UpdatePayload>("update-available", (e) => {
          setInfo(e.payload);
          setPhase("available");
        });
      } catch (e) {
        console.warn("[update] failed to install update-available listener:", e);
      }
    })();
    return () => { unlistenFn?.(); };
  }, []);

  if (phase === "idle" || !info) return null;

  const onInstall = async () => {
    setPhase("installing");
    setError(null);
    try {
      await invoke("install_update");
      setPhase("ready_to_restart");
    } catch (e) {
      console.error("[update] install failed:", e);
      setError(String(e));
      setPhase("error");
    }
  };

  const onRestart = async () => {
    try {
      await invoke("relaunch_app");
    } catch (e) {
      console.error("[update] relaunch failed:", e);
      setError(String(e));
      setPhase("error");
    }
  };

  const onDismiss = () => {
    setPhase("idle");
    setInfo(null);
  };

  return (
    <Card role="status" aria-live="polite">
      <Header>
        <Title>
          {phase === "ready_to_restart" ? "Update ready to install" : "GridPath update available"}
        </Title>
        <DismissBtn onClick={onDismiss} title="Dismiss" aria-label="Dismiss update notification">
          <X size={12} />
        </DismissBtn>
      </Header>

      <Version>v{info.version}</Version>

      {phase === "available" && info.body && (
        <Notes>{truncate(info.body, 220)}</Notes>
      )}

      {phase === "installing" && (
        <Status>
          <Spinner size={11} />
          Downloading and installing…
        </Status>
      )}

      {phase === "ready_to_restart" && (
        <Notes>Restart GridPath to start using the new version.</Notes>
      )}

      {phase === "error" && error && (
        <ErrorText>Update failed: {error}</ErrorText>
      )}

      <Actions>
        {phase === "available" && (
          <>
            <SecondaryBtn onClick={onDismiss}>Later</SecondaryBtn>
            <PrimaryBtn onClick={onInstall}>
              <Download size={12} />
              Install &amp; restart
            </PrimaryBtn>
          </>
        )}
        {phase === "installing" && (
          <SecondaryBtn disabled>Installing…</SecondaryBtn>
        )}
        {phase === "ready_to_restart" && (
          <PrimaryBtn onClick={onRestart}>
            <RotateCw size={12} />
            Restart now
          </PrimaryBtn>
        )}
        {phase === "error" && (
          <>
            <SecondaryBtn onClick={onDismiss}>Close</SecondaryBtn>
            <PrimaryBtn onClick={onInstall}>Retry</PrimaryBtn>
          </>
        )}
      </Actions>
    </Card>
  );
};

function truncate(s: string, n: number): string {
  if (s.length <= n) return s;
  return s.slice(0, n - 1).trimEnd() + "…";
}

const slideIn = keyframes`
  from { transform: translateY(20px); opacity: 0; }
  to   { transform: translateY(0);   opacity: 1; }
`;

const spin = keyframes`
  from { transform: rotate(0deg); }
  to   { transform: rotate(360deg); }
`;

const Card = styled.div`
  position: fixed;
  right: 16px;
  bottom: 36px;
  z-index: 200;
  width: 340px;
  background: #1e1e1e;
  border: 1px solid #333;
  border-radius: 10px;
  padding: 14px 14px 12px;
  box-shadow: 0 12px 32px rgba(0, 0, 0, 0.4);
  color: #e4e4e4;
  font-family: -apple-system, BlinkMacSystemFont, "Inter", system-ui, sans-serif;
  animation: ${slideIn} 200ms ease-out;
`;

const Header = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 4px;
`;

const Title = styled.div`
  font-size: 13px;
  font-weight: 600;
  letter-spacing: -0.01em;
`;

const DismissBtn = styled.button`
  background: transparent;
  border: none;
  color: #7c7c7c;
  cursor: pointer;
  padding: 4px;
  border-radius: 4px;
  display: inline-flex;
  align-items: center;
  &:hover { color: #e4e4e4; background: rgba(255, 255, 255, 0.06); }
`;

const Version = styled.div`
  font-size: 11px;
  color: #7c7c7c;
  font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
  margin-bottom: 8px;
`;

const Notes = styled.div`
  font-size: 12px;
  color: #b3b3b3;
  line-height: 1.45;
  margin-bottom: 12px;
  max-height: 90px;
  overflow: hidden;
`;

const Status = styled.div`
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
  color: #b3b3b3;
  margin-bottom: 12px;
`;

const ErrorText = styled.div`
  font-size: 12px;
  color: #f87171;
  margin-bottom: 12px;
  line-height: 1.45;
  word-break: break-word;
`;

const Actions = styled.div`
  display: flex;
  justify-content: flex-end;
  gap: 8px;
`;

const Btn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-size: 12px;
  font-weight: 500;
  padding: 6px 12px;
  border-radius: 6px;
  cursor: pointer;
  border: 1px solid transparent;
  font-family: inherit;
  &:disabled { opacity: 0.6; cursor: default; }
`;

const PrimaryBtn = styled(Btn)`
  background: #2563eb;
  color: #fff;
  &:hover:not(:disabled) { background: #1d4ed8; }
`;

const SecondaryBtn = styled(Btn)`
  background: transparent;
  color: #b3b3b3;
  border-color: #3a3a3a;
  &:hover:not(:disabled) { color: #e4e4e4; border-color: #4a4a4a; }
`;

const Spinner = styled(RotateCw)`
  animation: ${spin} 1s linear infinite;
`;
