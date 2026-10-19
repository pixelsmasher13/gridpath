use log::{error, info, warn};
use once_cell::sync::Lazy;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::{HashMap, HashSet, VecDeque};
use std::sync::{Arc, Mutex};
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};
use std::time::Duration;
use tauri::{AppHandle, Emitter, Manager};
use tokio::sync::oneshot;
use tokio::time::timeout;

use crate::configuration::state::ServiceAccess;
use crate::engine::llm_providers::claude::{
    stream_claude_with_tools, StreamEvent, StreamMessage,
};
use crate::engine::llm_providers::openai_codex::stream_codex_with_tools;
use crate::engine::spreadsheet_agent::tools::{agent_tools, system_prompt};
use crate::repository::settings_repository::get_setting;

/// Workbook context shipped from the webview each turn. Kept minimal so this
/// stays a thin v1 wrapper — the webview owns the workbook model and decides
/// what to ship as context (sheet names, used range, sample cells).
#[derive(Debug, Deserialize)]
pub struct WorkbookContext {
    pub path: String,
    pub sheets: Vec<SheetContext>,
    /// Optional "User focus" block — selection + @-mentions, serialized by
    /// the webview at submit time. When present we inject it into the user
    /// message above the prompt so the agent treats it as ground truth.
    #[serde(default)]
    pub focus: Option<String>,
    /// Changed-cells block from the webview's base+delta capture: `sheets`
    /// still carries the byte-identical cached base preview; this lists the
    /// cells edited since that snapshot. Rendered into the UNCACHED turn
    /// tail — putting it in the context block would defeat the whole point
    /// (one changed byte re-bills the cached prefix).
    #[serde(default)]
    pub delta: Option<String>,
    /// One-line live-vs-file-saved divergence warning from the webview's
    /// calc-health scan (engine can't evaluate INDIRECT/external refs/add-ins
    /// on some imported models). Computed once per loaded workbook so its
    /// bytes are stable — rendered inside the CACHED context block.
    #[serde(default)]
    pub calc_health: Option<String>,
    /// Read-only reference workbooks the user attached to the session (other
    /// xlsx files, e.g. analyst models to compare against). Each ships a
    /// compact preview; detail is pulled on demand via `read_reference`.
    #[serde(default)]
    pub references: Vec<ReferenceWorkbookContext>,
}

#[derive(Debug, Deserialize)]
pub struct SheetContext {
    pub name: String,
    pub row_count: u32,
    pub column_count: u32,
    /// Used data extent in A1 ("A1:Q48"). None on empty sheets or payloads
    /// from older frontends (fall back to row/column counts).
    #[serde(default)]
    pub used_range: Option<String>,
    /// Compact "A1 = 42" rows, optionally with `[fmt …]` annotations and a
    /// leading "Style conventions" palette. The webview decides how to
    /// truncate so we don't have to encode windowing logic in Rust.
    pub cells_preview: String,
}

#[derive(Debug, Deserialize)]
pub struct ReferenceWorkbookContext {
    pub path: String,
    /// Display label the agent addresses this workbook by in
    /// `read_reference` calls — the filename, e.g. "AAPL-GS.xlsx".
    pub label: String,
    pub sheets: Vec<SheetContext>,
}

/// Payload broadcast for every agent event. The frontend dispatches by
/// `tab_id` + `batch_id`. Keeping a single event name (`spreadsheet:event`)
/// + a discriminator field makes parallel agents and Tauri listener cleanup
/// straightforward — one listener per tab is enough.
#[derive(Debug, Serialize, Clone)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum AgentEvent {
    Started {
        tab_id: String,
        batch_id: String,
    },
    TextDelta {
        tab_id: String,
        batch_id: String,
        delta: String,
    },
    /// A chunk of the model's reasoning/plan. Rendered in the UI as a
    /// collapsible "Plan" block, separate from user-facing prose (`TextDelta`).
    Reasoning {
        tab_id: String,
        batch_id: String,
        delta: String,
    },
    ToolCall {
        tab_id: String,
        batch_id: String,
        tool_use_id: String,
        name: String,
        input: Value,
    },
    Done {
        tab_id: String,
        batch_id: String,
        stop_reason: String,
        input_tokens: u32,
        output_tokens: u32,
        /// Tokens served from Anthropic's prompt-cache this run (sum across turns).
        cache_read_tokens: u32,
        /// Tokens written into the cache this run (cache miss → refresh).
        cache_creation_tokens: u32,
    },
    Error {
        tab_id: String,
        batch_id: String,
        message: String,
    },
}

const EVENT_NAME: &str = "spreadsheet:event";

/// Global registry of in-flight agent runs keyed by batch_id. The agent loop
/// checks the AtomicBool between turns and bails if the user clicked Stop.
/// We can't cancel an in-flight HTTP body read mid-token without dropping
/// the response, so cancellation is granular at the *turn* boundary, not the
/// individual SSE chunk. That's fine — turns are short.
static CANCEL_TOKENS: Lazy<Mutex<HashMap<String, Arc<AtomicBool>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

fn register_cancel_token(batch_id: &str) -> Arc<AtomicBool> {
    let token = Arc::new(AtomicBool::new(false));
    if let Ok(mut map) = CANCEL_TOKENS.lock() {
        map.insert(batch_id.to_string(), token.clone());
    }
    token
}

fn unregister_cancel_token(batch_id: &str) {
    if let Ok(mut map) = CANCEL_TOKENS.lock() {
        map.remove(batch_id);
    }
}

/// Registry of pending tool-result senders, keyed by tool_use_id. The agent
/// loop registers a oneshot::Sender before emitting a tool_call event; the
/// webview later calls `spreadsheet_tool_result` to deliver the evaluated
/// cell values. The loop awaits all expected results (with a timeout) before
/// composing the next turn's user message.
static TOOL_RESULT_SENDERS: Lazy<Mutex<HashMap<String, oneshot::Sender<String>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Registry of pending "the webview dequeued this tool" acks, keyed the same
/// way. Tool calls run through a per-tab SERIAL queue in the webview (see
/// `toolCallQueueByTab`), so a tool emitted now can sit behind several others
/// before it starts executing — see `await_tool_result` for why that
/// distinction is load-bearing.
static TOOL_START_SENDERS: Lazy<Mutex<HashMap<String, oneshot::Sender<()>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// The pair of receivers the loop holds for one in-flight webview tool call:
/// `start` fires when the webview begins executing it, `result` when it
/// finishes and reports back.
struct ToolWaiters {
    start: oneshot::Receiver<()>,
    result: oneshot::Receiver<String>,
}

fn register_tool_result(tool_use_id: &str) -> ToolWaiters {
    let (result_tx, result) = oneshot::channel();
    let (start_tx, start) = oneshot::channel();
    if let Ok(mut map) = TOOL_RESULT_SENDERS.lock() {
        // If a sender already exists for this id (would be weird), drop it.
        map.insert(tool_use_id.to_string(), result_tx);
    }
    if let Ok(mut map) = TOOL_START_SENDERS.lock() {
        map.insert(tool_use_id.to_string(), start_tx);
    }
    ToolWaiters { start, result }
}

/// Drop both pending channels for a tool_use_id (turn teardown). Without the
/// start map being cleared too, an abandoned turn's acks would accumulate.
fn unregister_tool_waiters(tool_use_id: &str) {
    if let Ok(mut map) = TOOL_RESULT_SENDERS.lock() {
        map.remove(tool_use_id);
    }
    if let Ok(mut map) = TOOL_START_SENDERS.lock() {
        map.remove(tool_use_id);
    }
}

/// Called by the webview the moment a tool_call task reaches the head of its
/// per-tab queue and begins executing — NOT when the event arrives. This is
/// what lets the loop bill queue time and execution time separately.
#[tauri::command]
pub async fn spreadsheet_tool_started(tool_use_id: String) -> Result<(), String> {
    if let Ok(mut map) = TOOL_START_SENDERS.lock() {
        if let Some(tx) = map.remove(&tool_use_id) {
            let _ = tx.send(());
        }
    }
    // No registered sender: either a skip-readback tool (the loop never waits
    // on those) or a turn that already tore down. Nothing to do either way.
    Ok(())
}

#[tauri::command]
pub async fn spreadsheet_tool_result(tool_use_id: String, content: String) -> Result<(), String> {
    if let Ok(mut map) = TOOL_RESULT_SENDERS.lock() {
        if let Some(tx) = map.remove(&tool_use_id) {
            // If the receiver was dropped (turn already moved on), the send
            // returns Err — that's fine, we just discard.
            let _ = tx.send(content);
            return Ok(());
        }
    }
    // Late delivery: the loop already gave up on this tool and moved on, so
    // the webview just did work whose result nobody will ever see. Loud on
    // purpose — a silent drop here hides a full loop/webview desync (the grid
    // keeps updating correctly while the model sees a placeholder for every
    // read), which is exactly the failure the two-phase wait below prevents.
    warn!(
        "spreadsheet_tool_result: no pending sender for {} — result discarded ({}B). \
         The loop timed out on this tool before the webview answered.",
        tool_use_id,
        content.len()
    );
    Ok(())
}

/// How long the loop will wait for the webview to merely START a tool, on top
/// of that tool's own execution budget. Sized for a deep serial queue (a turn
/// can emit ~10 tools and each must drain in order), not for execution.
const TOOL_QUEUE_WAIT_SECS: u64 = 180;

/// Two-phase wait for one webview tool result.
///
/// The webview applies tool calls through a per-tab SERIAL queue while this
/// loop awaits every tool of a turn CONCURRENTLY. So the wall-clock until the
/// k-th tool of a batch answers is roughly the sum of the k-1 before it —
/// but each tool used to get only its own execution budget, measured from
/// emit. Any batch that ran deeper than a tool or two blew that budget, and
/// the fallout was permanent rather than local: the loop wrote a synthetic
/// result, tore down the sender, and started the next turn while the webview
/// was still draining the previous one, so from then on every correct answer
/// arrived at a deregistered id and was dropped. The grid kept updating
/// perfectly; the model saw a placeholder for every read for the rest of the
/// run and concluded its own read tooling was broken.
///
/// This is the root cause behind two earlier symptom fixes: the auto-snapshot
/// guard on `pendingToolCalls` ("cascading 25-28s timeouts") and the tool-task
/// crash wrapper ("the agent built on writes that never landed"). Both removed
/// a *source* of queue delay; neither fixed the accounting.
///
/// Splitting the wait fixes it: queue time is billed to `TOOL_QUEUE_WAIT_SECS`
/// (shared across the batch, since the waits overlap) and only execution is
/// billed to the per-tool budget it was sized for. Both phases tick at 1s and
/// re-check the cancel token: cancellation is otherwise turn-granular, and a
/// queue wait is long enough that a user pressing Stop must not sit through it.
/// Returns None when the webview never started or never answered.
async fn await_tool_result(
    id: &str,
    name: &str,
    waiters: ToolWaiters,
    exec_secs: u64,
    cancel: &AtomicBool,
) -> Option<String> {
    let ToolWaiters { start, result } = waiters;
    let mut start = start;
    let mut result = result;

    // Phase 1 — queued, not yet running. Billed to the batch-wide queue
    // budget, since every tool of the turn is waiting here concurrently.
    let mut acked = false;
    for _ in 0..TOOL_QUEUE_WAIT_SECS {
        if cancel.load(Ordering::SeqCst) {
            return None;
        }
        tokio::select! {
            biased;
            // A result can legitimately beat its own ack — any path that
            // answers straight from the event handler without dequeuing.
            // Polling `&mut result` never consumes the channel, so losing
            // this race costs nothing.
            r = &mut result => return r.ok(),
            tick = timeout(Duration::from_secs(1), &mut start) => match tick {
                Ok(Ok(())) => { acked = true; break; }
                Ok(Err(_)) => return None,
                Err(_) => continue,
            },
        }
    }
    if !acked {
        warn!(
            "agent_loop: {} ({}) was never dequeued by the webview after {}s \
             — its tool queue is stalled or the view unmounted",
            id, name, TOOL_QUEUE_WAIT_SECS
        );
        return None;
    }

    // Phase 2 — actually executing. This is the interval the per-tool budget
    // in readback_timeout_secs was sized against.
    for _ in 0..exec_secs {
        if cancel.load(Ordering::SeqCst) {
            return None;
        }
        match timeout(Duration::from_secs(1), &mut result).await {
            Ok(Ok(payload)) => return Some(payload),
            Ok(Err(_)) => {
                warn!("agent_loop: tool_result sender for {} dropped", id);
                return None;
            }
            Err(_) => continue,
        }
    }
    warn!(
        "agent_loop: tool_result for {} ({}) timed out after {}s of execution",
        id, name, exec_secs
    );
    None
}

/// Substituted when a webview tool never reports back. Deliberately NOT "ok":
/// that read as a successful call returning nothing, which is indistinguishable
/// from a successful read of an empty range — the model treated timed-out reads
/// as evidence that its work had vanished. It must also not invite a blind
/// retry, since a timed-out WRITE has very often applied.
fn readback_timeout_msg(waited_secs: u64) -> String {
    json!({
        "ok": false,
        "error": "readback_timeout",
        "message": format!(
            "The spreadsheet did not report back within {waited_secs}s. This says nothing \
             about whether the operation succeeded — it very likely applied. Do NOT blindly \
             repeat a write; read the target range first to see its current state, and if \
             reads keep timing out, say so plainly rather than assuming the data is missing."
        ),
    })
    .to_string()
}

#[tauri::command]
pub async fn spreadsheet_agent_stop(batch_id: String) -> Result<(), String> {
    info!("spreadsheet_agent_stop: batch={}", batch_id);
    if let Ok(map) = CANCEL_TOKENS.lock() {
        if let Some(t) = map.get(&batch_id) {
            t.store(true, Ordering::SeqCst);
            return Ok(());
        }
    }
    // Not finding the token is fine — the run probably already completed.
    Ok(())
}

fn emit(app: &AppHandle, ev: AgentEvent) {
    // Log a one-liner per event so we can correlate what Claude actually
    // produces vs. what the frontend renders. Verbose during v1 dev — drop
    // to debug! level once the loop is stable.
    match &ev {
        AgentEvent::Started { tab_id, batch_id } =>
            info!("agent_event: started tab={} batch={}", tab_id, batch_id),
        AgentEvent::TextDelta { delta, batch_id, .. } =>
            info!("agent_event: text_delta batch={} delta={:?}", batch_id, delta),
        AgentEvent::Reasoning { delta, batch_id, .. } =>
            info!("agent_event: reasoning batch={} delta={:?}", batch_id, delta),
        AgentEvent::ToolCall { name, input, batch_id, .. } =>
            info!("agent_event: tool_call batch={} name={} input={}", batch_id, name, input),
        AgentEvent::Done { stop_reason, input_tokens, output_tokens, cache_read_tokens, cache_creation_tokens, batch_id, .. } =>
            info!(
                "agent_event: done batch={} reason={} in={} out={} cache_read={} cache_creation={}",
                batch_id, stop_reason, input_tokens, output_tokens, cache_read_tokens, cache_creation_tokens
            ),
        AgentEvent::Error { message, batch_id, .. } =>
            error!("agent_event: error batch={} msg={}", batch_id, message),
    }
    if let Err(e) = app.emit(EVENT_NAME, ev) {
        error!("Failed to emit spreadsheet event: {}", e);
    }
}

/// Which LLM backend drives this turn.
#[derive(Debug, Clone, Copy)]
enum Provider {
    /// Anthropic — either API key (sk-ant-api03-) or subscription OAuth
    /// (sk-ant-oat01-). claude.rs's `apply_auth_headers` figures out which.
    Claude,
    /// OpenAI Codex (ChatGPT Plus/Pro subscription) — OAuth access token
    /// from `claude::auth::openai_codex_oauth`.
    Codex,
}

/// Resolve which provider to use + the credential string. Picks based on
/// `api_choice` setting; falls back to whichever credential is actually
/// configured if the choice is ambiguous or missing.
async fn resolve_provider_and_credential(app: &AppHandle) -> Result<(Provider, String), String> {
    let api_choice = app
        .db(|db| get_setting(db, "api_choice"))
        .map(|s| s.setting_value)
        .unwrap_or_default();

    match api_choice.as_str() {
        "openai-codex" => {
            let (token, _acct) =
                crate::auth::openai_codex_oauth::get_active_credentials(app)
                    .await
                    .map_err(|e| format!("ChatGPT subscription auth: {}", e))?;
            Ok((Provider::Codex, token))
        }
        "claude-subscription" => {
            let key = app
                .db(|db| get_setting(db, "api_key_claude_oauth"))
                .map(|s| s.setting_value)
                .unwrap_or_default();
            if key.is_empty() {
                Err("Claude OAuth token not configured. Run `claude setup-token` and paste it in Settings.".to_string())
            } else {
                Ok((Provider::Claude, key))
            }
        }
        "claude" => {
            let key = app
                .db(|db| get_setting(db, "api_key_claude"))
                .map(|s| s.setting_value)
                .unwrap_or_default();
            if key.is_empty() {
                Err("Anthropic API key not configured. Add one in Settings.".to_string())
            } else {
                Ok((Provider::Claude, key))
            }
        }
        _ => {
            // No explicit provider set — fall back by inspecting which
            // credential the user actually pasted, preferring OAuth subs.
            let codex_signed_in = crate::auth::openai_codex_oauth::load(app).is_some();
            if codex_signed_in {
                if let Ok((token, _acct)) =
                    crate::auth::openai_codex_oauth::get_active_credentials(app).await
                {
                    return Ok((Provider::Codex, token));
                }
            }
            let oauth = app
                .db(|db| get_setting(db, "api_key_claude_oauth"))
                .map(|s| s.setting_value)
                .unwrap_or_default();
            if !oauth.is_empty() {
                return Ok((Provider::Claude, oauth));
            }
            let api_key = app
                .db(|db| get_setting(db, "api_key_claude"))
                .map(|s| s.setting_value)
                .unwrap_or_default();
            if !api_key.is_empty() {
                return Ok((Provider::Claude, api_key));
            }
            Err("No LLM credential configured. Open Settings and connect either Claude (API key or subscription) or ChatGPT (subscription).".to_string())
        }
    }
}

/// The STABLE workbook context: sheet previews + reference workbooks.
///
/// This is the dominant token cost per request (a 13-sheet bank model
/// produces an ~860 KB / ~444K-token preview). It's shipped as its own
/// content block with a `cache_control` breakpoint so that, as long as its
/// bytes don't change between turns, every turn after the first READS it
/// from cache at ~10% of input price instead of re-WRITING it at 125%.
/// Anything that changes per turn (prior conversation, focus, the request)
/// must stay OUT of this block — one changed byte re-bills the whole prefix.
fn format_context_block(ctx: &WorkbookContext) -> String {
    let mut sheets_block = String::new();
    for s in &ctx.sheets {
        // Used range beats grid capacity ("10000 rows × 200 columns" read as
        // data extent), and NEVER mention 0-indexing — every coordinate the
        // model reads or writes is A1, where row 1 is the first row. The old
        // "(0-indexed)" note here contradicted system-prompt rule 2 and fed
        // the chronic off-by-one formula references.
        match &s.used_range {
            Some(used) => sheets_block.push_str(&format!(
                "## Sheet: {}\nUsed range: {} (A1 notation — row 1 is the first row)\nCurrent cells:\n{}\n\n",
                s.name, used, s.cells_preview
            )),
            None => sheets_block.push_str(&format!(
                "## Sheet: {}\nDimensions: {} rows × {} columns\nCurrent cells:\n{}\n\n",
                s.name, s.row_count, s.column_count, s.cells_preview
            )),
        }
    }

    // Reference workbooks — read-only context the user attached (other xlsx
    // files, e.g. analyst models). Previews are deliberately small; the
    // agent pulls detail on demand with the `read_reference` tool.
    let mut references_block = String::new();
    if !ctx.references.is_empty() {
        references_block.push_str(
            "# Reference workbooks (READ-ONLY)\n\
The user attached the following workbooks as read-only reference context — \
use them for comparisons, benchmarks, and pulling figures. You can NEVER \
write to them; all write tools target the active workbook above only. Each \
sheet below shows a truncated preview — use `read_reference(workbook, sheet, range)` \
to pull any range in full (address the workbook by its label).\n\n",
        );
        for r in &ctx.references {
            references_block.push_str(&format!("## Reference workbook: {} ({})\n", r.label, r.path));
            for s in &r.sheets {
                references_block.push_str(&format!(
                    "### Sheet: {}\nDimensions: {} rows × {} columns\nPreview:\n{}\n\n",
                    s.name, s.row_count, s.column_count, s.cells_preview
                ));
            }
        }
    }

    // Calc-health warning — which cells the live engine can't evaluate the
    // way Excel does. Computed once per loaded workbook in the webview, so
    // the bytes are stable and safe for the cached block.
    let calc_health_block = ctx
        .calc_health
        .as_deref()
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("# Calc health (live evaluation vs file)\n{}\n\n", s))
        .unwrap_or_default();

    // Today's date — the agent has no temporal grounding otherwise, so it
    // pattern-matches on its training cutoff and treats that as "current
    // year," breaking historical-vs-projected boundaries in models. Stable
    // within a day, so it can live in the cached block.
    let today = chrono::Local::now().format("%A, %B %-d, %Y").to_string();

    format!(
        "Today is {}.\n\nWorkbook: {}\n\n\
Preview notation: `A5:Z5 = <formula>, then <second formula>, …  (filled right ×N)` \
means the formula is FILLED across the range — each column's copy shifts relative \
references by one (the second column's exact formula is shown; it is NOT the same \
formula repeated). Sampled values follow after →. Numbers are rounded for display \
and long text is clipped; read the real cells before precision-sensitive work. \
Truncated sheets end with an index of the remaining row labels — those rows exist; \
read them by range when relevant.\n\n{}{}{}",
        today, ctx.path, sheets_block, calc_health_block, references_block
    )
}

/// The PER-TURN tail: prior conversation, workbook delta, focus hint, and
/// the user request. Small and changes every turn — lives in its own
/// uncached block after the cached context block. The delta precedes focus:
/// it's workbook state (supersedes the preview), focus is user intent.
fn format_turn_block(
    prompt: &str,
    delta: Option<&str>,
    focus: Option<&str>,
    prior_batches: Option<&str>,
    prior_sources: Option<&str>,
) -> String {
    let prior = prior_batches
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("# Prior turns in this session\n{}\n\n", s))
        .unwrap_or_default();

    let sources = prior_sources
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("{}\n\n", s))
        .unwrap_or_default();

    let delta = delta
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("{}\n\n", s))
        .unwrap_or_default();

    let focus = focus
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("{}\n", s))
        .unwrap_or_default();

    format!("{}{}{}{}# User request\n{}", prior, sources, delta, focus, prompt)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sheet(name: &str, preview: &str) -> SheetContext {
        SheetContext {
            name: name.to_string(),
            row_count: 100,
            column_count: 26,
            used_range: Some("A1:D10".to_string()),
            cells_preview: preview.to_string(),
        }
    }

    #[test]
    fn context_block_without_references_has_no_reference_block() {
        let ctx = WorkbookContext {
            path: "/tmp/model.xlsx".to_string(),
            sheets: vec![sheet("Model", "A1 = \"Revenue\"")],
            focus: None,
            delta: None,
            calc_health: None,
            references: vec![],
        };
        let out = format_context_block(&ctx);
        assert!(!out.contains("Reference workbooks"));
        assert!(!out.contains("# Calc health"));
        assert!(out.contains("## Sheet: Model"));
        let turn = format_turn_block("add a row", None, None, None, None);
        assert!(turn.contains("# User request\nadd a row"));
    }

    #[test]
    fn context_block_renders_references_after_sheets() {
        let ctx = WorkbookContext {
            path: "/tmp/model.xlsx".to_string(),
            sheets: vec![sheet("Model", "A1 = \"Revenue\"")],
            focus: None,
            delta: None,
            calc_health: None,
            references: vec![ReferenceWorkbookContext {
                path: "/tmp/AAPL-GS.xlsx".to_string(),
                label: "AAPL-GS.xlsx".to_string(),
                sheets: vec![sheet("Estimates", "B2 = 394328")],
            }],
        };
        let out = format_context_block(&ctx);
        assert!(out.contains("# Reference workbooks (READ-ONLY)"));
        assert!(out.contains("## Reference workbook: AAPL-GS.xlsx (/tmp/AAPL-GS.xlsx)"));
        assert!(out.contains("### Sheet: Estimates"));
        assert!(out.contains("B2 = 394328"));
        assert!(out.contains("read_reference"));
        let sheets_pos = out.find("## Sheet: Model").unwrap();
        let refs_pos = out.find("# Reference workbooks").unwrap();
        assert!(sheets_pos < refs_pos);
    }

    #[test]
    fn turn_block_orders_prior_then_sources_then_delta_then_focus_then_request() {
        let out = format_turn_block(
            "compare revenue",
            Some("# Workbook changes since the preview snapshot\nA1 = 2"),
            Some("focus hint"),
            Some("prior stuff"),
            Some("# Previously fetched sources (this session)\n- src1 — 3 page(s) — https://example.com"),
        );
        let prior_pos = out.find("# Prior turns in this session").unwrap();
        let sources_pos = out.find("# Previously fetched sources").unwrap();
        let delta_pos = out.find("# Workbook changes since the preview snapshot").unwrap();
        let focus_pos = out.find("focus hint").unwrap();
        let req_pos = out.find("# User request").unwrap();
        assert!(prior_pos < sources_pos && sources_pos < delta_pos && delta_pos < focus_pos && focus_pos < req_pos);
    }

    /// The whole point of the split: identical workbook context must produce
    /// byte-identical blocks across turns, or the cache prefix breaks.
    #[test]
    fn context_block_is_stable_across_turns() {
        let ctx = WorkbookContext {
            path: "/tmp/model.xlsx".to_string(),
            sheets: vec![sheet("Model", "A1 = \"Revenue\"")],
            focus: None,
            delta: None,
            calc_health: None,
            references: vec![],
        };
        assert_eq!(format_context_block(&ctx), format_context_block(&ctx));
    }

    #[test]
    fn context_block_renders_calc_health_between_sheets_and_references() {
        let ctx = WorkbookContext {
            path: "/tmp/model.xlsx".to_string(),
            sheets: vec![sheet("Model", "A1 = \"Revenue\"")],
            focus: None,
            delta: None,
            calc_health: Some("Live evaluation diverges on 12 of 400 checked cells (Model: 12).".to_string()),
            references: vec![ReferenceWorkbookContext {
                path: "/tmp/AAPL-GS.xlsx".to_string(),
                label: "AAPL-GS.xlsx".to_string(),
                sheets: vec![sheet("Estimates", "B2 = 394328")],
            }],
        };
        let out = format_context_block(&ctx);
        let sheets_pos = out.find("## Sheet: Model").unwrap();
        let health_pos = out.find("# Calc health (live evaluation vs file)").unwrap();
        let refs_pos = out.find("# Reference workbooks").unwrap();
        assert!(sheets_pos < health_pos && health_pos < refs_pos);
        assert!(out.contains("12 of 400"));
    }

    #[test]
    fn workbook_context_deserializes_without_references_field() {
        // Older frontends (or payloads with no attachments) omit the fields.
        let v: WorkbookContext = serde_json::from_value(json!({
            "path": "/tmp/model.xlsx",
            "sheets": [],
        }))
        .unwrap();
        assert!(v.references.is_empty());
        assert!(v.delta.is_none());
    }
}

/// Kicks off a single agent turn. Spawns a background task so the webview
/// gets streaming events without blocking the Tauri command return.
#[tauri::command]
pub async fn spreadsheet_agent_turn(
    app_handle: AppHandle,
    tab_id: String,
    batch_id: String,
    prompt: String,
    workbook_context: WorkbookContext,
    prior_batches_context: Option<String>,
) -> Result<(), String> {
    info!(
        "spreadsheet_agent_turn: tab={} batch={} prompt_len={} prior_ctx_len={}",
        tab_id,
        batch_id,
        prompt.len(),
        prior_batches_context.as_ref().map(|s| s.len()).unwrap_or(0),
    );

    let (provider, api_key) = resolve_provider_and_credential(&app_handle).await?;
    info!("spreadsheet_agent_turn: provider={:?}", provider);

    // Two content blocks: the big workbook context (byte-stable across turns,
    // cache_control breakpoint → cache READ from turn 2 on) and the small
    // per-turn tail. Codex flattens blocks to sequential messages; the
    // cache_control key is Anthropic-only and ignored there.
    let context_block = format_context_block(&workbook_context);
    let prior_sources_block = format_prior_sources_block(&tab_id);
    let turn_block = format_turn_block(
        &prompt,
        workbook_context.delta.as_deref(),
        workbook_context.focus.as_deref(),
        prior_batches_context.as_deref(),
        prior_sources_block.as_deref(),
    );
    // The context block must be byte-identical across turns for the prompt
    // cache to hit. Log its hash so a cache_creation spike in the logs is
    // attributable: same hash + rewrite = TTL/provider issue; changed hash =
    // the capture actually changed.
    info!(
        "agent context block: {}B hash={:016x} (turn tail {}B)",
        context_block.len(),
        xxhash_rust::xxh64::xxh64(context_block.as_bytes(), 0),
        turn_block.len(),
    );
    let initial_messages = vec![StreamMessage {
        role: "user".to_string(),
        content: json!([
            { "type": "text", "text": context_block, "cache_control": { "type": "ephemeral" } },
            { "type": "text", "text": turn_block },
        ]),
    }];
    let tools = agent_tools();
    // Claude's claude.rs injects a hosted web_search server tool; Codex
    // doesn't (ChatGPT-sub rejects it). Teach the agent about web_search
    // in the system prompt only when it can actually call it.
    let has_web_search = matches!(provider, Provider::Claude);
    let system = system_prompt(has_web_search);

    let app_for_task = app_handle.clone();
    let tab_id_owned = tab_id.clone();
    let batch_id_owned = batch_id.clone();

    let cancel_token = register_cancel_token(&batch_id_owned);

    tokio::spawn(async move {
        emit(
            &app_for_task,
            AgentEvent::Started {
                tab_id: tab_id_owned.clone(),
                batch_id: batch_id_owned.clone(),
            },
        );

        run_agent_loop(
            app_for_task,
            tab_id_owned,
            batch_id_owned.clone(),
            provider,
            api_key,
            system,
            tools,
            initial_messages,
            cancel_token,
        )
        .await;

        unregister_cancel_token(&batch_id_owned);
    });

    Ok(())
}

// ============================================================================
// Fetched-source paging + compaction
//
// A `fetch_web` result can be huge (a 10-K is ~hundreds of KB). Left in the
// conversation verbatim it re-bills the cache prefix and dilutes attention for
// the rest of the build, even though the model only needs a handful of figures
// from it. Strategy: the model REVIEWS the full source on the turn it's
// delivered, extracts what it needs, then `keep_pages` marks the few pages
// worth retaining — we evict the rest, collapsing the blob to a small
// index-only stub. Raw pages stay stashed so `read_source` can recover any
// page — within the run (a last resort) AND on later runs in the same
// session, where the stash + the per-session registry (SESSION_SOURCES) let
// the agent re-read a source instead of re-fetching it. An auto-evict
// fallback covers a forgotten keep_pages.
//
// Sources are addressed by a SHORT id ("src1", "src2", …), NOT the 25-char
// `toolu_` tool_use_id: models reliably echo a short token but mangle the long
// opaque id (observed in logs: toolu_ -> toulu_), which silently broke pruning
// and forced wasteful re-fetches. We keep an internal alias->tool_use_id map
// because the compaction still has to match the real id on the tool_result.
// ============================================================================

/// Raw fetched pages keyed by short source id ("src1"). Globally unique via
/// FETCH_SEQ, so no per-session collision. Entries PERSIST after the owning
/// agent run finishes — a later turn in the same session can `read_source`
/// any page without re-fetching (the run's structured history is discarded,
/// but the stash isn't). Bounded by MAX_STASHED_SOURCES (oldest-first
/// eviction via FETCH_ORDER).
static FETCHED_SOURCES: Lazy<Mutex<HashMap<String, Vec<String>>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Insertion order of stashed sources, backing the MAX_STASHED_SOURCES cap.
static FETCH_ORDER: Lazy<Mutex<VecDeque<String>>> =
    Lazy::new(|| Mutex::new(VecDeque::new()));

/// Cap on sources retained across the app's lifetime. Each source is at most
/// MAX_SOURCE_PAGES × SOURCE_PAGE_BYTES ≈ 600 KB, so this bounds the stash
/// at ~20 MB worst-case.
const MAX_STASHED_SOURCES: usize = 32;

/// Stash a fetched source's pages, evicting the oldest sources beyond the cap.
fn stash_source(alias: &str, pages: Vec<String>) {
    if let (Ok(mut store), Ok(mut order)) = (FETCHED_SOURCES.lock(), FETCH_ORDER.lock()) {
        if store.insert(alias.to_string(), pages).is_none() {
            order.push_back(alias.to_string());
        }
        while order.len() > MAX_STASHED_SOURCES {
            if let Some(old) = order.pop_front() {
                store.remove(&old);
            }
        }
    }
}

/// One session's (tab's) cross-run source registry: the URL-dedup map so a
/// later run short-circuits a re-fetch of an already-pulled page, plus
/// display metadata for the "# Previously fetched sources" block injected
/// into later turns' user messages.
#[derive(Default)]
struct SessionSources {
    url_to_alias: HashMap<String, String>,
    metas: Vec<SourceMeta>,
}

struct SourceMeta {
    alias: String,
    urls: Vec<String>,
}

static SESSION_SOURCES: Lazy<Mutex<HashMap<String, SessionSources>>> =
    Lazy::new(|| Mutex::new(HashMap::new()));

/// Render the "# Previously fetched sources" block for a tab: one line per
/// source fetched on an earlier turn whose pages are still stashed. Injected
/// into the (uncached) turn tail so the model reaches for `read_source`
/// instead of re-fetching. None when the session has no live prior sources.
fn format_prior_sources_block(tab_id: &str) -> Option<String> {
    let sessions = SESSION_SOURCES.lock().ok()?;
    let entry = sessions.get(tab_id)?;
    let store = FETCHED_SOURCES.lock().ok()?;
    let lines: Vec<String> = entry
        .metas
        .iter()
        .filter_map(|m| {
            store.get(&m.alias).map(|pages| {
                format!("- {} — {} page(s) — {}", m.alias, pages.len(), m.urls.join(", "))
            })
        })
        .collect();
    if lines.is_empty() {
        return None;
    }
    Some(format!(
        "# Previously fetched sources (this session)\n\
These web pages were fetched on earlier turns and their full page text is still stored. \
Do NOT re-fetch these URLs — call read_source(source_id, [pages]) to read any page again.\n{}",
        lines.join("\n")
    ))
}

/// Monotonic source-id sequence. Short + globally unique so the model can
/// reproduce it verbatim in keep_pages / read_source.
static FETCH_SEQ: AtomicUsize = AtomicUsize::new(1);

/// Target byte size of one source page. Small enough that the auto-built
/// index is specific (the model can tell which page holds the income
/// statement); large enough that a filing doesn't explode into hundreds of
/// pages.
const SOURCE_PAGE_BYTES: usize = 4000;
/// Hard cap on pages per source — bounds a pathologically large document.
const MAX_SOURCE_PAGES: usize = 150;
/// How many turns a fetched source may stay in full before we auto-evict it
/// to an index-only stub. The review turn is fetch_turn+1, so 2 leaves the
/// full text live for review plus one extra working turn.
const AUTO_COMPACT_AFTER_TURNS: usize = 2;

/// Split extracted text into ~SOURCE_PAGE_BYTES chunks on line boundaries
/// (never mid-UTF8-char). A single oversized line is char-split.
fn paginate_source(text: &str) -> Vec<String> {
    let mut pages: Vec<String> = Vec::new();
    let mut cur = String::new();
    for line in text.split_inclusive('\n') {
        if pages.len() >= MAX_SOURCE_PAGES {
            break;
        }
        if line.len() > SOURCE_PAGE_BYTES {
            if !cur.is_empty() {
                pages.push(std::mem::take(&mut cur));
            }
            for ch in line.chars() {
                if cur.len() + ch.len_utf8() > SOURCE_PAGE_BYTES && !cur.is_empty() {
                    pages.push(std::mem::take(&mut cur));
                    if pages.len() >= MAX_SOURCE_PAGES {
                        break;
                    }
                }
                cur.push(ch);
            }
            continue;
        }
        if cur.len() + line.len() > SOURCE_PAGE_BYTES && !cur.is_empty() {
            pages.push(std::mem::take(&mut cur));
        }
        cur.push_str(line);
    }
    if !cur.is_empty() && pages.len() < MAX_SOURCE_PAGES {
        pages.push(cur);
    }
    pages
}

/// One navigational line per page (cleaned snippet of the page head, no model
/// call) so the model knows what each evicted page held.
fn source_index(pages: &[String]) -> String {
    pages
        .iter()
        .enumerate()
        .map(|(i, p)| {
            let snippet: String = p.split_whitespace().collect::<Vec<_>>().join(" ");
            let snippet: String = snippet.chars().take(100).collect();
            format!("PAGE {}: {}", i + 1, snippet)
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Normalized key for a fetch_web call, used to detect re-fetches of a page
/// already pulled this run. `fetch_web` returns the WHOLE page regardless of
/// any `#fragment`, so we strip the fragment before keying — a model that
/// re-fetches `…htm#section` thinking it'll jump to a section would otherwise
/// re-download the identical 450 KB. Multi-URL calls are keyed on the sorted
/// set so order doesn't matter.
fn normalize_fetch_key(input: &Value) -> String {
    let mut urls: Vec<String> = input
        .get("urls")
        .and_then(|v| v.as_array())
        .map(|a| {
            a.iter()
                .filter_map(|v| v.as_str())
                .map(|s| s.split('#').next().unwrap_or(s).trim().to_string())
                .collect()
        })
        .unwrap_or_default();
    urls.sort();
    urls.join("\n")
}

/// Render the given 1-indexed pages with `===== PAGE n =====` markers.
fn render_pages(pages: &[String], which: impl Iterator<Item = usize>) -> String {
    which
        .filter_map(|n| pages.get(n.wrapping_sub(1)).map(|p| format!("===== PAGE {} =====\n{}", n, p)))
        .collect::<Vec<_>>()
        .join("\n\n")
}

/// Replace the fetch_web tool_result identified by `tool_use_id` with a
/// compact stub: the full page index plus only the retained pages' text.
/// `display_id` is the short id shown to the model in the stub. Returns true
/// if a matching block was found.
fn compact_source_in_messages(
    messages: &mut [StreamMessage],
    tool_use_id: &str,
    display_id: &str,
    keep: &HashSet<usize>,
    pages: &[String],
) -> bool {
    let index = source_index(pages);
    let mut kept: Vec<usize> = keep
        .iter()
        .copied()
        .filter(|n| *n >= 1 && *n <= pages.len())
        .collect();
    kept.sort_unstable();
    let retained = if kept.is_empty() {
        "(none — all pages evicted)".to_string()
    } else {
        render_pages(pages, kept.into_iter())
    };
    let stub = format!(
        "source_id: {} — full text evicted from context to keep it small ({} pages total). \
Use read_source(source_id, [pages]) to recover any page (LAST RESORT — you should have kept \
what you needed during review).\n\nINDEX:\n{}\n\nRETAINED PAGES:\n{}",
        display_id,
        pages.len(),
        index,
        retained
    );
    let mut found = false;
    for m in messages.iter_mut() {
        if let Some(arr) = m.content.as_array_mut() {
            for blk in arr.iter_mut() {
                let is_match = blk.get("type").and_then(|v| v.as_str()) == Some("tool_result")
                    && blk.get("tool_use_id").and_then(|v| v.as_str()) == Some(tool_use_id);
                if is_match {
                    if let Some(obj) = blk.as_object_mut() {
                        obj.insert("content".to_string(), Value::String(stub.clone()));
                        found = true;
                    }
                }
            }
        }
    }
    found
}

/// Cap on agent turns within one user request. Anthropic charges per turn so
/// runaway loops are expensive; 45 is enough headroom for a full model build
/// (fetch + data + bulk format + merges + widths + done) while still bounding
/// runaway agents.
const MAX_AGENT_TURNS: usize = 45;

/// How many times to transparently retry a turn whose provider stream broke
/// *before* emitting any content (text, reasoning, or tool calls). Such breaks
/// are transient connection failures (TLS reset, idle-connection close) and a
/// retry hits the cached prefix cheaply. We deliberately do NOT auto-retry once
/// content has streamed to the UI — re-running would duplicate the visible
/// reasoning/prose, so those surface as an error for the user to re-run.
const MAX_TURN_RETRIES: u32 = 2;

/// Per-turn output-token ceiling passed to the providers. Must cover a long
/// extended-thinking block PLUS the largest tool payload we expect (a
/// full-model `set_range`) — thinking is billed as output and, on Opus, the
/// thinking budget is guidance rather than a hard cap. The old 4096 (floored
/// to thinking+8K by the provider-side guard, i.e. 12,192) was exhausted by
/// a single thinking block on a model-build turn: the turn died at
/// stop_reason=max_tokens having emitted no text and no tool call.
const TURN_MAX_TOKENS: usize = 48_000;

/// Tools that don't need a tool_result with evaluated cell values — formats,
/// merges, dimensions all just succeed-or-not. We supply "ok" immediately
/// instead of waiting up to 8s for a frontend round-trip that has nothing to
/// say. Was a major source of wasted turns: each set_format burned a full
/// timeout window before the loop could continue.
fn tool_skips_readback(name: &str) -> bool {
    matches!(
        name,
        "set_format"
            | "set_column_width"
            | "set_row_height"
            | "merge_cells"
            | "unmerge_cells"
            | "set_note"
            | "delete_note"
            | "create_sheet"
            | "delete_sheet"
            | "rename_sheet"
            | "clear_range"
            | "insert_rows"
            | "delete_rows"
            | "insert_columns"
            | "delete_columns"
            | "freeze_panes"
            | "unfreeze_panes"
            | "hide_rows"
            | "show_rows"
            | "hide_columns"
            | "show_columns"
            | "define_name"
    )
}

/// How long to wait for the webview to ship evaluated cell values back
/// after a write tool. Scales with the size of the operation — a 50×9
/// `set_range` needs Univer to apply each cell + recompute formulas + run
/// a React commit before it can serialize values back through Tauri IPC.
/// Base 15s covers tiny writes; we add ~1s per 30 cells touched, capped
/// at 90s so a runaway never hangs the agent loop.
/// Cell count of an A1 range like "B5:D20" (single cell → 1, malformed → 1).
/// Used only for timeout sizing, so parse failures err small and cheap.
fn a1_range_cell_count(range: &str) -> usize {
    fn parse(cell: &str) -> Option<(u64, u64)> {
        let cell = cell.trim();
        let split = cell.find(|c: char| c.is_ascii_digit())?;
        let (letters, digits) = cell.split_at(split);
        let col = letters
            .chars()
            .try_fold(0u64, |acc, c| {
                let d = (c.to_ascii_uppercase() as u64).checked_sub('A' as u64 - 1)?;
                (1..=26).contains(&d).then(|| acc * 26 + d)
            })?;
        let row: u64 = digits.parse().ok()?;
        Some((row, col))
    }
    let mut parts = range.splitn(2, ':');
    let a = parts.next().and_then(parse);
    let b = parts.next().and_then(parse);
    match (a, b) {
        (Some((r1, c1)), Some((r2, c2))) => {
            ((r1.abs_diff(r2) + 1) * (c1.abs_diff(c2) + 1)) as usize
        }
        (Some(_), None) => 1,
        _ => 1,
    }
}

fn readback_timeout_secs(name: &str, input: &Value) -> u64 {
    // Index lookups are served from a memoized in-memory structure on the
    // frontend — near-instant, no Univer apply/recalc involved. A short
    // window keeps a hung webview from stalling the loop for 25s.
    if matches!(name, "describe_workbook" | "find_rows") {
        return 10;
    }
    // A script's write volume is unknown until it runs — it can legally touch
    // 20k cells. The webview enforces its own 5s CPU cap on execution; the
    // long pole is the same Univer apply + recalc + readback as a huge
    // set_range, so give it the ceiling outright.
    if name == "run_script" {
        return 90;
    }
    let cells: usize = match name {
        "set_range" => input
            .get("values")
            .and_then(|v| v.as_array())
            .map(|rows| {
                let row_count = rows.len();
                let col_count = rows
                    .first()
                    .and_then(|r| r.as_array())
                    .map(|c| c.len())
                    .unwrap_or(1);
                row_count * col_count
            })
            .unwrap_or(1),
        "set_cell" => 1,
        // copy_range expands server-side into one mutation per source cell —
        // the webview's apply cost scales with the SOURCE rectangle.
        "copy_range" => input
            .get("source")
            .and_then(|v| v.as_str())
            .map(a1_range_cell_count)
            .unwrap_or(1),
        _ => 1,
    };
    // Base 25s covers the typical small write + the Univer/React commit
    // overhead. Add ~1s per ~20 cells so a 150-cell set_range gets ~32s,
    // a 500-cell write gets ~50s. Capped at 90s so a pathological prompt
    // never wedges the loop.
    let secs = 25u64 + (cells as u64) / 20;
    secs.min(90)
}

/// Per-turn size audit: prints what we're about to ship and (when
/// `GRIDPATH_AGENT_DEBUG=1`) dumps the full `messages` array to a JSONL
/// file under app-data so we can rerun a tokenizer offline.
///
/// Goal is diagnosing cache_read inflation: by logging every turn's
/// payload size we can see whether the prefix is growing as expected
/// (linear in tool history) or blowing up unexpectedly.
fn log_turn_request(
    app: &AppHandle,
    batch_id: &str,
    turn: usize,
    system: &str,
    tools: &[Value],
    messages: &[StreamMessage],
) {
    let sys_bytes = system.len();
    let tools_bytes: usize = tools.iter().map(|t| t.to_string().len()).sum();
    let mut total_msg_bytes: usize = 0;
    let mut per_msg: Vec<String> = Vec::with_capacity(messages.len());
    for (i, m) in messages.iter().enumerate() {
        let raw = m.content.to_string();
        total_msg_bytes += raw.len();
        // Summarize content: count text / tool_use / tool_result blocks
        // and the largest tool_use input or tool_result content size, so
        // we can spot which historical entry is dominating.
        let mut text_b = 0usize;
        let mut tu_n = 0usize;
        let mut tu_max = 0usize;
        let mut tu_name = String::new();
        let mut tr_n = 0usize;
        let mut tr_max = 0usize;
        if let Some(arr) = m.content.as_array() {
            for blk in arr {
                let ty = blk.get("type").and_then(|v| v.as_str()).unwrap_or("");
                match ty {
                    "text" => {
                        text_b += blk.get("text").and_then(|v| v.as_str()).map(|s| s.len()).unwrap_or(0);
                    }
                    "tool_use" => {
                        tu_n += 1;
                        let isize_ = blk.get("input").map(|v| v.to_string().len()).unwrap_or(0);
                        if isize_ > tu_max {
                            tu_max = isize_;
                            tu_name = blk.get("name").and_then(|v| v.as_str()).unwrap_or("").to_string();
                        }
                    }
                    "tool_result" => {
                        tr_n += 1;
                        let csize = blk.get("content").map(|v| match v.as_str() {
                            Some(s) => s.len(),
                            None => v.to_string().len(),
                        }).unwrap_or(0);
                        if csize > tr_max { tr_max = csize; }
                    }
                    _ => {}
                }
            }
        }
        per_msg.push(format!(
            "  msg[{}] {:<9} | bytes={} text={} tool_use={}(max={}B,{}) tool_result={}(max={}B)",
            i, m.role, raw.len(), text_b, tu_n, tu_max, tu_name, tr_n, tr_max
        ));
    }
    let total = sys_bytes + tools_bytes + total_msg_bytes;
    info!(
        "agent_loop: turn={} req_size sys={}B tools={}B msgs={}B msg_count={} total={}KB",
        turn, sys_bytes, tools_bytes, total_msg_bytes, messages.len(), total / 1024
    );
    for line in &per_msg {
        info!("{}", line);
    }

    if std::env::var("GRIDPATH_AGENT_DEBUG").ok().as_deref() == Some("1") {
        if let Ok(dir) = app.path().app_data_dir() {
            let dump_dir = dir.join("agent_debug");
            if std::fs::create_dir_all(&dump_dir).is_ok() {
                let path = dump_dir.join(format!("{}.jsonl", batch_id));
                let record = json!({
                    "turn": turn,
                    "sys_bytes": sys_bytes,
                    "tools_bytes": tools_bytes,
                    "msgs_bytes": total_msg_bytes,
                    "messages": messages.iter().map(|m| json!({
                        "role": m.role,
                        "content": m.content,
                    })).collect::<Vec<_>>(),
                });
                if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(&path) {
                    use std::io::Write;
                    let _ = writeln!(f, "{}", record);
                }
            }
        }
    }
}

/// Multi-turn driver: keep calling Claude until it emits `done`, returns
/// `end_turn`, or we hit the turn cap. Between turns we synthesize
/// `tool_result` messages so Claude can keep working — for spreadsheet
/// writes the result is always `"ok"` (we accept everything optimistically
/// and the user reviews the final batch).
async fn run_agent_loop(
    app: AppHandle,
    tab_id: String,
    batch_id: String,
    provider: Provider,
    api_key: String,
    system: String,
    tools: Vec<Value>,
    mut messages: Vec<StreamMessage>,
    cancel_token: Arc<AtomicBool>,
) {

    let mut total_input_tokens: u32 = 0;
    let mut total_output_tokens: u32 = 0;
    let mut total_cache_read_tokens: u32 = 0;
    let mut total_cache_creation_tokens: u32 = 0;
    let mut last_stop_reason = String::new();

    // Fetched-source bookkeeping for this run, all keyed by the short source
    // id ("src1"). `sources_seen` maps id -> turn fetched (drives auto-evict);
    // `compacted_sources` records the ones already collapsed; the alias maps
    // bridge the short id the model uses and the real tool_use_id the
    // compaction must match on the tool_result block.
    let mut sources_seen: HashMap<String, usize> = HashMap::new();
    let mut compacted_sources: HashSet<String> = HashSet::new();
    let mut alias_to_toolid: HashMap<String, String> = HashMap::new();
    let mut toolid_to_alias: HashMap<String, String> = HashMap::new();
    // Dedup: normalized-URL -> alias of the source that already fetched it,
    // and tool_use_id -> that alias for fetches we short-circuit instead of
    // re-downloading. Seeded from the session registry so a URL fetched on an
    // EARLIER run dedups too (the model gets pointed at read_source instead
    // of re-downloading); entries whose pages were evicted from the stash are
    // dropped from the seed — those genuinely need a re-fetch.
    let mut url_to_alias: HashMap<String, String> = SESSION_SOURCES
        .lock()
        .ok()
        .and_then(|s| s.get(&tab_id).map(|ss| ss.url_to_alias.clone()))
        .unwrap_or_default();
    if !url_to_alias.is_empty() {
        if let Ok(store) = FETCHED_SOURCES.lock() {
            url_to_alias.retain(|_, alias| store.contains_key(alias));
        }
    }
    let mut toolid_to_dup: HashMap<String, String> = HashMap::new();
    // One free retry for a turn that dies at max_tokens with nothing usable
    // (see the check below the turn dispatch).
    let mut max_tokens_retries: u32 = 0;

    for turn in 0..MAX_AGENT_TURNS {
        if cancel_token.load(Ordering::SeqCst) {
            info!("agent_loop: cancelled before turn {}", turn);
            emit(
                &app,
                AgentEvent::Done {
                    tab_id: tab_id.clone(),
                    batch_id: batch_id.clone(),
                    stop_reason: "stopped".to_string(),
                    input_tokens: total_input_tokens,
                    output_tokens: total_output_tokens,
                    cache_read_tokens: total_cache_read_tokens,
                    cache_creation_tokens: total_cache_creation_tokens,
                },
            );
            return;
        }

        // Auto-evict fallback: any fetched source that's lingered in full for
        // AUTO_COMPACT_AFTER_TURNS without an explicit keep_pages gets
        // collapsed to an index-only stub now. Guards against the model
        // forgetting to prune (or mangling the source id).
        let to_evict: Vec<String> = sources_seen
            .iter()
            .filter(|(id, &t)| {
                !compacted_sources.contains(*id) && turn.saturating_sub(t) >= AUTO_COMPACT_AFTER_TURNS
            })
            .map(|(id, _)| id.clone())
            .collect();
        for alias in to_evict {
            let tool_use_id = alias_to_toolid.get(&alias).cloned();
            let pages = FETCHED_SOURCES.lock().ok().and_then(|s| s.get(&alias).cloned());
            if let (Some(tuid), Some(pages)) = (tool_use_id, pages) {
                if compact_source_in_messages(&mut messages, &tuid, &alias, &HashSet::new(), &pages) {
                    compacted_sources.insert(alias.clone());
                    info!(
                        "agent_loop: auto-evicted source {} (un-pruned after {} turns)",
                        alias, AUTO_COMPACT_AFTER_TURNS
                    );
                }
            }
        }

        // Forced prune: if any fetched source hasn't been pruned yet, the
        // model MUST call `keep_pages` before it can do anything else — that's
        // what bounds context the turn right after a fetch instead of letting
        // a 450 KB blob ride along for the rest of the build. Enforcement is
        // at RESULT time, not request time: the tools array is never touched
        // (tool definitions are the first bytes of the prompt-cache prefix, so
        // swapping in a keep_pages-only array invalidated the ENTIRE cached
        // conversation — context block + filing — as an orphaned ~250K-token
        // write per source, keyed under a tools prefix no later turn could
        // read). Instead, any non-keep_pages tool call this turn is refused
        // with an error tool_result (never emitted to the webview, so nothing
        // executes) and the model re-issues it after pruning, with the whole
        // prefix still cached. `tool_choice` forcing isn't an option here —
        // extended thinking is incompatible with it on Anthropic — but the
        // refusal loop gives the same hard guarantee, and the auto-evict above
        // still bounds a model that never complies.
        let awaiting_prune: Vec<String> = sources_seen
            .keys()
            .filter(|a| !compacted_sources.contains(*a))
            .cloned()
            .collect();
        let forcing_prune = !awaiting_prune.is_empty();
        if forcing_prune {
            info!(
                "agent_loop: turn {} demanding keep_pages for un-pruned source(s): {:?} (other tools refused)",
                turn, awaiting_prune
            );
        }

        // Per-turn captures (shared with the streaming callback).
        // block_texts: index -> accumulated text  (for text blocks only)
        // assistant_blocks_buf: ordered (index, json-block) so we can stitch
        //   the assistant message back together in the order Claude emitted.
        let block_texts: Arc<Mutex<HashMap<u32, String>>> = Arc::new(Mutex::new(HashMap::new()));
        let assistant_blocks: Arc<Mutex<Vec<(u32, Value)>>> = Arc::new(Mutex::new(Vec::new()));
        let tool_uses_this_turn: Arc<Mutex<Vec<(String, String, Value)>>> = Arc::new(Mutex::new(Vec::new()));
        // Receivers parallel to tool_uses_this_turn — one per non-`done` tool.
        // Registered BEFORE the tool_call event is emitted so we never miss
        // a fast result that arrives during the SSE loop itself.
        let tool_result_rxs: Arc<Mutex<HashMap<String, ToolWaiters>>> =
            Arc::new(Mutex::new(HashMap::new()));
        let turn_stop_reason: Arc<Mutex<String>> = Arc::new(Mutex::new(String::new()));
        let turn_in: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_out: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_cache_read: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_cache_creation: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        // Set true once anything user-visible (text, reasoning, or a tool call)
        // has streamed this turn. Gates the transparent-retry path: we only
        // retry a broken stream if nothing was emitted yet (see MAX_TURN_RETRIES).
        let emitted_content = Arc::new(AtomicBool::new(false));

        let mut attempt: u32 = 0;
        let result = loop {
        let app_cb = app.clone();
        let tab_cb = tab_id.clone();
        let batch_cb = batch_id.clone();
        let block_texts_cb = block_texts.clone();
        let assistant_blocks_cb = assistant_blocks.clone();
        let tool_uses_cb = tool_uses_this_turn.clone();
        let rxs_cb = tool_result_rxs.clone();
        let stop_reason_cb = turn_stop_reason.clone();
        let in_cb = turn_in.clone();
        let out_cb = turn_out.clone();
        let cache_read_cb = turn_cache_read.clone();
        let cache_creation_cb = turn_cache_creation.clone();
        let emitted_cb = emitted_content.clone();

        let on_event = move |event: StreamEvent| match event {
                StreamEvent::TextDelta { index, delta } => {
                    emitted_cb.store(true, Ordering::SeqCst);
                    if let Ok(mut bt) = block_texts_cb.lock() {
                        bt.entry(index).or_default().push_str(&delta);
                    }
                    emit(
                        &app_cb,
                        AgentEvent::TextDelta {
                            tab_id: tab_cb.clone(),
                            batch_id: batch_cb.clone(),
                            delta,
                        },
                    );
                }
                StreamEvent::ReasoningDelta { delta } => {
                    emitted_cb.store(true, Ordering::SeqCst);
                    // Stream the model's plan to the UI live. Not accumulated
                    // into the assistant message — the complete signature-bearing
                    // ThinkingBlock (Claude) is stitched back separately; Codex
                    // reasoning is summary-only and not replayed.
                    emit(
                        &app_cb,
                        AgentEvent::Reasoning {
                            tab_id: tab_cb.clone(),
                            batch_id: batch_cb.clone(),
                            delta,
                        },
                    );
                }
                StreamEvent::ToolCall {
                    index,
                    tool_use_id,
                    name,
                    input,
                } => {
                    emitted_cb.store(true, Ordering::SeqCst);
                    // On a forced-prune turn, everything except keep_pages is
                    // refused: the block still goes into the assistant message
                    // (Anthropic verifies the replayed turn matches what it
                    // emitted), but it's never sent to the webview — nothing
                    // executes — and result collection answers it with an
                    // error tool_result instead.
                    let blocked = forcing_prune && name != "keep_pages";
                    if let Ok(mut tools_l) = tool_uses_cb.lock() {
                        tools_l.push((tool_use_id.clone(), name.clone(), input.clone()));
                    }
                    if let Ok(mut blocks) = assistant_blocks_cb.lock() {
                        blocks.push((
                            index,
                            json!({
                                "type": "tool_use",
                                "id": tool_use_id,
                                "name": name,
                                "input": input,
                            }),
                        ));
                    }
                    if blocked {
                        info!(
                            "agent_loop: refused tool {} ({}) — keep_pages pending",
                            name, tool_use_id
                        );
                        return;
                    }
                    // Register a frontend-result receiver for tools that need
                    // evaluated-value read-back (set_cell, set_range) AND for
                    // `done` (layout validation gate). Format / dimension /
                    // merge tools succeed-or-not — no point waiting for the
                    // webview to confirm, we just supply "ok" immediately
                    // during result collection. `fetch_web` / keep_pages /
                    // read_source are handled entirely in Rust.
                    if name != "fetch_web"
                        && name != "keep_pages"
                        && name != "read_source"
                        && name != "edgar_lookup"
                        && (name == "done" || !tool_skips_readback(&name))
                    {
                        let rx = register_tool_result(&tool_use_id);
                        if let Ok(mut map) = rxs_cb.lock() {
                            map.insert(tool_use_id.clone(), rx);
                        }
                    }
                    emit(
                        &app_cb,
                        AgentEvent::ToolCall {
                            tab_id: tab_cb.clone(),
                            batch_id: batch_cb.clone(),
                            tool_use_id,
                            name,
                            input,
                        },
                    );
                }
                StreamEvent::MessageStop {
                    stop_reason,
                    input_tokens,
                    output_tokens,
                    cache_read_tokens,
                    cache_creation_tokens,
                } => {
                    if let Ok(mut s) = stop_reason_cb.lock() { *s = stop_reason; }
                    if let Ok(mut i) = in_cb.lock() { *i = input_tokens; }
                    if let Ok(mut o) = out_cb.lock() { *o = output_tokens; }
                    if let Ok(mut c) = cache_read_cb.lock() { *c = cache_read_tokens; }
                    if let Ok(mut c) = cache_creation_cb.lock() { *c = cache_creation_tokens; }
                }
                StreamEvent::ThinkingBlock { index, thinking, signature } => {
                    // Stitch the reasoning block back into the assistant
                    // message at its original index (it sorts ahead of the
                    // tool_use blocks). Anthropic verifies the signature and
                    // rejects the next tool-use turn if we drop it.
                    info!(
                        "agent_loop: thinking block — {} chars (sig {})",
                        thinking.len(),
                        if signature.is_empty() { "MISSING" } else { "ok" }
                    );
                    if let Ok(mut blocks) = assistant_blocks_cb.lock() {
                        blocks.push((
                            index,
                            json!({
                                "type": "thinking",
                                "thinking": thinking,
                                "signature": signature,
                            }),
                        ));
                    }
                }
                StreamEvent::RedactedThinking { index, data } => {
                    if let Ok(mut blocks) = assistant_blocks_cb.lock() {
                        blocks.push((
                            index,
                            json!({
                                "type": "redacted_thinking",
                                "data": data,
                            }),
                        ));
                    }
                }
                StreamEvent::ServerToolUse { index, block, .. } => {
                    // Hosted tool call (e.g. Anthropic's built-in web_search,
                    // which the system prompt tells the model to use for
                    // external data). Must be stitched back into the
                    // assistant message at its original index, same as
                    // ThinkingBlock/ToolCall — dropping it here corrupted the
                    // replayed message and got the next turn rejected with a
                    // 400 ("thinking blocks ... cannot be modified"), because
                    // Anthropic checks the whole block sequence, not just the
                    // thinking block itself.
                    if let Ok(mut blocks) = assistant_blocks_cb.lock() {
                        blocks.push((index, block));
                    }
                }
                StreamEvent::ServerToolResult { index, block, .. } => {
                    if let Ok(mut blocks) = assistant_blocks_cb.lock() {
                        blocks.push((index, block));
                    }
                }
                StreamEvent::MessageStart | StreamEvent::BlockStart { .. } => {}
            };

        // Pre-flight size audit — see log_turn_request for what's printed
        // and how to enable the JSONL dump.
        log_turn_request(&app, &batch_id, turn, &system, &tools, &messages);

        // Dispatch to the right provider. Each one consumes the same
        // `on_event` callback and produces the same StreamEvent stream
        // so the rest of the agent loop is provider-agnostic.
        let attempt_result = match provider {
            Provider::Claude => {
                stream_claude_with_tools(
                    &api_key,
                    &system,
                    messages.clone(),
                    tools.clone(),
                    TURN_MAX_TOKENS,
                    on_event,
                ).await
            }
            Provider::Codex => {
                stream_codex_with_tools(
                    &api_key,
                    &system,
                    messages.clone(),
                    tools.clone(),
                    TURN_MAX_TOKENS,
                    // Use tab_id as the OpenAI prompt_cache_key — stable
                    // across all turns of a single spreadsheet session.
                    &tab_id,
                    on_event,
                ).await
            }
        };

        match attempt_result {
            Ok(()) => break Ok::<(), String>(()),
            Err(e) => {
                // Retry only transient breaks that happened before anything
                // streamed to the UI — re-running after content is visible
                // would duplicate the reasoning/prose. A cancel mid-stream
                // also short-circuits the retry.
                let nothing_emitted = !emitted_content.load(Ordering::SeqCst);
                let cancelled = cancel_token.load(Ordering::SeqCst);
                if attempt < MAX_TURN_RETRIES && nothing_emitted && !cancelled {
                    attempt += 1;
                    warn!(
                        "agent_loop: turn {} stream failed before any output (attempt {}/{}), retrying: {}",
                        turn, attempt, MAX_TURN_RETRIES, e
                    );
                    // Reset the per-turn accumulators before the next attempt.
                    // Safe because nothing was emitted, so none of these hold
                    // partial state the downstream code would double-count.
                    if let Ok(mut m) = block_texts.lock() { m.clear(); }
                    if let Ok(mut m) = assistant_blocks.lock() { m.clear(); }
                    if let Ok(mut m) = tool_uses_this_turn.lock() { m.clear(); }
                    if let Ok(mut m) = tool_result_rxs.lock() { m.clear(); }
                    if let Ok(mut s) = turn_stop_reason.lock() { s.clear(); }
                    if let Ok(mut v) = turn_in.lock() { *v = 0; }
                    if let Ok(mut v) = turn_out.lock() { *v = 0; }
                    if let Ok(mut v) = turn_cache_read.lock() { *v = 0; }
                    if let Ok(mut v) = turn_cache_creation.lock() { *v = 0; }
                    // Brief linear backoff so we don't hammer a flapping
                    // connection; the cached prefix makes the retry cheap.
                    tokio::time::sleep(Duration::from_millis(400 * attempt as u64)).await;
                    continue;
                }
                break Err(e);
            }
        }
        };

        if let Err(e) = result {
            emit(
                &app,
                AgentEvent::Error {
                    tab_id: tab_id.clone(),
                    batch_id: batch_id.clone(),
                    message: e,
                },
            );
            return;
        }

        let stop_reason = turn_stop_reason.lock().map(|s| s.clone()).unwrap_or_default();
        let turn_in_v = *turn_in.lock().map(|g| *g).as_ref().unwrap_or(&0);
        let turn_out_v = *turn_out.lock().map(|g| *g).as_ref().unwrap_or(&0);
        let turn_cr_v = *turn_cache_read.lock().map(|g| *g).as_ref().unwrap_or(&0);
        let turn_cc_v = *turn_cache_creation.lock().map(|g| *g).as_ref().unwrap_or(&0);
        total_input_tokens += turn_in_v;
        total_output_tokens += turn_out_v;
        total_cache_read_tokens += turn_cr_v;
        total_cache_creation_tokens += turn_cc_v;
        info!(
            "agent_loop: turn={} tokens in={} out={} cache_read={} cache_creation={}",
            turn, turn_in_v, turn_out_v, turn_cr_v, turn_cc_v
        );
        last_stop_reason = stop_reason.clone();

        let tools_called: Vec<(String, String, Value)> =
            tool_uses_this_turn.lock().map(|g| g.clone()).unwrap_or_default();

        let done_called = tools_called.iter().any(|(_, name, _)| name == "done");

        // Termination conditions:
        //  - Claude explicitly called `done` → we still need to collect its
        //    tool_result (layout validation). Don't break yet — fall through
        //    to result collection, then decide based on validation.
        //  - stop_reason == "end_turn" with no pending tool calls → finish.
        //  - No tools at all this turn → can't continue, treat as finished.
        let no_tools_this_turn = tools_called.is_empty();

        // A max_tokens death with no tool calls produced nothing to build on
        // — typically the whole budget went to one long thinking block on a
        // model-build turn. The truncated output can't be replayed (a cut-off
        // thinking block carries no signature), so discard it and retry the
        // turn once; `messages` hasn't been appended to yet, so this is a
        // clean re-request. A second death becomes an explicit error — the
        // old behavior fell through to the soft-stop and ended the batch
        // silently, which read as a hang in the UI.
        if stop_reason == "max_tokens" && !done_called && no_tools_this_turn {
            if max_tokens_retries == 0 {
                max_tokens_retries += 1;
                warn!(
                    "agent_loop: turn {} hit max_tokens ({} out) with no tool calls — discarding partial output and retrying",
                    turn, turn_out_v
                );
                emit(
                    &app,
                    AgentEvent::Reasoning {
                        tab_id: tab_id.clone(),
                        batch_id: batch_id.clone(),
                        delta: "\n[Ran out of output budget mid-response — retrying the turn.]\n".to_string(),
                    },
                );
                continue;
            }
            emit(
                &app,
                AgentEvent::Error {
                    tab_id: tab_id.clone(),
                    batch_id: batch_id.clone(),
                    message: format!(
                        "The model ran out of output budget (max_tokens, {} tokens) mid-response twice in a row without producing an edit. Re-run the prompt, or break the request into smaller steps.",
                        turn_out_v
                    ),
                },
            );
            return;
        }
        if !no_tools_this_turn {
            // Productive turn — replenish the max_tokens retry allowance so
            // an isolated death later in a long session still gets one.
            max_tokens_retries = 0;
        }

        // Soft-stop (no tools / end_turn without done): finish immediately.
        // Hard-stop via `done` waits for validation below.
        let soft_stop = !done_called && (stop_reason == "end_turn" || no_tools_this_turn);

        if soft_stop {
            info!(
                "agent_loop: stopping after turn {} (done={}, stop_reason={}, tools_this_turn={})",
                turn, done_called, stop_reason, tools_called.len()
            );
            break;
        }

        // Otherwise stop_reason was "tool_use" (or similar), OR the model
        // called `done` and we need to collect its validation result before
        // deciding whether to actually terminate. Build the assistant
        // message + a synthetic tool_result for each tool_use so Claude can
        // take the next turn (or we break after a successful done).
        let mut assistant_content: Vec<Value> = Vec::new();
        // Interleave text blocks and tool_use blocks in their original order.
        // We have text per index from block_texts; tool_use blocks are stored
        // with their indices in assistant_blocks. Merge by index.
        let texts: HashMap<u32, String> = block_texts
            .lock()
            .map(|g| g.clone())
            .unwrap_or_default();
        let mut blocks: Vec<(u32, Value)> = assistant_blocks
            .lock()
            .map(|g| g.clone())
            .unwrap_or_default();
        for (idx, text) in &texts {
            if !text.is_empty() {
                blocks.push((*idx, json!({ "type": "text", "text": text })));
            }
        }
        blocks.sort_by_key(|(idx, _)| *idx);
        for (_, b) in blocks {
            assistant_content.push(b);
        }

        messages.push(StreamMessage {
            role: "assistant".to_string(),
            content: Value::Array(assistant_content),
        });

        // Pull receivers we registered during the streaming callback.
        let mut rxs: HashMap<String, ToolWaiters> =
            tool_result_rxs.lock().map(|mut g| std::mem::take(&mut *g)).unwrap_or_default();

        // Resolve every tool's result CONCURRENTLY. The model now routinely
        // emits a whole batch of tool calls in one turn (rule 12a), so a
        // sequential await would stack each tool's timeout window — one
        // stalled set_range would delay the next tool's wait from even
        // starting. With join_all the total wait is the SLOWEST single tool,
        // not the sum. (The webview still applies the mutations through its
        // own per-tab serial queue; we're only parallelizing the *waiting*,
        // not the application.)
        //
        // Pre-extract each readback receiver in order so the async blocks
        // below don't all need a mutable borrow of `rxs`. `done` NOW produces
        // a tool_result (layout validation); everything else maps to exactly
        // one, tagged with its original position so we can restore order
        // after the join.
        let jobs: Vec<(usize, &String, &String, &Value, Option<ToolWaiters>)> =
            tools_called
                .iter()
                .enumerate()
                .map(|(i, (id, name, input))| {
                    let rust_handled = name == "fetch_web"
                        || name == "keep_pages"
                        || name == "read_source"
                        || name == "edgar_lookup";
                    let needs_rx = name == "done"
                        || (!rust_handled && !tool_skips_readback(name));
                    let rx = if needs_rx { rxs.remove(id) } else { None };
                    (i, id, name, input, rx)
                })
                .collect();

        // Assign a short, stable source id to each new fetch this turn BEFORE
        // spawning the futures, so the fetch can embed it. Models reliably
        // reproduce "src7" but mangle the long toolu_ id (see header note).
        // A fetch whose (fragment-stripped) URL was already pulled this run is
        // marked a dup — the future short-circuits instead of re-downloading.
        // On a forced-prune turn every fetch_web is refused, so none of them
        // get an alias or a stash entry.
        for (id, name, input) in &tools_called {
            if !forcing_prune
                && name == "fetch_web"
                && !toolid_to_alias.contains_key(id)
                && !toolid_to_dup.contains_key(id)
            {
                let key = normalize_fetch_key(input);
                if let Some(existing) = url_to_alias.get(&key) {
                    toolid_to_dup.insert(id.clone(), existing.clone());
                } else {
                    let alias = format!("src{}", FETCH_SEQ.fetch_add(1, Ordering::SeqCst));
                    url_to_alias.insert(key, alias.clone());
                    alias_to_toolid.insert(alias.clone(), id.clone());
                    toolid_to_alias.insert(id.clone(), alias);
                }
            }
        }

        let batch_id_ref = batch_id.as_str();
        let toolid_to_alias_ref = &toolid_to_alias;
        let toolid_to_dup_ref = &toolid_to_dup;
        let awaiting_prune_ref = &awaiting_prune;
        let cancel_ref: &AtomicBool = &cancel_token;
        let result_futs = jobs.into_iter().map(|(i, id, name, input, rx)| async move {
            let content = if forcing_prune && name != "keep_pages" {
                // Refused during a forced-prune turn: nothing was executed
                // (the call never reached the webview). `ok:false` so a
                // blocked `done` can't fail-open through the validation
                // parser and end the batch mid-prune.
                json!({
                    "ok": false,
                    "error": "blocked_pending_prune",
                    "note": format!(
                        "Refused — you must first call keep_pages(source_id, [pages], extracted_notes) for pending source(s): {}. This call was NOT executed; re-issue it after pruning.",
                        awaiting_prune_ref.join(", ")
                    ),
                })
                .to_string()
            } else if name == "fetch_web" {
                // Dedup: this exact page (fragment-stripped) was already pulled
                // this run — short-circuit instead of re-downloading 450 KB and
                // re-flooding context. Point the model back at the existing
                // source rather than handing it a second copy.
                if let Some(existing) = toolid_to_dup_ref.get(id) {
                    let n = FETCHED_SOURCES
                        .lock()
                        .ok()
                        .and_then(|s| s.get(existing).map(|p| p.len()))
                        .unwrap_or(0);
                    info!("fetch_web: deduped — same URL already fetched as {}", existing);
                    json!({
                        "source_id": existing,
                        "deduped": true,
                        "pages": n,
                        "note": format!("You already fetched this exact page as {} ({} pages). fetch_web returns the WHOLE page regardless of any #fragment, so you already have every section of it. Review the pages you already pulled, or call read_source({}, [pages]) to re-surface specific ones. Do NOT re-fetch.", existing, n, existing),
                    })
                    .to_string()
                } else {
                // `fetch_web` is handled entirely in Rust. The result is
                // paginated and stashed so it can be evicted from context
                // after the model's review turn (keep_pages / auto-compact)
                // while staying recoverable via read_source.
                let urls: Vec<String> = input
                    .get("urls")
                    .and_then(|v| v.as_array())
                    .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                    .unwrap_or_default();
                if urls.is_empty() {
                    "{\"error\": \"no urls provided\"}".to_string()
                } else {
                    let alias = toolid_to_alias_ref.get(id).cloned().unwrap_or_else(|| id.clone());
                    let capped: Vec<String> = urls.into_iter().take(5).collect();
                    info!("fetch_web: fetching {} url(s) for batch {}", capped.len(), batch_id_ref);
                    let (text, succeeded) =
                        crate::engine::web_fetcher::fetch_and_extract_pages(capped.clone()).await;
                    let pages = paginate_source(&text);
                    let n = pages.len();
                    stash_source(&alias, pages.clone());
                    info!("fetch_web: source {} → {} pages ({} bytes)", alias, n, text.len());
                    json!({
                        "source_id": alias,
                        "fetched": succeeded.len(),
                        "requested": capped.len(),
                        "pages": n,
                        "instructions": "Review ALL pages NOW and identify every figure you'll need. On your VERY NEXT turn you MUST call keep_pages(source_id, [pages to retain in context], extracted_notes=\"...actual figures with units and years...\") FIRST — every other tool call will be refused with an error until you do. The `extracted_notes` field is your durable scratchpad: it lives in the assistant message forever even after page text is evicted. Be specific (\"FY2024A revenue $97.69B; gross margin 19.0%\") not vague (\"revenue figures\"). Keep pages worth re-reading; pass [] only if extracted_notes already captures everything. read_source can recover an evicted page but is a LAST RESORT.",
                        "content": render_pages(&pages, 1..=n),
                    })
                    .to_string()
                }
                }
            } else if name == "edgar_lookup" {
                // Handled entirely in Rust: resolve ticker/company → CIK via
                // SEC's public map, list recent filings, return exact primary
                // document URLs for fetch_web. Kills guessed-URL 404s and
                // aggregator-basis sourcing (benchmark round 3 root cause).
                let company = input.get("company").and_then(|v| v.as_str()).unwrap_or("").to_string();
                let forms: Vec<String> = input
                    .get("forms")
                    .and_then(|v| v.as_array())
                    .map(|a| a.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                    .unwrap_or_default();
                let count = input.get("count").and_then(|v| v.as_u64()).unwrap_or(6).clamp(1, 12) as usize;
                if company.is_empty() {
                    json!({ "error": "edgar_lookup: 'company' is required (ticker or company name)" }).to_string()
                } else {
                    info!("edgar_lookup: {:?} forms={:?} count={}", company, forms, count);
                    crate::engine::edgar::lookup(&company, &forms, count).await.to_string()
                }
            } else if name == "read_source" {
                // LAST-RESORT recovery: re-surface specific pages of an
                // already-reviewed source from the stash.
                let sid = input.get("source_id").and_then(|v| v.as_str()).unwrap_or("");
                let want: Vec<usize> = input
                    .get("pages")
                    .and_then(|v| v.as_array())
                    .map(|a| a.iter().filter_map(|v| v.as_u64().map(|x| x as usize)).collect())
                    .unwrap_or_default();
                let pages = FETCHED_SOURCES.lock().ok().and_then(|s| s.get(sid).cloned());
                match pages {
                    Some(pages) => {
                        let body = render_pages(&pages, want.into_iter());
                        if body.is_empty() {
                            json!({ "error": "no matching pages", "source_id": sid, "available_pages": pages.len() }).to_string()
                        } else {
                            json!({ "source_id": sid, "content": body }).to_string()
                        }
                    }
                    None => {
                        warn!("read_source: unknown source_id '{}'", sid);
                        json!({ "error": "unknown source_id (it may have expired)", "source_id": sid }).to_string()
                    }
                }
            } else if name == "keep_pages" {
                // The eviction side-effect on `messages` happens after the
                // join (it needs &mut messages); here we just echo back what
                // the model committed to durable memory. Echoing the notes
                // anchors them in the user-side tool_result too — they're
                // already in the assistant tool_use.input forever, and the
                // duplicate appearance gives every model (including ones whose
                // reasoning isn't replayed) a strong attention signal that
                // these figures are the post-eviction ground truth.
                let sid = input.get("source_id").and_then(|v| v.as_str()).unwrap_or("");
                let kept: Vec<i64> = input
                    .get("pages")
                    .and_then(|v| v.as_array())
                    .map(|a| a.iter().filter_map(|v| v.as_i64()).collect())
                    .unwrap_or_default();
                let notes = input
                    .get("extracted_notes")
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                json!({
                    "source_id": sid,
                    "kept_pages": kept,
                    "extracted_notes": notes,
                    "note": "Pages outside `kept_pages` are now evicted from context. Treat `extracted_notes` as the durable record of figures from this source — page text may not be available next turn.",
                })
                .to_string()
            } else if tool_skips_readback(name) {
                // Format/dimension/merge tools never had a receiver registered
                // (see ToolCall handler above) — they succeed-or-not, "ok".
                "ok".to_string()
            } else if let Some(waiters) = rx {
                // Data-writing tools (set_cell, set_range) and `done` (layout
                // validation) registered a receiver; wait for the frontend.
                // For writes the wait is proportional to operation size; for
                // `done` we give a fixed window for the formula engine to
                // settle and the validation scan to complete. Either way it
                // budgets EXECUTION only — time spent queued behind other
                // tools of the same batch is billed separately (see
                // await_tool_result).
                let wait_secs = if *name == "done" {
                    30
                } else {
                    readback_timeout_secs(name, input)
                };
                match await_tool_result(id, name, waiters, wait_secs, cancel_ref).await {
                    Some(payload) => payload,
                    // Fail-open on `done`: a webview that never answered must
                    // not strand the agent mid-done. Everything else reports
                    // the timeout honestly rather than as a bare "ok".
                    None if *name == "done" => {
                        r#"{"ok":true,"note":"validation_timed_out"}"#.to_string()
                    }
                    None => readback_timeout_msg(wait_secs),
                }
            } else {
                if *name == "done" {
                    r#"{"ok":true,"note":"no_validation_receiver"}"#.to_string()
                } else {
                    "ok".to_string()
                }
            };
            (
                i,
                name.clone(),
                json!({
                    "type": "tool_result",
                    "tool_use_id": id,
                    "content": content,
                }),
            )
        });

        let mut indexed: Vec<(usize, String, Value)> = futures::future::join_all(result_futs).await;
        indexed.sort_by_key(|(i, _, _)| *i);

        // Did `done` pass layout validation? Only then do we terminate.
        let mut done_validated = false;
        if done_called {
            for (_, name, result_val) in &indexed {
                if name != "done" {
                    continue;
                }
                let content = result_val
                    .get("content")
                    .and_then(|c| c.as_str())
                    .unwrap_or("");
                // Frontend ships `{"ok":true}` on pass, `{"ok":false,...}` on fail.
                match serde_json::from_str::<Value>(content) {
                    Ok(v) if v.get("ok").and_then(|o| o.as_bool()) == Some(true) => {
                        done_validated = true;
                    }
                    Ok(v) if v.get("ok").and_then(|o| o.as_bool()) == Some(false) => {
                        warn!(
                            "agent_loop: done rejected by layout validation ({} errors) — continuing",
                            v.get("error_count").and_then(|n| n.as_u64()).unwrap_or(0)
                        );
                        done_validated = false;
                    }
                    _ => {
                        // Non-JSON / unexpected — fail-open so we don't loop forever.
                        warn!("agent_loop: done validation returned unexpected payload; treating as ok");
                        done_validated = true;
                    }
                }
            }
        }

        let tool_results: Vec<Value> = indexed.into_iter().map(|(_, _, v)| v).collect();

        // Clean up any senders that are still hanging around (e.g. if a
        // tool_use_id never got a result and the webview reports late).
        for (id, _, _) in &tools_called {
            unregister_tool_waiters(id);
        }

        messages.push(StreamMessage {
            role: "user".to_string(),
            content: Value::Array(tool_results),
        });

        if done_called && done_validated {
            info!(
                "agent_loop: stopping after turn {} (done validated, stop_reason={})",
                turn, stop_reason
            );
            break;
        }
        if done_called && !done_validated {
            info!(
                "agent_loop: done failed layout validation on turn {} — agent must fix",
                turn
            );
            // Fall through to next turn with the validation errors in history.
        }

        // Source bookkeeping: note new fetches (for auto-evict aging) and
        // apply any keep_pages the model issued this turn — evicting every
        // page it didn't retain from the fetch_web result now in history. The
        // blob lives for exactly the review turn, then collapses.
        for (id, name, input) in &tools_called {
            match name.as_str() {
                "fetch_web" => {
                    if let Some(alias) = toolid_to_alias.get(id) {
                        sources_seen.entry(alias.clone()).or_insert(turn);
                    }
                }
                "keep_pages" => {
                    let alias = input
                        .get("source_id")
                        .and_then(|v| v.as_str())
                        .unwrap_or("")
                        .to_string();
                    if alias.is_empty() {
                        continue;
                    }
                    let keep: HashSet<usize> = input
                        .get("pages")
                        .and_then(|v| v.as_array())
                        .map(|a| a.iter().filter_map(|v| v.as_u64().map(|x| x as usize)).collect())
                        .unwrap_or_default();
                    let tool_use_id = alias_to_toolid.get(&alias).cloned();
                    let pages = FETCHED_SOURCES.lock().ok().and_then(|s| s.get(&alias).cloned());
                    match (tool_use_id, pages) {
                        (Some(tuid), Some(pages)) => {
                            if compact_source_in_messages(&mut messages, &tuid, &alias, &keep, &pages) {
                                compacted_sources.insert(alias.clone());
                                info!(
                                    "agent_loop: pruned source {} → kept {} of {} pages",
                                    alias,
                                    keep.len(),
                                    pages.len()
                                );
                            } else {
                                warn!("agent_loop: keep_pages {} matched no tool_result block", alias);
                            }
                        }
                        _ => warn!(
                            "agent_loop: keep_pages for unknown source_id '{}' (model may have mangled it)",
                            alias
                        ),
                    }
                }
                _ => {}
            }
        }

        // No backstop eviction here: a source the model failed to prune this
        // turn (mangled source_id, partial coverage of several pending
        // sources) just keeps the refusal loop alive next turn — cheap now
        // that the prompt prefix stays cached — and the auto-evict at the top
        // of the loop still guarantees the blob collapses within
        // AUTO_COMPACT_AFTER_TURNS regardless.

        info!(
            "agent_loop: continuing — turn {} stop_reason={} tools_processed={}",
            turn, stop_reason, tools_called.len()
        );
    }

    // Persist this run's sources into the session registry (instead of
    // dropping them, as we used to) so later turns can read_source any page
    // and dedup re-fetches of the same URL. The stash itself is bounded by
    // MAX_STASHED_SOURCES; format_prior_sources_block filters to aliases
    // whose pages are still live.
    if let Ok(mut sessions) = SESSION_SOURCES.lock() {
        let entry = sessions.entry(tab_id.clone()).or_default();
        for (key, alias) in url_to_alias.iter() {
            entry.url_to_alias.insert(key.clone(), alias.clone());
            if !entry.metas.iter().any(|m| &m.alias == alias) {
                entry.metas.push(SourceMeta {
                    alias: alias.clone(),
                    urls: key.split('\n').map(|s| s.to_string()).collect(),
                });
            }
        }
    }

    emit(
        &app,
        AgentEvent::Done {
            tab_id,
            batch_id,
            stop_reason: last_stop_reason,
            input_tokens: total_input_tokens,
            output_tokens: total_output_tokens,
            cache_read_tokens: total_cache_read_tokens,
            cache_creation_tokens: total_cache_creation_tokens,
        },
    );
}

/// One-shot LLM call that turns the user's first prompt (plus the agent's
/// final justification when available) into a 3-5 word session title. Falls
/// back to Err so the frontend can keep the heuristic name we already set.
///
/// Claude provider: uses Haiku for speed/cost. Codex provider: not wired up
/// — we return Err and the heuristic name stays. Easy to extend later if
/// needed.
#[tauri::command]
pub async fn generate_session_title(
    app_handle: AppHandle,
    prompt: String,
    justification: Option<String>,
) -> Result<String, String> {
    let (provider, credential) = resolve_provider_and_credential(&app_handle).await?;

    // Compose a single short user message. Justification carries what the
    // agent actually built ("simple income model for Tesla, 2022A-2028E…"),
    // which usually yields a tighter title than the prompt alone.
    let mut user = String::new();
    user.push_str("Suggest a 3-5 word session title for the work below. ");
    user.push_str("Output ONLY the title — no quotes, no punctuation, no leading 'Title:'. ");
    user.push_str("Title-case the first letter of each significant word.\n\n");
    user.push_str("User request:\n");
    user.push_str(prompt.trim());
    if let Some(j) = justification.as_ref().map(|s| s.trim()).filter(|s| !s.is_empty()) {
        user.push_str("\n\nWhat was built:\n");
        user.push_str(j);
    }

    let raw = match provider {
        Provider::Claude => {
            crate::engine::llm_providers::claude::complete_claude_brief(
                &credential,
                "claude-haiku-4-5",
                "",
                &user,
                32,
            )
            .await?
        }
        Provider::Codex => {
            // Title generation on the Codex provider isn't wired yet — the
            // Responses API would need a separate non-streaming helper. For
            // now we keep the heuristic session name in that case.
            return Err("title generation not supported on Codex provider".to_string());
        }
    };

    // Normalize: collapse whitespace, strip surrounding quotes/punct the
    // model sometimes adds, cap to 60 chars so it fits the tab UI.
    let mut t = raw.trim().to_string();
    t = t.trim_matches(|c: char| c == '"' || c == '\'' || c == '`' || c == '.' || c == ':').to_string();
    t = t.split_whitespace().collect::<Vec<_>>().join(" ");
    if t.chars().count() > 60 {
        t = t.chars().take(60).collect::<String>().trim_end().to_string();
        t.push('…');
    }
    if t.is_empty() {
        return Err("title generation returned empty".to_string());
    }
    info!("generate_session_title: \"{}\"", t);
    Ok(t)
}
