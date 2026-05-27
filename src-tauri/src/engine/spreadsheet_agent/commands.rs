use log::{error, info, warn};
use once_cell::sync::Lazy;
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use std::sync::atomic::{AtomicBool, Ordering};
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
}

#[derive(Debug, Deserialize)]
pub struct SheetContext {
    pub name: String,
    pub row_count: u32,
    pub column_count: u32,
    /// Compact "A1 = 42" rows. The webview decides how to truncate so we
    /// don't have to encode windowing logic in Rust.
    pub cells_preview: String,
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

fn register_tool_result(tool_use_id: &str) -> oneshot::Receiver<String> {
    let (tx, rx) = oneshot::channel();
    if let Ok(mut map) = TOOL_RESULT_SENDERS.lock() {
        // If a sender already exists for this id (would be weird), drop it.
        map.insert(tool_use_id.to_string(), tx);
    }
    rx
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
    // No registered sender — late delivery, ignore.
    Ok(())
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

fn format_user_prompt(
    prompt: &str,
    ctx: &WorkbookContext,
    prior_batches: Option<&str>,
) -> String {
    let mut sheets_block = String::new();
    for s in &ctx.sheets {
        sheets_block.push_str(&format!(
            "## Sheet: {}\nDimensions: {} rows × {} columns (0-indexed)\nCurrent cells:\n{}\n\n",
            s.name, s.row_count, s.column_count, s.cells_preview
        ));
    }

    let prior = prior_batches
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("# Prior turns in this session\n{}\n\n", s))
        .unwrap_or_default();

    let focus = ctx
        .focus
        .as_ref()
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| format!("{}\n", s))
        .unwrap_or_default();

    // Today's date — the agent has no temporal grounding otherwise, so it
    // pattern-matches on its training cutoff (often 2024) and treats that
    // as "current year." That breaks financial models: it picks the wrong
    // historical-vs-projected boundary, projects from stale numbers, and
    // cites years that are already in the past as "estimated." A short
    // date header at the top of every user message fixes this in one line.
    let today = chrono::Local::now().format("%A, %B %-d, %Y").to_string();

    format!(
        "Today is {}.\n\nWorkbook: {}\n\n{}{}{}# User request\n{}",
        today, ctx.path, sheets_block, prior, focus, prompt
    )
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

    let user_message = format_user_prompt(&prompt, &workbook_context, prior_batches_context.as_deref());
    let initial_messages = vec![StreamMessage {
        role: "user".to_string(),
        content: json!([{ "type": "text", "text": user_message }]),
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

/// Cap on agent turns within one user request. Anthropic charges per turn so
/// runaway loops are expensive; 45 is enough headroom for a full model build
/// (fetch + data + bulk format + merges + widths + done) while still bounding
/// runaway agents.
const MAX_AGENT_TURNS: usize = 45;

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
    )
}

/// How long to wait for the webview to ship evaluated cell values back
/// after a write tool. Scales with the size of the operation — a 50×9
/// `set_range` needs Univer to apply each cell + recompute formulas + run
/// a React commit before it can serialize values back through Tauri IPC.
/// Base 15s covers tiny writes; we add ~1s per 30 cells touched, capped
/// at 90s so a runaway never hangs the agent loop.
fn readback_timeout_secs(name: &str, input: &Value) -> u64 {
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
        let tool_result_rxs: Arc<Mutex<HashMap<String, oneshot::Receiver<String>>>> =
            Arc::new(Mutex::new(HashMap::new()));
        let turn_stop_reason: Arc<Mutex<String>> = Arc::new(Mutex::new(String::new()));
        let turn_in: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_out: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_cache_read: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));
        let turn_cache_creation: Arc<Mutex<u32>> = Arc::new(Mutex::new(0));

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

        let on_event = move |event: StreamEvent| match event {
                StreamEvent::TextDelta { index, delta } => {
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
                StreamEvent::ToolCall {
                    index,
                    tool_use_id,
                    name,
                    input,
                } => {
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
                    // Register a frontend-result receiver only for tools
                    // that benefit from evaluated-value read-back
                    // (set_cell, set_range). `done` terminates the loop.
                    // `fetch_web` is handled entirely in Rust. Format /
                    // dimension / merge tools succeed-or-not — no point
                    // waiting for the webview to confirm, we just supply
                    // "ok" immediately during result collection.
                    if name != "done" && name != "fetch_web" && !tool_skips_readback(&name) {
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
                StreamEvent::MessageStart | StreamEvent::BlockStart { .. } => {}
            };

        // Pre-flight size audit — see log_turn_request for what's printed
        // and how to enable the JSONL dump.
        log_turn_request(&app, &batch_id, turn, &system, &tools, &messages);

        // Dispatch to the right provider. Each one consumes the same
        // `on_event` callback and produces the same StreamEvent stream
        // so the rest of the agent loop is provider-agnostic.
        let result = match provider {
            Provider::Claude => {
                stream_claude_with_tools(
                    &api_key,
                    &system,
                    messages.clone(),
                    tools.clone(),
                    4096,
                    on_event,
                ).await
            }
            Provider::Codex => {
                stream_codex_with_tools(
                    &api_key,
                    &system,
                    messages.clone(),
                    tools.clone(),
                    4096,
                    // Use tab_id as the OpenAI prompt_cache_key — stable
                    // across all turns of a single spreadsheet session.
                    &tab_id,
                    on_event,
                ).await
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
        //  - Claude explicitly called `done` → we have a justification, finish.
        //  - stop_reason == "end_turn" with no pending tool calls → finish.
        //  - No tools at all this turn → can't continue, treat as finished.
        let no_tools_this_turn = tools_called.is_empty();
        let should_stop =
            done_called || stop_reason == "end_turn" || no_tools_this_turn;

        if should_stop {
            info!(
                "agent_loop: stopping after turn {} (done={}, stop_reason={}, tools_this_turn={})",
                turn, done_called, stop_reason, tools_called.len()
            );
            break;
        }

        // Otherwise stop_reason was "tool_use" (or similar) — Claude wants
        // to keep working. Build the assistant message + a synthetic
        // tool_result for each tool_use so Claude can take the next turn.
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
        let mut rxs: HashMap<String, oneshot::Receiver<String>> =
            tool_result_rxs.lock().map(|mut g| std::mem::take(&mut *g)).unwrap_or_default();

        // For each tool_use, await the webview's tool_result (with a per-tool
        // timeout). If it never comes (timeout or dropped sender), fall back
        // to "ok" so Claude keeps moving rather than hang the loop. The
        // timeout is generous because Univer's formula engine + React commit
        // can take a moment on large set_range calls.
        let mut tool_results: Vec<Value> = Vec::with_capacity(tools_called.len());
        for (id, name, input) in &tools_called {
            if name == "done" {
                continue;
            }

            // `fetch_web` is handled entirely in Rust. Pull URLs out of
            // the tool input, call the existing fetcher, return its text.
            // We don't gate on a frontend receiver — the agent loop owns
            // the fetch directly.
            if name == "fetch_web" {
                let urls: Vec<String> = input
                    .get("urls")
                    .and_then(|v| v.as_array())
                    .map(|arr| arr.iter().filter_map(|v| v.as_str().map(|s| s.to_string())).collect())
                    .unwrap_or_default();
                let content = if urls.is_empty() {
                    "{\"error\": \"no urls provided\"}".to_string()
                } else {
                    let capped: Vec<String> = urls.into_iter().take(5).collect();
                    info!("fetch_web: fetching {} url(s) for batch {}", capped.len(), batch_id);
                    let (text, succeeded) =
                        crate::engine::web_fetcher::fetch_and_extract_pages(capped.clone()).await;
                
                    let truncated = if text.len() > 350_000 {
                        let mut t = text[..350_000].to_string();
                        t.push_str("\n...[truncated for token budget]");
                        t
                    } else {
                        text
                    };
                    json!({
                        "fetched": succeeded.len(),
                        "requested": capped.len(),
                        "content": truncated,
                    })
                    .to_string()
                };
                tool_results.push(json!({
                    "type": "tool_result",
                    "tool_use_id": id,
                    "content": content,
                }));
                continue;
            }

            // Format/dimension/merge tools never had a receiver registered
            // (see ToolCall handler above) — they get "ok" instantly.
            // Data-writing tools (set_cell, set_range) registered a receiver
            // and we wait for the frontend to ship back evaluated values
            // so Claude can self-correct on the next turn.
            //
            // The wait is **proportional to the operation size** — a 50×9
            // set_range needs Univer to apply cells, recalc formulas, and
            // commit a React render before shipping the snapshot back.
            // The flat 20s we used to wait was too short on big initial
            // model writes, causing the agent to fly blind on recovery
            // turns (it would call set_range, time out without seeing
            // evaluated values, then make wrong guesses about what landed).
            let content = if tool_skips_readback(name) {
                "ok".to_string()
            } else if let Some(rx) = rxs.remove(id) {
                let wait_secs = readback_timeout_secs(name, input);
                match timeout(Duration::from_secs(wait_secs), rx).await {
                    Ok(Ok(payload)) => payload,
                    Ok(Err(_)) => {
                        warn!("agent_loop: tool_result sender for {} dropped", id);
                        "ok".to_string()
                    }
                    Err(_) => {
                        warn!(
                            "agent_loop: tool_result for {} ({}) timed out after {}s",
                            id, name, wait_secs
                        );
                        "ok".to_string()
                    }
                }
            } else {
                "ok".to_string()
            };
            tool_results.push(json!({
                "type": "tool_result",
                "tool_use_id": id,
                "content": content,
            }));
        }

        // Clean up any senders that are still hanging around (e.g. if a
        // tool_use_id never got a result and the webview reports late).
        if let Ok(mut global) = TOOL_RESULT_SENDERS.lock() {
            for (id, _, _) in &tools_called {
                global.remove(id);
            }
        }

        messages.push(StreamMessage {
            role: "user".to_string(),
            content: Value::Array(tool_results),
        });

        info!(
            "agent_loop: continuing — turn {} stop_reason={} tools_processed={}",
            turn, stop_reason, tools_called.len()
        );
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
