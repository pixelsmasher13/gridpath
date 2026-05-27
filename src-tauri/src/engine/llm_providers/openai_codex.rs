// OpenAI Codex Responses client (ChatGPT Plus/Pro subscription path).
//
// Hits `https://chatgpt.com/backend-api/codex/responses` using SSE, aggregates streamed
// `response.output_text.delta` events into a final string, and returns the same
// `(text, input_tokens, output_tokens)` shape as `openai::call_llm_api*` so callers
// can swap providers with a single match arm.
//
// Auth: caller passes the access token (`api_key`). The `chatgpt-account-id` header is
// derived on the fly from that token's JWT payload, so callers don't have to thread a
// second value through every dispatch site.

use base64::engine::general_purpose::URL_SAFE_NO_PAD;
use base64::Engine as _;
use eventsource_stream::Eventsource;
use futures_util::StreamExt;
use log::{error, info, warn};
use reqwest::Client;
use serde::Deserialize;
use std::time::Duration;

const JWT_CLAIM_PATH: &str = "https://api.openai.com/auth";

#[derive(Deserialize)]
struct JwtAuthClaim {
    chatgpt_account_id: Option<String>,
}

/// Pull `chatgpt_account_id` out of a Codex OAuth access token JWT.
fn account_id_from_token(access_token: &str) -> Result<String, String> {
    let mut parts = access_token.split('.');
    let _header = parts.next();
    let payload = parts
        .next()
        .ok_or_else(|| "ChatGPT access token is not a valid JWT".to_string())?;
    let decoded = URL_SAFE_NO_PAD
        .decode(payload)
        .map_err(|e| format!("Could not decode ChatGPT access token payload: {}", e))?;
    let value: serde_json::Value = serde_json::from_slice(&decoded)
        .map_err(|e| format!("ChatGPT access token payload was not JSON: {}", e))?;
    let claim = value.get(JWT_CLAIM_PATH).cloned().ok_or_else(|| {
        format!(
            "ChatGPT access token missing `{}` claim (re-sign in)",
            JWT_CLAIM_PATH
        )
    })?;
    let parsed: JwtAuthClaim = serde_json::from_value(claim).map_err(|e| {
        format!(
            "ChatGPT access token had unexpected `{}` shape: {}",
            JWT_CLAIM_PATH, e
        )
    })?;
    parsed
        .chatgpt_account_id
        .filter(|s| !s.is_empty())
        .ok_or_else(|| "ChatGPT access token missing chatgpt_account_id; please sign in again".into())
}

const CODEX_URL: &str = "https://chatgpt.com/backend-api/codex/responses";
const ORIGINATOR: &str = "gridpath";

#[derive(Deserialize, Debug)]
struct ResponseUsage {
    #[serde(default)]
    input_tokens: u32,
    #[serde(default)]
    output_tokens: u32,
    #[serde(default)]
    input_tokens_details: Option<InputTokensDetails>,
}

#[derive(Deserialize, Debug)]
struct InputTokensDetails {
    #[serde(default)]
    cached_tokens: u32,
}

fn user_agent() -> String {
    format!(
        "gridpath/{} ({}; {})",
        env!("CARGO_PKG_VERSION"),
        std::env::consts::OS,
        std::env::consts::ARCH
    )
}

// Streaming time budgets.
//
// `gpt-5.5` and other reasoning models can spend many minutes on hard prompts
// before any output bytes arrive, so a single global wall-clock timeout (which
// reqwest counts against the entire request including body read) is the wrong
// shape — it manifests as a confusing "Transport error: error decoding response
// body" the moment the deadline hits mid-stream.
//
// Instead we use:
//   * `connect_timeout`  – cap the initial TCP/TLS handshake.
//   * `read_timeout`     – cap each individual read on the body. As long as the
//                          server keeps trickling SSE keepalives or events, the
//                          stream stays alive.
const CONNECT_TIMEOUT: Duration = Duration::from_secs(30);
const READ_TIMEOUT: Duration = Duration::from_secs(180);

// ============================================================================
// Streaming + tool-use for the spreadsheet agent
//
// Mirrors `claude::stream_claude_with_tools` so the spreadsheet agent loop
// can consume either provider behind the same `StreamEvent` interface.
//
// Differences vs. Claude that this layer abstracts:
//   * Tool schema shape — OpenAI's Responses API wants
//       { type: "function", name, description, parameters }
//     vs. Anthropic's { name, description, input_schema }.
//   * Message format — OpenAI uses `input: [...]` with bare items like
//       { role, content: [{type:"input_text", text}] }
//       { type: "function_call", call_id, name, arguments: <json string> }
//       { type: "function_call_output", call_id, output: <string> }
//     vs. Anthropic's role-tagged content arrays with tool_use / tool_result.
//   * SSE events:
//       response.output_text.delta          -> StreamEvent::TextDelta
//       response.output_item.added type=function_call  -> setup
//       response.function_call_arguments.delta -> buffer JSON
//       response.function_call_arguments.done  -> StreamEvent::ToolCall
//       response.completed                  -> StreamEvent::MessageStop
// ============================================================================

use crate::engine::llm_providers::claude::{StreamEvent, StreamMessage};

/// Stream a Codex turn with tool-use. Same callback contract as
/// `stream_claude_with_tools` so the spreadsheet agent loop can drive either.
// `cache_key` is passed to OpenAI as `prompt_cache_key` so all turns of a
// session route to the same cache shard, increasing the chance of prefix-
// cache hits on tools + system across turns. Per OpenAI docs: keep each
// (prefix, cache_key) pair below ~15 req/min to avoid overflow —
// session-grained satisfies that.
pub async fn stream_codex_with_tools<F>(
    api_key: &str,
    system: &str,
    messages: Vec<StreamMessage>,
    tools: Vec<serde_json::Value>,
    max_tokens: usize,
    cache_key: &str,
    mut on_event: F,
) -> Result<(), String>
where
    F: FnMut(StreamEvent) + Send,
{
    let _ = max_tokens; // Codex backend ignores caps; left unused for parity.

    let account_id = account_id_from_token(api_key)?;
    let model = crate::engine::provider_config::get_model("openai-codex");

    // Convert Anthropic-shaped tool schemas to OpenAI's Responses function spec.
    let codex_tools: Vec<serde_json::Value> = tools
        .iter()
        .map(|t| {
            let name = t.get("name").cloned().unwrap_or(serde_json::Value::Null);
            let description = t.get("description").cloned().unwrap_or(serde_json::Value::Null);
            // Anthropic calls it `input_schema`; OpenAI calls it `parameters`.
            // Same JSON Schema shape inside.
            let parameters = t
                .get("input_schema")
                .or_else(|| t.get("parameters"))
                .cloned()
                .unwrap_or_else(|| serde_json::json!({"type": "object", "properties": {}}));
            serde_json::json!({
                "type": "function",
                "name": name,
                "description": description,
                "parameters": parameters,
                // OpenAI's Responses-API tools usually require strict: true; the
                // Codex backend tolerates either, but explicit strict=false lets
                // us evolve schemas without per-property `required` updates.
                "strict": false,
            })
        })
        .collect();
    // Hosted web_search is intentionally NOT injected on the Codex path —
    // the ChatGPT-subscription tier rejected it in testing. Claude
    // (claude.rs) gets the equivalent server tool. Codex agents fall back
    // to fetch_web with whatever URL they can produce on their own.

    // Translate Anthropic-shaped messages into Codex's input list.
    let mut codex_input: Vec<serde_json::Value> = Vec::new();
    for msg in &messages {
        let role = msg.role.as_str();
        // `content` is a Value — usually an array of blocks but sometimes a string.
        // Walk it and translate per-block to the right Codex shape.
        if let Some(arr) = msg.content.as_array() {
            for block in arr {
                let btype = block.get("type").and_then(|v| v.as_str()).unwrap_or("");
                match btype {
                    "text" | "input_text" | "output_text" => {
                        let text = block
                            .get("text")
                            .and_then(|v| v.as_str())
                            .unwrap_or("")
                            .to_string();
                        codex_input.push(serde_json::json!({
                            "type": "message",
                            "role": role,
                            "content": [{
                                "type": if role == "assistant" { "output_text" } else { "input_text" },
                                "text": text,
                            }],
                        }));
                    }
                    "tool_use" => {
                        // Anthropic emits this in assistant messages: { type, id, name, input }.
                        // OpenAI's equivalent is a top-level function_call item.
                        let call_id = block.get("id").and_then(|v| v.as_str()).unwrap_or("");
                        let name = block.get("name").and_then(|v| v.as_str()).unwrap_or("");
                        let input = block.get("input").cloned().unwrap_or(serde_json::json!({}));
                        codex_input.push(serde_json::json!({
                            "type": "function_call",
                            "call_id": call_id,
                            "name": name,
                            "arguments": input.to_string(),
                        }));
                    }
                    "tool_result" => {
                        let call_id = block
                            .get("tool_use_id")
                            .and_then(|v| v.as_str())
                            .unwrap_or("");
                        let output = block
                            .get("content")
                            .and_then(|v| v.as_str())
                            .map(|s| s.to_string())
                            .unwrap_or_else(|| block.get("content").map(|v| v.to_string()).unwrap_or_default());
                        codex_input.push(serde_json::json!({
                            "type": "function_call_output",
                            "call_id": call_id,
                            "output": output,
                        }));
                    }
                    _ => {
                        // Unknown block type — best effort: stringify and ship as text.
                        let text = block.to_string();
                        codex_input.push(serde_json::json!({
                            "type": "message",
                            "role": role,
                            "content": [{ "type": "input_text", "text": text }],
                        }));
                    }
                }
            }
        } else if let Some(text) = msg.content.as_str() {
            codex_input.push(serde_json::json!({
                "type": "message",
                "role": role,
                "content": [{
                    "type": if role == "assistant" { "output_text" } else { "input_text" },
                    "text": text,
                }],
            }));
        }
    }

    let reasoning = {
        let effort = crate::engine::provider_config::get_openai_codex_reasoning_effort();
        let trimmed = effort.trim().to_lowercase();
        if trimmed.is_empty() || trimmed == "none" || trimmed == "default" {
            None
        } else {
            Some(serde_json::json!({ "effort": trimmed, "summary": "auto" }))
        }
    };

    // store: false keeps responses out of OpenAI's 30-day server-side
    // retention — orthogonal to prompt caching, which is automatic and
    // free. The prompt_cache_key is what actually helps caching: it
    // co-routes all turns of a session to the same cache shard so the
    // tools + system prefix written in turn 1 is hit in turn 2+.
    let body = serde_json::json!({
        "model": model,
        "instructions": system,
        "input": codex_input,
        "tools": codex_tools,
        "tool_choice": "auto",
        "stream": true,
        "store": false,
        "prompt_cache_key": cache_key,
        "reasoning": reasoning,
    });

    info!(
        "Codex stream request: model={} tools={} input_items={}",
        model,
        codex_tools.len(),
        codex_input.len()
    );

    let client = Client::builder()
        .connect_timeout(CONNECT_TIMEOUT)
        .read_timeout(READ_TIMEOUT)
        .build()
        .map_err(|e| format!("Failed to build HTTP client: {}", e))?;

    let resp = client
        .post(CODEX_URL)
        .header("Authorization", format!("Bearer {}", api_key))
        .header("chatgpt-account-id", &account_id)
        .header("originator", ORIGINATOR)
        .header("User-Agent", user_agent())
        .header("Accept", "text/event-stream")
        .header("Content-Type", "application/json")
        .header("OpenAI-Beta", "responses=experimental")
        .json(&body)
        .send()
        .await
        .map_err(|e| format!("Failed to send request to OpenAI Codex: {}", e))?;

    let status = resp.status();
    info!("Codex stream HTTP {}", status);
    if !status.is_success() {
        let text = resp.text().await.unwrap_or_default();
        error!("OpenAI Codex stream error: HTTP {} body={}", status, text);
        if status.as_u16() == 401 || status.as_u16() == 403 {
            return Err(format!(
                "ChatGPT subscription auth failed ({}). Sign in again from Settings. Body: {}",
                status.as_u16(),
                text
            ));
        }
        return Err(format!("OpenAI Codex error {}: {}", status.as_u16(), text));
    }

    // Per-call_id JSON-args buffers. Codex streams function_call_arguments.delta
    // as raw partial-JSON chunks; we concatenate until function_call_arguments.done.
    use std::collections::HashMap;
    struct PendingCall {
        name: String,
        args: String,
    }
    let mut pending: HashMap<String, PendingCall> = HashMap::new();
    // item_id -> call_id mapping so delta/done events (which reference item_id)
    // can find their call payload.
    let mut item_to_call: HashMap<String, String> = HashMap::new();

    let mut input_tokens: u32 = 0;
    let mut output_tokens: u32 = 0;
    // Codex Responses API DOES report cache hits — they live in
    // `usage.input_tokens_details.cached_tokens`. Until now we hard-coded
    // 0 for the streaming path; that made the Usage tab show "cached = 0"
    // for every Codex run even when prompt caching was working. Propagate
    // it through the MessageStop event the same way Anthropic does.
    let mut cached_tokens: u32 = 0;
    let mut stop_reason = String::new();

    let stream_started = std::time::Instant::now();
    let mut events_received: u64 = 0;
    let mut last_event_type = String::new();
    let mut last_event_at = stream_started;

    let mut stream = resp.bytes_stream().eventsource();
    while let Some(ev) = stream.next().await {
        let event = match ev {
            Ok(e) => e,
            Err(e) => {
                // The reqwest `error decoding response body` message is a
                // generic wrapper from hyper — the actually interesting bit
                // is the source chain (TLS read closed, h2 stream reset,
                // body deadline exceeded, etc.). Walk it and log all.
                use std::error::Error as _;
                let mut chain: Vec<String> = vec![format!("{}", e)];
                let mut src: Option<&(dyn std::error::Error + 'static)> = e.source();
                let mut hops = 0;
                while let Some(s) = src {
                    chain.push(format!("caused_by: {}", s));
                    src = s.source();
                    hops += 1;
                    if hops > 6 { break; }
                }
                let now = std::time::Instant::now();
                warn!(
                    "Codex SSE error after {} events ({:?} since start, {:?} since last event '{}'): {}",
                    events_received,
                    now.duration_since(stream_started),
                    now.duration_since(last_event_at),
                    last_event_type,
                    chain.join(" | "),
                );
                break;
            }
        };
        events_received += 1;
        last_event_at = std::time::Instant::now();
        let payload: serde_json::Value = match serde_json::from_str(&event.data) {
            Ok(v) => v,
            Err(_) => continue,
        };
        let kind = payload
            .get("type")
            .and_then(|v| v.as_str())
            .unwrap_or("");
        last_event_type = kind.to_string();
        match kind {
            "response.created" => {
                on_event(StreamEvent::MessageStart);
            }
            "response.output_text.delta" => {
                if let Some(delta) = payload.get("delta").and_then(|v| v.as_str()) {
                    on_event(StreamEvent::TextDelta {
                        index: 0,
                        delta: delta.to_string(),
                    });
                }
            }
            "response.output_item.added" => {
                if let Some(item) = payload.get("item") {
                    if item.get("type").and_then(|v| v.as_str()) == Some("function_call") {
                        let item_id = item
                            .get("id")
                            .and_then(|v| v.as_str())
                            .unwrap_or("")
                            .to_string();
                        let call_id = item
                            .get("call_id")
                            .and_then(|v| v.as_str())
                            .unwrap_or("")
                            .to_string();
                        let name = item
                            .get("name")
                            .and_then(|v| v.as_str())
                            .unwrap_or("")
                            .to_string();
                        if !call_id.is_empty() {
                            pending.insert(call_id.clone(), PendingCall { name, args: String::new() });
                            if !item_id.is_empty() {
                                item_to_call.insert(item_id, call_id);
                            }
                        }
                    }
                }
            }
            "response.function_call_arguments.delta" => {
                let delta = payload
                    .get("delta")
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                // Find the call this delta belongs to via item_id.
                let item_id = payload
                    .get("item_id")
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                if let Some(call_id) = item_to_call.get(item_id) {
                    if let Some(p) = pending.get_mut(call_id) {
                        p.args.push_str(delta);
                    }
                }
            }
            "response.function_call_arguments.done" => {
                let item_id = payload
                    .get("item_id")
                    .and_then(|v| v.as_str())
                    .unwrap_or("");
                let final_args = payload
                    .get("arguments")
                    .and_then(|v| v.as_str())
                    .map(|s| s.to_string());
                if let Some(call_id) = item_to_call.remove(item_id) {
                    if let Some(mut p) = pending.remove(&call_id) {
                        if let Some(a) = final_args {
                            p.args = a;
                        }
                        let input: serde_json::Value = if p.args.trim().is_empty() {
                            serde_json::json!({})
                        } else {
                            serde_json::from_str(&p.args).unwrap_or(serde_json::json!({}))
                        };
                        on_event(StreamEvent::ToolCall {
                            index: 0,
                            tool_use_id: call_id,
                            name: p.name,
                            input,
                        });
                    }
                }
            }
            "response.completed" => {
                if let Some(response) = payload.get("response") {
                    if let Some(usage) = response.get("usage") {
                        if let Ok(parsed) =
                            serde_json::from_value::<ResponseUsage>(usage.clone())
                        {
                            input_tokens = parsed.input_tokens;
                            output_tokens = parsed.output_tokens;
                            cached_tokens = parsed
                                .input_tokens_details
                                .as_ref()
                                .map(|d| d.cached_tokens)
                                .unwrap_or(0);
                        }
                    }
                    if let Some(s) = response.get("stop_reason").and_then(|v| v.as_str()) {
                        stop_reason = s.to_string();
                    } else if let Some(s) = response.get("status").and_then(|v| v.as_str()) {
                        stop_reason = s.to_string();
                    }
                }
                // Don't fabricate "end_turn" — the agent loop reads that as
                // "stop, don't continue." Codex's response.completed signals
                // the response is done, but if there were function calls the
                // loop should keep going (the tool_results need to be sent
                // back). We emit a generic "completed" so the agent loop's
                // existing `no_tools_this_turn` and `done_called` checks
                // decide what to do.
                on_event(StreamEvent::MessageStop {
                    stop_reason: if stop_reason.is_empty() { "completed".to_string() } else { stop_reason.clone() },
                    input_tokens,
                    output_tokens,
                    cache_read_tokens: cached_tokens,
                    // Codex doesn't surface a separate "cache creation"
                    // counter the way Anthropic does — cache writes are
                    // implicit. Leave 0 so the Usage tab doesn't show a
                    // misleading number.
                    cache_creation_tokens: 0,
                });
            }
            "response.failed" | "response.error" => {
                let msg = payload
                    .get("error")
                    .and_then(|e| e.get("message"))
                    .and_then(|v| v.as_str())
                    .unwrap_or("OpenAI Codex returned a failed response");
                return Err(format!("OpenAI Codex stream failure: {}", msg));
            }
            _ => {}
        }
    }

    Ok(())
}
