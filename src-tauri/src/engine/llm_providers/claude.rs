use log::{info, error};
use reqwest::{Client, RequestBuilder};
use serde::Serialize;
use std::time::Duration;
use futures_util::StreamExt;

const ANTHROPIC_URL: &str = "https://api.anthropic.com/v1/messages";

/// Prefix of OAuth tokens issued by `claude setup-token` (Claude Code's
/// subscription auth flow). Anthropic API keys start with `sk-ant-api03-`,
/// so the two are unambiguous by prefix and we route auth accordingly:
///   * `sk-ant-oat01-...`  -> `Authorization: Bearer ...` + `anthropic-beta`
///   * anything else       -> `x-api-key: ...`
const OAUTH_TOKEN_PREFIX: &str = "sk-ant-oat01-";

/// Beta header required by Anthropic when authenticating with a Claude
/// Pro/Max OAuth token instead of an API key. Matches what the Claude Agent
/// SDK sets on subscription-billed `/v1/messages` calls.
const OAUTH_BETA_HEADER: &str = "oauth-2025-04-20";

/// Identity passphrase Anthropic requires as the **first** `system` block when
/// authenticating with a subscription OAuth token (`sk-ant-oat01-*`). Since
/// roughly March 2026 the server-side check rejects Sonnet/Opus OAuth calls
/// that don't lead with this exact string (any other wording -> 400/401). The
/// model still follows the user's real system prompt because subsequent
/// blocks override identity, but the first block must be byte-for-byte equal
/// to this. API-key callers must NOT send this prefix - they get billed for
/// the extra tokens and it's not needed.
const OAUTH_SYSTEM_IDENTITY: &str =
    "You are Claude Code, Anthropic's official CLI for Claude.";

fn is_oauth_token(token: &str) -> bool {
    sanitize_credential(token).starts_with(OAUTH_TOKEN_PREFIX)
}

/// Strip ALL whitespace from a credential before sending it to Anthropic.
///
/// The `claude setup-token` CLI prints OAuth tokens that wrap across the
/// terminal's column width, so copy-paste flows routinely deliver a token
/// with embedded newlines or surrounding spaces. Anthropic then 401s with
/// "Invalid bearer token" because the literal `\n` (or the truncated half)
/// isn't a valid token. OAuth tokens and `sk-ant-api03-*` API keys are both
/// guaranteed to be whitespace-free, so unconditional whitespace stripping
/// is safe defense-in-depth — much friendlier than asking users to paste
/// twice.
fn sanitize_credential(credential: &str) -> String {
    credential.chars().filter(|c| !c.is_whitespace()).collect()
}

/// Attach the right auth + beta headers for either API-key or subscription
/// (OAuth) credentials. All other headers (`Content-Type`,
/// `anthropic-version`) stay identical.
fn apply_auth_headers(builder: RequestBuilder, credential: &str) -> RequestBuilder {
    let clean = sanitize_credential(credential);
    if is_oauth_token(&clean) {
        info!(
            "Claude auth: Bearer (subscription OAuth), credential length={}",
            clean.len()
        );
        builder
            .header("Authorization", format!("Bearer {}", clean))
            .header("anthropic-beta", OAUTH_BETA_HEADER)
    } else {
        info!(
            "Claude auth: x-api-key (API credits), credential length={}",
            clean.len()
        );
        builder.header("x-api-key", clean)
    }
}

#[derive(Serialize)]
struct SystemBlock {
    #[serde(rename = "type")]
    block_type: &'static str,
    text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    cache_control: Option<CacheControl>,
}

#[derive(Serialize)]
struct CacheControl {
    #[serde(rename = "type")]
    cache_type: &'static str,
}

/// Build the `system` array for `/v1/messages`. On the OAuth subscription
/// path we MUST lead with the Claude Code identity block (see
/// `OAUTH_SYSTEM_IDENTITY`); the user's real system prompt follows as a
/// second cached block and supplies the actual instructions. API-key callers
/// just send the user's prompt with cache_control as before.
fn cached_system(text: String, oauth: bool) -> Option<Vec<SystemBlock>> {
    let identity = || SystemBlock {
        block_type: "text",
        text: OAUTH_SYSTEM_IDENTITY.to_string(),
        cache_control: None,
    };
    let user_block = |text: String| SystemBlock {
        block_type: "text",
        text,
        cache_control: Some(CacheControl { cache_type: "ephemeral" }),
    };

    match (text.is_empty(), oauth) {
        // API key + no system prompt -> omit the field entirely.
        (true, false) => None,
        // OAuth + no system prompt -> still required to lead with identity.
        (true, true) => Some(vec![identity()]),
        // API key + system prompt -> single cached user block.
        (false, false) => Some(vec![user_block(text)]),
        // OAuth + system prompt -> identity first, then user prompt cached.
        (false, true) => Some(vec![identity(), user_block(text)]),
    }
}

// ============================================================================
// One-shot non-streaming completion (used for tiny side calls like
// session-title generation — not the agent loop).
// ============================================================================

/// Issue a single non-streaming `/v1/messages` request and return the
/// concatenated text from the first assistant turn. No tools, no SSE,
/// no caching tricks — this is the simplest possible Claude call.
///
/// Caller-provided `model` lets us use Haiku for cheap helpers while the
/// main agent runs on Sonnet/Opus. Caller-provided `system` and `user`
/// strings are sent verbatim (OAuth identity block prepended automatically
/// when the credential is a subscription token).
pub async fn complete_claude_brief(
    api_key: &str,
    model: &str,
    system: &str,
    user: &str,
    max_tokens: usize,
) -> Result<String, String> {
    #[derive(Serialize)]
    struct OneShotRequest {
        model: String,
        max_tokens: usize,
        messages: Vec<StreamMessage>,
        #[serde(skip_serializing_if = "Option::is_none")]
        system: Option<Vec<SystemBlock>>,
        stream: bool,
    }

    let request_body = OneShotRequest {
        model: model.to_string(),
        max_tokens,
        messages: vec![StreamMessage {
            role: "user".to_string(),
            content: serde_json::json!([{ "type": "text", "text": user }]),
        }],
        system: cached_system(system.to_string(), is_oauth_token(api_key)),
        stream: false,
    };

    let client = Client::builder()
        .timeout(Duration::from_secs(20))
        .build()
        .map_err(|e| format!("Failed to create HTTP client: {}", e))?;

    let builder = client
        .post(ANTHROPIC_URL)
        .header("Content-Type", "application/json")
        .header("anthropic-version", "2023-06-01")
        .json(&request_body);
    let response = apply_auth_headers(builder, api_key)
        .send()
        .await
        .map_err(|e| format!("Failed to send brief Claude request: {}", e))?;

    let status = response.status();
    if !status.is_success() {
        let body = response.text().await.unwrap_or_default();
        return Err(format!("Claude brief HTTP {}: {}", status, body));
    }

    let body: serde_json::Value = response
        .json()
        .await
        .map_err(|e| format!("Claude brief: bad JSON: {}", e))?;

    // Concatenate every `text` block in `content` (usually just one).
    let mut out = String::new();
    if let Some(blocks) = body.get("content").and_then(|v| v.as_array()) {
        for blk in blocks {
            if blk.get("type").and_then(|v| v.as_str()) == Some("text") {
                if let Some(t) = blk.get("text").and_then(|v| v.as_str()) {
                    out.push_str(t);
                }
            }
        }
    }
    Ok(out)
}

// ============================================================================
// Streaming + tool-use (the only entry point used by the spreadsheet agent)
//
// Anthropic's SSE protocol is self-framing: each content block has
// start/delta/stop events with a stable `index`, and tool_use blocks
// accumulate JSON via `input_json_delta` until `content_block_stop`, so
// there's no partial-JSON parsing on our side — we just buffer the chunks
// and parse once the block ends.
//
// See https://docs.anthropic.com/en/api/messages-streaming for the event spec.
// ============================================================================

#[derive(Serialize)]
struct LLMStreamRequest {
    model: String,
    max_tokens: usize,
    messages: Vec<StreamMessage>,
    #[serde(skip_serializing_if = "Option::is_none")]
    system: Option<Vec<SystemBlock>>,
    tools: Vec<serde_json::Value>,
    stream: bool,
}

#[derive(Serialize, Debug, Clone)]
pub struct StreamMessage {
    pub role: String,
    pub content: serde_json::Value,
}

/// Events surfaced to the caller as a streaming Claude turn unfolds.
/// `index` is Claude's content_block index — useful when caller needs to
/// reconstruct interleaved text + tool_use ordering.
#[derive(Debug, Clone)]
pub enum StreamEvent {
    /// Streaming has begun. Emitted once.
    MessageStart,
    /// A new content block has begun (text or tool_use). Carried for future
    /// reconstruction of interleaved text + tool_use ordering — the current
    /// spreadsheet agent loop only branches on `ToolCall` / `TextDelta`.
    BlockStart {
        #[allow(dead_code)]
        index: u32,
        #[allow(dead_code)]
        kind: BlockKind,
    },
    /// Text delta arrived (only for text blocks).
    TextDelta { index: u32, delta: String },
    /// A complete tool_use block has finished streaming and its JSON input
    /// has been parsed. This is the event the agent loop dispatches on.
    ToolCall {
        index: u32,
        tool_use_id: String,
        name: String,
        input: serde_json::Value,
    },
    /// Stream finished cleanly. Final usage counts attached.
    /// `cache_read_tokens` / `cache_creation_tokens` are populated from
    /// Anthropic's `message_start.usage` block — non-zero whenever the
    /// ephemeral cached system prompt is hit (read) or refreshed (creation).
    /// On the Codex provider these are always 0 since the Responses API
    /// doesn't expose comparable cache hit stats.
    MessageStop {
        stop_reason: String,
        input_tokens: u32,
        output_tokens: u32,
        cache_read_tokens: u32,
        cache_creation_tokens: u32,
    },
}

#[derive(Debug, Clone)]
pub enum BlockKind {
    Text,
    ToolUse { tool_use_id: String, name: String },
}

/// Stream a Claude turn with tool-use. Invokes `on_event` for each event in
/// order; callers typically emit Tauri events from inside the callback.
///
/// `messages` is the full multi-turn history (caller manages it). `tools` is
/// the Anthropic tool schema array (each item: { name, description, input_schema }).
pub async fn stream_claude_with_tools<F>(
    api_key: &str,
    system_prompt: &str,
    messages: Vec<StreamMessage>,
    tools: Vec<serde_json::Value>,
    max_tokens: usize,
    mut on_event: F,
) -> Result<(), String>
where
    F: FnMut(StreamEvent) + Send,
{
    // Append Anthropic's hosted web_search tool so the agent can discover
    // sources (10-Ks, news, product pages) without us round-tripping for a
    // URL we'd just have to guess. Results stream back inline in the
    // model's response — no readback required, no frontend handler needed.
    // Hosted-tool name is fixed by Anthropic spec; do NOT rename to avoid
    // collision with any future user-defined tool called "web_search".
    let mut tools_with_search = tools;
    tools_with_search.push(serde_json::json!({
        "type": "web_search_20250305",
        "name": "web_search",
        "max_uses": 5,
    }));

    let request_body = LLMStreamRequest {
        model: crate::engine::provider_config::get_model("claude"),
        max_tokens,
        messages,
        system: cached_system(system_prompt.to_string(), is_oauth_token(api_key)),
        tools: tools_with_search,
        stream: true,
    };

    let client = Client::builder()
        .timeout(Duration::from_secs(120))
        .build()
        .map_err(|e| format!("Failed to create HTTP client: {}", e))?;

    let builder = client
        .post(ANTHROPIC_URL)
        .header("Content-Type", "application/json")
        .header("Accept", "text/event-stream")
        .header("anthropic-version", "2023-06-01")
        .json(&request_body);
    let response = apply_auth_headers(builder, api_key)
        .send()
        .await
        .map_err(|e| format!("Failed to send streaming request to Claude API: {}", e))?;

    let status = response.status();
    info!(
        "Claude stream: HTTP {} (tools={}, model={})",
        status,
        request_body.tools.len(),
        request_body.model
    );
    if !status.is_success() {
        let body = response.text().await.unwrap_or_default();
        error!("Claude stream error: HTTP {} body={}", status, body);
        return Err(format!("Claude API error {}: {}", status, body));
    }

    use std::collections::HashMap;
    struct BlockState {
        kind: BlockKind,
        json_buf: String,
    }
    let mut blocks: HashMap<u32, BlockState> = HashMap::new();
    let mut input_tokens: u32 = 0;
    let mut output_tokens: u32 = 0;
    let mut cache_read_tokens: u32 = 0;
    let mut cache_creation_tokens: u32 = 0;
    let mut stop_reason: String = String::new();

    let mut sse_buf = String::new();
    let mut stream = response.bytes_stream();

    while let Some(chunk) = stream.next().await {
        let chunk = chunk.map_err(|e| format!("stream read failed: {}", e))?;
        sse_buf.push_str(&String::from_utf8_lossy(&chunk));

        loop {
            let Some(sep_idx) = sse_buf.find("\n\n") else { break; };
            let record = sse_buf[..sep_idx].to_string();
            sse_buf.drain(..sep_idx + 2);

            for line in record.lines() {
                let Some(payload) = line.strip_prefix("data:") else { continue; };
                let payload = payload.trim();
                if payload.is_empty() || payload == "[DONE]" {
                    continue;
                }
                let parsed: serde_json::Value = match serde_json::from_str(payload) {
                    Ok(v) => v,
                    Err(e) => {
                        error!("Malformed SSE payload: {} (payload={})", e, payload);
                        continue;
                    }
                };
                let event_type = parsed.get("type").and_then(|v| v.as_str()).unwrap_or("");

                // Log structural events only (block_start/stop, message_*) — content_block_delta
                // fires for every text token and would flood the log.
                if event_type != "content_block_delta" && event_type != "ping" {
                    info!("Claude SSE: event_type={}", event_type);
                }
                match event_type {
                    "message_start" => {
                        if let Some(usage) = parsed.get("message").and_then(|m| m.get("usage")) {
                            input_tokens = usage.get("input_tokens").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                            cache_read_tokens = usage.get("cache_read_input_tokens").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                            cache_creation_tokens = usage.get("cache_creation_input_tokens").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        }
                        on_event(StreamEvent::MessageStart);
                    }
                    "content_block_start" => {
                        let idx = parsed.get("index").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        let block = parsed.get("content_block").cloned().unwrap_or_default();
                        let kind = match block.get("type").and_then(|v| v.as_str()).unwrap_or("") {
                            "tool_use" => BlockKind::ToolUse {
                                tool_use_id: block
                                    .get("id")
                                    .and_then(|v| v.as_str())
                                    .unwrap_or("")
                                    .to_string(),
                                name: block
                                    .get("name")
                                    .and_then(|v| v.as_str())
                                    .unwrap_or("")
                                    .to_string(),
                            },
                            _ => BlockKind::Text,
                        };
                        blocks.insert(
                            idx,
                            BlockState {
                                kind: kind.clone(),
                                json_buf: String::new(),
                            },
                        );
                        on_event(StreamEvent::BlockStart { index: idx, kind });
                    }
                    "content_block_delta" => {
                        let idx = parsed.get("index").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        let delta = parsed.get("delta").cloned().unwrap_or_default();
                        let delta_type = delta.get("type").and_then(|v| v.as_str()).unwrap_or("");
                        match delta_type {
                            "text_delta" => {
                                let text = delta
                                    .get("text")
                                    .and_then(|v| v.as_str())
                                    .unwrap_or("")
                                    .to_string();
                                on_event(StreamEvent::TextDelta { index: idx, delta: text });
                            }
                            "input_json_delta" => {
                                let partial = delta
                                    .get("partial_json")
                                    .and_then(|v| v.as_str())
                                    .unwrap_or("");
                                if let Some(b) = blocks.get_mut(&idx) {
                                    b.json_buf.push_str(partial);
                                }
                            }
                            _ => {}
                        }
                    }
                    "content_block_stop" => {
                        let idx = parsed.get("index").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        if let Some(state) = blocks.remove(&idx) {
                            if let BlockKind::ToolUse { tool_use_id, name } = state.kind {
                                let input: serde_json::Value = if state.json_buf.trim().is_empty() {
                                    serde_json::Value::Object(serde_json::Map::new())
                                } else {
                                    match serde_json::from_str(&state.json_buf) {
                                        Ok(v) => v,
                                        Err(e) => {
                                            error!(
                                                "Tool-use JSON parse failed (name={} id={}): {} — buf={}",
                                                name, tool_use_id, e, state.json_buf
                                            );
                                            continue;
                                        }
                                    }
                                };
                                on_event(StreamEvent::ToolCall {
                                    index: idx,
                                    tool_use_id,
                                    name,
                                    input,
                                });
                            }
                        }
                    }
                    "message_delta" => {
                        if let Some(delta) = parsed.get("delta") {
                            if let Some(reason) = delta.get("stop_reason").and_then(|v| v.as_str()) {
                                stop_reason = reason.to_string();
                            }
                        }
                        if let Some(usage) = parsed.get("usage") {
                            output_tokens = usage.get("output_tokens").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        }
                    }
                    "message_stop" => {
                        on_event(StreamEvent::MessageStop {
                            stop_reason: stop_reason.clone(),
                            input_tokens,
                            output_tokens,
                            cache_read_tokens,
                            cache_creation_tokens,
                        });
                    }
                    "ping" | "" => {}
                    other => {
                        info!("Claude stream: ignoring event '{}'", other);
                    }
                }
            }
        }
    }

    Ok(())
}

