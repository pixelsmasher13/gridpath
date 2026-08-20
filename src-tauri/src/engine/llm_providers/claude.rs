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

/// Place a single `cache_control: ephemeral` breakpoint on the last content
/// block of the last message. Anthropic caches the entire prefix up to that
/// block, so on the next turn the whole conversation so far — workbook
/// snapshot, fetched filings (the 350 KB S1), and all prior tool history —
/// is served from cache at ~10% of input price instead of being re-billed in
/// full on every formatting pass.
///
/// Without this, only `tools` + `system` are cached (the one breakpoint
/// `cached_system` sets); the entire `messages` array is re-sent at full
/// price each turn. That's the dominant cost in a long build.
///
/// The breakpoint slides forward every turn — the caller hands us the freshly
/// grown `messages` each time and we always mark the new tail. This is the
/// canonical multi-turn pattern: cache_control markers are NOT part of the
/// cached content, so dropping the marker from an earlier block on the next
/// request does not break the prefix match. Anthropic always reads the
/// longest cached prefix it can find regardless of where the current
/// breakpoint sits. Short prefixes below the model's minimum-cacheable token
/// count simply don't cache — a harmless no-op on the first turn or two.
fn mark_last_message_cached(messages: &mut [StreamMessage]) {
    let Some(last) = messages.last_mut() else {
        return;
    };
    // Every agent-loop message uses the structured array form, but promote a
    // bare string defensively so we always have a block to annotate.
    if last.content.is_string() {
        let text = last.content.as_str().unwrap_or_default().to_string();
        last.content = serde_json::json!([{ "type": "text", "text": text }]);
    }
    if let Some(last_block) = last
        .content
        .as_array_mut()
        .and_then(|blocks| blocks.last_mut())
        .and_then(|block| block.as_object_mut())
    {
        last_block.insert(
            "cache_control".to_string(),
            serde_json::json!({ "type": "ephemeral" }),
        );
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

/// Extended-thinking budget for LEGACY models only (Haiku 4.5, Sonnet/Opus
/// 4.5 and older) — the pre-adaptive `{type:"enabled", budget_tokens}` API.
/// Current models (4.6+) use adaptive thinking with `output_config.effort`
/// as the depth control; sending budget_tokens to Opus 4.7+/Sonnet 5 is a
/// 400. Thinking tokens are billed as output, so `max_tokens` must exceed
/// this on the legacy path.
const LEGACY_THINKING_BUDGET_TOKENS: usize = 4000;

/// Which thinking API generation a Claude model speaks. Unknown/future
/// models are assumed CURRENT (newest API) — the legacy list is closed,
/// new models only ever move forward.
#[derive(PartialEq)]
enum ThinkingGen {
    /// `{type:"enabled", budget_tokens}`; no `output_config.effort`.
    Legacy,
    /// Adaptive thinking; effort supported but `xhigh` and the `display`
    /// param are not (thinking summaries are the default here anyway).
    Adaptive46,
    /// Adaptive + `display:"summarized"` (default is omitted → empty
    /// thinking text) + full effort range including `xhigh`.
    Current,
}

fn thinking_gen(model: &str) -> ThinkingGen {
    const LEGACY: [&str; 7] = [
        "haiku", "sonnet-4-5", "sonnet-4-0", "opus-4-5", "opus-4-1", "opus-4-0", "claude-3",
    ];
    if LEGACY.iter().any(|m| model.contains(m)) {
        ThinkingGen::Legacy
    } else if model.contains("opus-4-6") || model.contains("sonnet-4-6") {
        ThinkingGen::Adaptive46
    } else {
        ThinkingGen::Current
    }
}

/// Build the per-model `thinking` + `output_config` request fields, plus the
/// MINIMUM `max_tokens` the request needs (budgeted paths must have
/// max_tokens exceed budget_tokens or the API 400s). Effort comes from
/// provider_config (Settings → Effort); invalid values fall back to the
/// default.
fn thinking_and_output_config(
    model: &str,
) -> (Option<serde_json::Value>, Option<serde_json::Value>, usize) {
    let gen = thinking_gen(model);
    if gen == ThinkingGen::Legacy {
        return (
            Some(serde_json::json!({
                "type": "enabled",
                "budget_tokens": LEGACY_THINKING_BUDGET_TOKENS,
            })),
            None,
            LEGACY_THINKING_BUDGET_TOKENS + 8192,
        );
    }
    let raw = crate::engine::provider_config::get_claude_effort();
    let effort = match raw.trim().to_lowercase().as_str() {
        e @ ("low" | "medium" | "high" | "xhigh" | "max") => e.to_string(),
        _ => crate::engine::provider_config::DEFAULT_CLAUDE_EFFORT.to_string(),
    };
    if gen == ThinkingGen::Adaptive46 {
        // The 4.6 family straddles thinking generations: it still ACCEPTS
        // the legacy budget_tokens API (verified live 2026-07-18), and a
        // hard budget is deterministic where adaptive effort is only a
        // policy — an adaptive "medium" turn can still think far past 4k on
        // a hard build, which is exactly the multi-minute-stall surprise
        // this avoids on the default tier. So map the effort picker onto
        // budgets: same knob, guaranteed ceiling. Newest models reject
        // budget_tokens and stay adaptive below.
        let budget: usize = match effort.as_str() {
            "low" => 2_000,
            "medium" => 4_000,
            "high" => 10_000,
            // "xhigh" | "max" — keep budget + payload headroom well under
            // TURN_MAX_TOKENS (32k) so a max-effort think can't eat the turn.
            _ => 16_000,
        };
        return (
            Some(serde_json::json!({ "type": "enabled", "budget_tokens": budget })),
            None,
            budget + 8192,
        );
    }
    // Newest models (Opus 4.7+/Sonnet 5/Fable 5): adaptive is the only mode.
    // display:"summarized" restores readable thinking summaries — the
    // default is "omitted", which streams thinking blocks with EMPTY text
    // (the "model is awfully quiet" failure: a 3-minute think
    // indistinguishable from a hang).
    (
        Some(serde_json::json!({ "type": "adaptive", "display": "summarized" })),
        Some(serde_json::json!({ "effort": effort })),
        LEGACY_THINKING_BUDGET_TOKENS + 8192,
    )
}

#[derive(Serialize)]
struct LLMStreamRequest {
    model: String,
    max_tokens: usize,
    messages: Vec<StreamMessage>,
    #[serde(skip_serializing_if = "Option::is_none")]
    system: Option<Vec<SystemBlock>>,
    tools: Vec<serde_json::Value>,
    stream: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    thinking: Option<serde_json::Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    output_config: Option<serde_json::Value>,
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
    /// A chunk of the model's reasoning/plan arrived. On Claude this is a
    /// plaintext `thinking_delta`; on Codex it's a `reasoning_summary_text`
    /// delta. Surfaced live to the UI as a collapsible "Plan" block — distinct
    /// from `TextDelta` (user-facing prose) and from `ThinkingBlock` (the
    /// complete, signature-bearing block stitched back for the next turn).
    ReasoningDelta { delta: String },
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
    /// A complete extended-thinking block. The caller MUST stitch this back
    /// into the assistant message (ahead of the tool_use blocks) when it
    /// composes the next turn — Anthropic verifies the `signature` and rejects
    /// a tool-use continuation whose assistant turn dropped its thinking block.
    ThinkingBlock {
        index: u32,
        thinking: String,
        signature: String,
    },
    /// An encrypted thinking block Anthropic occasionally returns instead of
    /// plaintext. Opaque to us, but must be replayed verbatim for the same
    /// signature-verification reason as `ThinkingBlock`.
    RedactedThinking { index: u32, data: String },
    /// A hosted (server-executed) tool call — e.g. Anthropic's built-in
    /// `web_search`. Unlike `ToolCall`, the caller does NOT execute this or
    /// send a `tool_result` back — Anthropic runs it and continues the turn
    /// on its own. `block` is the complete original content block (e.g.
    /// `{"type":"server_tool_use","id":"...","name":"web_search","input":{"query":"..."}}`)
    /// — callers that replay assistant-message history MUST stitch this back
    /// verbatim at its original position, same as `ThinkingBlock`: Anthropic
    /// rejects a continuation whose assistant turn doesn't match what it
    /// originally returned, and it's not just thinking blocks that trip that
    /// check — dropping a server-tool block from the middle of the sequence
    /// does too (this was the "thinking blocks ... cannot be modified" 400
    /// that turned out to be a dropped web_search block, not a thinking-block
    /// bug per se).
    ServerToolUse {
        index: u32,
        tool_use_id: String,
        name: String,
        block: serde_json::Value,
    },
    /// The result of a hosted tool call (e.g. web_search results). Arrives
    /// fully formed (no deltas) at `content_block_start`. `block` is the
    /// complete original content block — must be replayed verbatim for the
    /// same reason as `ServerToolUse`.
    ServerToolResult {
        index: u32,
        tool_use_id: String,
        result_count: usize,
        block: serde_json::Value,
    },
}

#[derive(Debug, Clone)]
pub enum BlockKind {
    Text,
    Thinking,
    RedactedThinking,
    ToolUse { tool_use_id: String, name: String },
    ServerToolUse { tool_use_id: String, name: String },
}

/// Stream a Claude turn with tool-use. Invokes `on_event` for each event in
/// order; callers typically emit Tauri events from inside the callback.
///
/// `messages` is the full multi-turn history (caller manages it). `tools` is
/// the Anthropic tool schema array (each item: { name, description, input_schema }).
pub async fn stream_claude_with_tools<F>(
    api_key: &str,
    system_prompt: &str,
    mut messages: Vec<StreamMessage>,
    tools: Vec<serde_json::Value>,
    max_tokens: usize,
    mut on_event: F,
) -> Result<(), String>
where
    F: FnMut(StreamEvent) + Send,
{
    // Cache the conversation prefix so the growing history (workbook snapshot
    // + fetched filings + prior tool turns) is read from cache on every
    // subsequent turn instead of re-billed at full input price. See
    // `mark_last_message_cached`.
    mark_last_message_cached(&mut messages);
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

    // Thinking is billed as output. Budgeted paths (legacy + the 4.6
    // family) must have max_tokens exceed budget_tokens or the API 400s;
    // adaptive paths keep a sane floor for a long think PLUS the tool-call
    // payload (a 12K thinking block once consumed a 12,192 ceiling whole —
    // see TURN_MAX_TOKENS in commands.rs).
    let model = crate::engine::provider_config::get_model("claude");
    let (thinking, output_config, min_max_tokens) = thinking_and_output_config(&model);
    let max_tokens = max_tokens.max(min_max_tokens);
    let request_body = LLMStreamRequest {
        model,
        max_tokens,
        messages,
        system: cached_system(system_prompt.to_string(), is_oauth_token(api_key)),
        tools: tools_with_search,
        stream: true,
        thinking,
        output_config,
    };

    // Diagnostic: dump the ordered block shape of every assistant message that
    // carries a thinking/redacted_thinking block. The "blocks cannot be
    // modified" 400 means one of these doesn't match what Anthropic originally
    // sent — this reveals whether it's an empty signature, empty redacted
    // `data`, or a block ordered after a tool_use. Cheap; only fires on the
    // multi-turn (thinking-replay) path, not the first turn.
    for (mi, m) in request_body.messages.iter().enumerate() {
        if m.role != "assistant" { continue; }
        let Some(blocks) = m.content.as_array() else { continue; };
        let has_thinking = blocks.iter().any(|b| {
            matches!(
                b.get("type").and_then(|v| v.as_str()),
                Some("thinking") | Some("redacted_thinking")
            )
        });
        if !has_thinking { continue; }
        let shape: Vec<String> = blocks
            .iter()
            .map(|b| match b.get("type").and_then(|v| v.as_str()).unwrap_or("?") {
                "thinking" => format!(
                    "thinking(text={}B,sig={}B{})",
                    b.get("thinking").and_then(|v| v.as_str()).map(|s| s.len()).unwrap_or(0),
                    b.get("signature").and_then(|v| v.as_str()).map(|s| s.len()).unwrap_or(0),
                    if b.get("cache_control").is_some() { ",CACHED!" } else { "" },
                ),
                "redacted_thinking" => format!(
                    "redacted_thinking(data={}B{})",
                    b.get("data").and_then(|v| v.as_str()).map(|s| s.len()).unwrap_or(0),
                    if b.get("cache_control").is_some() { ",CACHED!" } else { "" },
                ),
                "tool_use" => format!(
                    "tool_use({})",
                    b.get("name").and_then(|v| v.as_str()).unwrap_or("?")
                ),
                "text" => format!(
                    "text({}B)",
                    b.get("text").and_then(|v| v.as_str()).map(|s| s.len()).unwrap_or(0)
                ),
                other => other.to_string(),
            })
            .collect();
        info!("Claude thinking-replay: msg[{}] assistant blocks = [{}]", mi, shape.join(", "));
    }

    // NOTE: do NOT use `.timeout()` here — that's a *total-request* deadline
    // that includes streaming the whole SSE body, so a long but healthy turn
    // (extended thinking + a large in-context source like a 450 KB filing)
    // gets aborted mid-stream as "request or response body error". Use a
    // per-read idle timeout instead: it resets on every chunk, so a stream can
    // run for minutes as long as tokens keep arriving, while a genuinely
    // stalled connection still gets killed.
    let client = Client::builder()
        .connect_timeout(Duration::from_secs(20))
        .read_timeout(Duration::from_secs(120))
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
        /// Accumulated plaintext for a thinking block.
        thinking_buf: String,
        /// Accumulated signature for a thinking block (arrives via
        /// `signature_delta`, usually one chunk at the block's end).
        signature: String,
        /// Opaque payload for a `redacted_thinking` block (delivered whole in
        /// `content_block_start`, no deltas).
        redacted_data: String,
        /// The raw `content_block` JSON from `content_block_start`, kept for
        /// block kinds that must be replayed verbatim (`server_tool_use`) —
        /// its `input` field gets overwritten with the accumulated delta JSON
        /// once the block finishes streaming.
        raw_block: serde_json::Value,
    }
    let mut blocks: HashMap<u32, BlockState> = HashMap::new();
    let mut input_tokens: u32 = 0;
    let mut output_tokens: u32 = 0;
    let mut cache_read_tokens: u32 = 0;
    let mut cache_creation_tokens: u32 = 0;
    let mut stop_reason: String = String::new();
    // Transport errors already propagate via `?` on the chunk read below; this
    // guards the other failure mode — the body ending cleanly without a
    // terminal `message_stop`, which would otherwise return Ok and let the
    // agent loop mistake a truncated turn for a clean stop.
    let mut saw_message_stop = false;

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
                        let block_type = block.get("type").and_then(|v| v.as_str()).unwrap_or("");
                        let kind = match block_type {
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
                            // Hosted (server-executed) tool call, e.g. Anthropic's
                            // built-in web_search — its `input` (the query) streams
                            // in via the same input_json_delta path as `tool_use`.
                            "server_tool_use" => BlockKind::ServerToolUse {
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
                            "thinking" => BlockKind::Thinking,
                            "redacted_thinking" => BlockKind::RedactedThinking,
                            _ => BlockKind::Text,
                        };
                        // A web_search_tool_result block arrives fully formed here
                        // (no deltas) — surface the raw block immediately (verbatim,
                        // for replay) rather than waiting for content_block_stop.
                        if block_type == "web_search_tool_result" {
                            let tool_use_id = block
                                .get("tool_use_id")
                                .and_then(|v| v.as_str())
                                .unwrap_or("")
                                .to_string();
                            let result_count = block
                                .get("content")
                                .and_then(|v| v.as_array())
                                .map(|a| a.len())
                                .unwrap_or(0);
                            on_event(StreamEvent::ServerToolResult {
                                index: idx,
                                tool_use_id,
                                result_count,
                                block: block.clone(),
                            });
                        }
                        // redacted_thinking carries its opaque `data` payload
                        // inline at start (no deltas follow).
                        let redacted_data = block
                            .get("data")
                            .and_then(|v| v.as_str())
                            .unwrap_or("")
                            .to_string();
                        blocks.insert(
                            idx,
                            BlockState {
                                kind: kind.clone(),
                                json_buf: String::new(),
                                thinking_buf: String::new(),
                                signature: String::new(),
                                redacted_data,
                                raw_block: block,
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
                            "thinking_delta" => {
                                let t = delta.get("thinking").and_then(|v| v.as_str()).unwrap_or("");
                                if let Some(b) = blocks.get_mut(&idx) {
                                    b.thinking_buf.push_str(t);
                                }
                                // Stream the reasoning to the UI live. The buffer
                                // above is still needed for the complete
                                // signature-bearing ThinkingBlock at block_stop.
                                if !t.is_empty() {
                                    on_event(StreamEvent::ReasoningDelta {
                                        delta: t.to_string(),
                                    });
                                }
                            }
                            "signature_delta" => {
                                let s = delta.get("signature").and_then(|v| v.as_str()).unwrap_or("");
                                if let Some(b) = blocks.get_mut(&idx) {
                                    b.signature.push_str(s);
                                }
                            }
                            _ => {}
                        }
                    }
                    "content_block_stop" => {
                        let idx = parsed.get("index").and_then(|v| v.as_u64()).unwrap_or(0) as u32;
                        if let Some(state) = blocks.remove(&idx) {
                            match &state.kind {
                                BlockKind::Thinking => {
                                    on_event(StreamEvent::ThinkingBlock {
                                        index: idx,
                                        thinking: state.thinking_buf.clone(),
                                        signature: state.signature.clone(),
                                    });
                                }
                                BlockKind::RedactedThinking => {
                                    on_event(StreamEvent::RedactedThinking {
                                        index: idx,
                                        data: state.redacted_data.clone(),
                                    });
                                }
                                BlockKind::ServerToolUse { tool_use_id, name } => {
                                    let input: serde_json::Value = if state.json_buf.trim().is_empty() {
                                        serde_json::Value::Object(serde_json::Map::new())
                                    } else {
                                        serde_json::from_str(&state.json_buf)
                                            .unwrap_or(serde_json::Value::Null)
                                    };
                                    // Patch the accumulated input into the raw block
                                    // captured at content_block_start (which has an
                                    // empty/stale `input`) so the emitted block is a
                                    // faithful, replayable copy of what Anthropic sent.
                                    let mut block = state.raw_block.clone();
                                    if let Some(obj) = block.as_object_mut() {
                                        obj.insert("input".to_string(), input);
                                    }
                                    on_event(StreamEvent::ServerToolUse {
                                        index: idx,
                                        tool_use_id: tool_use_id.clone(),
                                        name: name.clone(),
                                        block,
                                    });
                                }
                                _ => {}
                            }
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
                        saw_message_stop = true;
                    }
                    "ping" | "" => {}
                    other => {
                        info!("Claude stream: ignoring event '{}'", other);
                    }
                }
            }
        }
    }

    if !saw_message_stop {
        return Err(
            "Claude stream ended without a terminal message_stop (connection likely dropped mid-response)".to_string(),
        );
    }

    Ok(())
}


#[cfg(test)]
mod thinking_tests {
    use super::*;

    #[test]
    fn model_generations_classify_correctly() {
        assert!(matches!(thinking_gen("claude-haiku-4-5"), ThinkingGen::Legacy));
        assert!(matches!(thinking_gen("claude-haiku-4-5-20251001"), ThinkingGen::Legacy));
        assert!(matches!(thinking_gen("claude-sonnet-4-5"), ThinkingGen::Legacy));
        assert!(matches!(thinking_gen("claude-opus-4-5"), ThinkingGen::Legacy));
        assert!(matches!(thinking_gen("claude-sonnet-4-6"), ThinkingGen::Adaptive46));
        assert!(matches!(thinking_gen("claude-opus-4-6"), ThinkingGen::Adaptive46));
        assert!(matches!(thinking_gen("claude-opus-4-8"), ThinkingGen::Current));
        assert!(matches!(thinking_gen("claude-sonnet-5"), ThinkingGen::Current));
        assert!(matches!(thinking_gen("claude-fable-5"), ThinkingGen::Current));
        // Unknown/future models must land on the NEWEST API, never legacy —
        // sending budget_tokens to a future model would 400 every request.
        assert!(matches!(thinking_gen("claude-opus-5"), ThinkingGen::Current));
    }

    #[test]
    fn request_fields_per_generation() {
        // Legacy: budget shape, no output_config (effort unsupported there);
        // max_tokens floor covers budget + payload headroom.
        let (t, oc, floor) = thinking_and_output_config("claude-haiku-4-5");
        assert_eq!(t.unwrap()["type"], "enabled");
        assert!(oc.is_none());
        assert_eq!(floor, LEGACY_THINKING_BUDGET_TOKENS + 8192);

        // Current: adaptive + summarized display (omitted default = invisible
        // thinking in the UI), effort passed through.
        crate::engine::provider_config::set_claude_effort("xhigh");
        let (t, oc, _) = thinking_and_output_config("claude-opus-4-8");
        let t = t.unwrap();
        assert_eq!(t["type"], "adaptive");
        assert_eq!(t["display"], "summarized");
        assert_eq!(oc.unwrap()["effort"], "xhigh");

        // 4.6 family: effort maps onto legacy budget_tokens (deterministic
        // thinking ceiling — the default tier must never surprise with a
        // multi-minute adaptive think). No output_config on this path.
        crate::engine::provider_config::set_claude_effort("medium");
        let (t, oc, floor) = thinking_and_output_config("claude-sonnet-4-6");
        let t = t.unwrap();
        assert_eq!(t["type"], "enabled");
        assert_eq!(t["budget_tokens"], 4000);
        assert!(oc.is_none());
        assert_eq!(floor, 4000 + 8192);

        crate::engine::provider_config::set_claude_effort("high");
        let (t, _, floor) = thinking_and_output_config("claude-sonnet-4-6");
        assert_eq!(t.unwrap()["budget_tokens"], 10_000);
        assert_eq!(floor, 10_000 + 8192);

        crate::engine::provider_config::set_claude_effort("max");
        let (t, _, _) = thinking_and_output_config("claude-sonnet-4-6");
        assert_eq!(t.unwrap()["budget_tokens"], 16_000);

        // Garbage effort value falls back to the default (medium).
        crate::engine::provider_config::set_claude_effort("turbo");
        let (_, oc, _) = thinking_and_output_config("claude-opus-4-8");
        assert_eq!(oc.unwrap()["effort"], "medium");
        crate::engine::provider_config::set_claude_effort("");
    }
}
