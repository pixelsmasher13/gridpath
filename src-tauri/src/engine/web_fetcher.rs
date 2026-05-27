//! HTTP/HTML utilities for the `FetchPages` action — fetch a list of URLs in
//! parallel, strip boilerplate, and return readable text. Originally lived in
//! `automation_agent_engine.rs`; extracted here because nothing in this module
//! touches engine state, globals, or sibling modules.

use log::info;
use std::time::Duration;

/// Fetch and extract readable text from multiple URLs in parallel.
/// Returns (formatted output, list of URLs that succeeded).
pub async fn fetch_and_extract_pages(urls: Vec<String>) -> (String, Vec<String>) {
    use futures::future::join_all;

    let client = reqwest::Client::builder()
        .timeout(Duration::from_secs(10))
        .redirect(reqwest::redirect::Policy::limited(5))
        .build()
        .unwrap_or_default();

    let futures: Vec<_> = urls.iter().map(|url| {
        let client = client.clone();
        let url = url.clone();
        async move {
            fetch_single_page(&client, &url).await
        }
    }).collect();

    let results = join_all(futures).await;
    let success_count = results.iter().filter(|r| r.0).count();

    info!("FETCH_PAGES: {}/{} succeeded", success_count, urls.len());

    let mut output = format!("## Fetched Page Content ({}/{} succeeded)\n\n", success_count, urls.len());

    let mut failed_urls = Vec::new();
    let mut succeeded_urls = Vec::new();

    for (i, ((ok, title, text, error), url)) in results.into_iter().zip(urls.iter()).enumerate() {
        if ok {
            output.push_str(&format!("[Page {}: {}]\nTitle: {}\n\n{}\n\n---\n\n", i + 1, url, title, text));
            succeeded_urls.push(url.clone());
        } else {
            output.push_str(&format!("[Page {}: {}]\nFailed: {}\n\n---\n\n", i + 1, url, error));
            failed_urls.push(url.as_str());
        }
    }

    if !failed_urls.is_empty() {
        output.push_str(&format!(
            "\n⚠️ {} page(s) could not be fetched (blocked or JS-rendered). Use URL navigation to access them via the browser:\n",
            failed_urls.len()
        ));
        for url in &failed_urls {
            output.push_str(&format!("  - URL:1:{}\n", url));
        }
    }

    // Gentle reminder to save findings — this content won't persist in action history
    if success_count > 0 {
        output.push_str("\n💡 Use MEMORY_SAVE now to keep important findings — this content will not be available on the next turn.\n");
    }

    // Truncate total output to avoid blowing up context
    if output.len() > 1_500_000 {
        output.truncate(1_500_000);
        output.push_str("\n...[Output truncated]");
    }

    (output, succeeded_urls)
}

/// Fetch a single page and extract readable text.
/// Returns (ok, title, text, error)
async fn fetch_single_page(client: &reqwest::Client, url: &str) -> (bool, String, String, String) {
    // SEC EDGAR requires an identifying User-Agent (Name email) per their fair-access
    // policy — a generic browser UA gets a 403. https://www.sec.gov/os/accessing-edgar-data
    let user_agent = if is_sec_host(url) {
        "GridPath help@gridpath.dev"
    } else {
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    };

    let response = match client
        .get(url)
        .header("User-Agent", user_agent)
        .header("Accept", "text/html,application/xhtml+xml,*/*;q=0.9")
        .header("Accept-Language", "en-US,en;q=0.9")
        .send()
        .await
    {
        Ok(r) => r,
        Err(e) => return (false, String::new(), String::new(), e.to_string()),
    };

    let status = response.status();
    if !status.is_success() {
        return (false, String::new(), String::new(), format!("HTTP {}", status));
    }

    // Check content type — skip binary
    if let Some(ct) = response.headers().get("content-type") {
        let ct_str = ct.to_str().unwrap_or("");
        if !ct_str.contains("text/") && !ct_str.contains("application/json") && !ct_str.contains("application/xml") {
            return (false, String::new(), String::new(), format!("Non-text content type: {}", ct_str));
        }
    }

    let html = match response.text().await {
        Ok(t) => t,
        Err(e) => return (false, String::new(), String::new(), format!("Failed to read body: {}", e)),
    };

    // Extract title
    let title = extract_html_title(&html);

    // Strip HTML to readable text
    let text = html_to_readable_text(&html);

    // Truncate per page
    let text = if text.len() > 450000 {
        format!("{}...", &text[..450000])
    } else {
        text
    };

    if text.len() < 50 {
        return (false, title, String::new(), "Page content too short (likely blocked or empty)".to_string());
    }

    (true, title, text, String::new())
}

/// Decode `&#NNN;` and `&#xHH;` numeric HTML entities to their UTF-8 chars.
/// Leaves malformed sequences and unknown codepoints untouched.
fn decode_numeric_entities(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(amp) = rest.find("&#") {
        out.push_str(&rest[..amp]);
        let after = &rest[amp + 2..];
        if let Some(semi) = after.find(';') {
            // Cap entity body length to avoid scanning huge regions on stray '&'
            if semi <= 8 {
                let body = &after[..semi];
                let parsed = if let Some(hex) = body.strip_prefix('x').or_else(|| body.strip_prefix('X')) {
                    u32::from_str_radix(hex, 16).ok()
                } else if !body.is_empty() && body.bytes().all(|b| b.is_ascii_digit()) {
                    body.parse::<u32>().ok()
                } else {
                    None
                };
                if let Some(c) = parsed.and_then(char::from_u32) {
                    out.push(c);
                    rest = &after[semi + 1..];
                    continue;
                }
            }
        }
        // Not a valid entity — emit "&#" literally and continue past it
        out.push_str("&#");
        rest = after;
    }
    out.push_str(rest);
    out
}

/// True for sec.gov and its subdomains (e.g. www.sec.gov, efts.sec.gov).
fn is_sec_host(url: &str) -> bool {
    let lower = url.to_lowercase();
    let after_scheme = lower.split("://").nth(1).unwrap_or(&lower);
    let host = after_scheme.split('/').next().unwrap_or("");
    host == "sec.gov" || host.ends_with(".sec.gov")
}

/// Extract <title> from HTML
fn extract_html_title(html: &str) -> String {
    // Case-insensitive search for <title>...</title>
    let lower = html.to_lowercase();
    if let Some(start) = lower.find("<title") {
        if let Some(tag_end) = lower[start..].find('>') {
            let content_start = start + tag_end + 1;
            if let Some(end) = lower[content_start..].find("</title>") {
                return html[content_start..content_start + end].trim().to_string();
            }
        }
    }
    String::new()
}

/// Convert HTML to readable text by stripping tags and cleaning whitespace
fn html_to_readable_text(html: &str) -> String {
    let mut text = html.to_string();

    // Remove script, style, nav, header, footer blocks (case-insensitive via regex-like approach)
    // Use simple iterative approach since we don't have regex crate for this
    for tag in &["script", "style", "nav", "header", "footer", "noscript"] {
        loop {
            let lower = text.to_lowercase();
            let open = format!("<{}", tag);
            if let Some(start) = lower.find(&open) {
                let close = format!("</{}>", tag);
                if let Some(end_offset) = lower[start..].find(&close) {
                    let end = start + end_offset + close.len();
                    text = format!("{} {}", &text[..start], &text[end..]);
                } else {
                    // No closing tag — remove from open tag to end of next >
                    if let Some(gt) = text[start..].find('>') {
                        text = format!("{} {}", &text[..start], &text[start + gt + 1..]);
                    } else {
                        break;
                    }
                }
            } else {
                break;
            }
        }
    }

    // Strip all remaining HTML tags
    let mut result = String::with_capacity(text.len());
    let mut in_tag = false;
    for ch in text.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => {
                in_tag = false;
                result.push(' ');
            }
            _ if !in_tag => result.push(ch),
            _ => {}
        }
    }

    // Decode common named HTML entities, then any remaining numeric entities
    // (SEC EDGAR HTML uses &#160; / &#8217; / &#9744; etc. heavily — without
    // this they leak through and waste tokens).
    let result = result
        .replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&apos;", "'");
    let result = decode_numeric_entities(&result);
    // Normalize remaining non-breaking spaces (e.g. from &#160;) to regular spaces
    let result = result.replace('\u{a0}', " ");

    // Collapse whitespace
    let mut cleaned = String::with_capacity(result.len());
    let mut prev_newline_count = 0;
    let mut prev_space = false;
    for ch in result.chars() {
        match ch {
            '\n' | '\r' => {
                prev_newline_count += 1;
                if prev_newline_count <= 2 {
                    cleaned.push('\n');
                }
                prev_space = false;
            }
            ' ' | '\t' => {
                if !prev_space && prev_newline_count == 0 {
                    cleaned.push(' ');
                }
                prev_space = true;
            }
            _ => {
                cleaned.push(ch);
                prev_newline_count = 0;
                prev_space = false;
            }
        }
    }

    cleaned.trim().to_string()
}
