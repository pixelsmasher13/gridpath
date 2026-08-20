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
        "Linefox help@linefox.ai"
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

    let ct = response
        .headers()
        .get("content-type")
        .and_then(|v| v.to_str().ok())
        .unwrap_or("")
        .to_lowercase();
    let texty = ct.contains("text/")
        || ct.contains("application/json")
        || ct.contains("application/xml")
        || ct.contains("+xml");

    let html = if texty {
        // Declared-text responses keep reqwest's charset-aware decoding.
        match response.text().await {
            Ok(t) => t,
            Err(e) => return (false, String::new(), String::new(), format!("Failed to read body: {}", e)),
        }
    } else {
        // Binary or undeclared content. The one binary format we understand
        // is PDF (press releases, IR decks, filings) — sniffed by magic, not
        // header, because IR CDNs routinely serve PDFs as
        // application/octet-stream.
        if response.content_length().unwrap_or(0) as usize > PDF_MAX_BYTES {
            return (false, String::new(), String::new(), format!("File too large (> {} MB)", PDF_MAX_BYTES / 1024 / 1024));
        }
        let bytes = match response.bytes().await {
            Ok(b) => b,
            Err(e) => return (false, String::new(), String::new(), format!("Failed to read body: {}", e)),
        };
        if bytes.len() > PDF_MAX_BYTES {
            return (false, String::new(), String::new(), format!("File too large (> {} MB)", PDF_MAX_BYTES / 1024 / 1024));
        }
        if looks_like_pdf(&bytes) {
            return pdf_page_result(bytes.to_vec()).await;
        }
        if ct.is_empty() {
            // No content-type header and not a PDF: assume text, as the
            // pre-PDF code did for header-less responses (no charset info
            // existed either way, so lossy UTF-8 matches the old decode).
            String::from_utf8_lossy(&bytes).into_owned()
        } else {
            return (false, String::new(), String::new(), format!("Non-text content type: {}", ct));
        }
    };

    // Extract title
    let title = extract_html_title(&html);

    // Strip HTML to readable text
    let text = html_to_readable_text(&html);

    // Truncate per page
    let text = truncate_on_char_boundary(text, 450_000);

    if text.len() < 50 {
        return (false, title, String::new(), "Page content too short (likely blocked or empty)".to_string());
    }

    (true, title, text, String::new())
}

/// Largest binary body we'll download for PDF extraction. Press releases and
/// IR decks are single-digit MB; anything bigger is likely a scan dump.
const PDF_MAX_BYTES: usize = 25 * 1024 * 1024;

/// PDF magic within the first 1KB — the spec tolerates junk before the
/// header, and some CDNs prepend a BOM.
fn looks_like_pdf(bytes: &[u8]) -> bool {
    let head = &bytes[..bytes.len().min(1024)];
    head.windows(5).any(|w| w == b"%PDF-")
}

/// Extract text from PDF bytes. pdf-extract panics on some malformed files
/// (unwraps inside its lopdf paths), so the panic is contained here rather
/// than taking down the fetch task.
fn extract_pdf_text(bytes: &[u8]) -> Result<String, String> {
    let extracted = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        pdf_extract::extract_text_from_mem(bytes)
    }))
    .map_err(|_| "PDF parser panicked (malformed file)".to_string())?;
    extracted.map_err(|e| format!("PDF text extraction failed: {}", e))
}

/// Build the (ok, title, text, error) tuple for a fetched PDF body.
async fn pdf_page_result(bytes: Vec<u8>) -> (bool, String, String, String) {
    // Extraction is CPU-bound (tens of ms on a press release, seconds on a
    // big deck) — run it off the async runtime.
    let extracted = match tauri::async_runtime::spawn_blocking(move || extract_pdf_text(&bytes)).await {
        Ok(r) => r,
        Err(e) => return (false, String::new(), String::new(), format!("PDF task join error: {}", e)),
    };
    let text = match extracted {
        Ok(t) => collapse_whitespace(&t),
        Err(e) => return (false, String::new(), String::new(), e),
    };
    if text.len() < 50 {
        return (
            false,
            String::new(),
            String::new(),
            "PDF contained no extractable text (likely a scanned/image PDF)".to_string(),
        );
    }
    // A PDF has no <title>; the first non-empty line is almost always the
    // document heading ("MANHATTAN ASSOCIATES REPORTS SECOND QUARTER ...").
    let title: String = text
        .lines()
        .find(|l| !l.trim().is_empty())
        .unwrap_or("PDF document")
        .trim()
        .chars()
        .take(120)
        .collect();
    let text = truncate_on_char_boundary(text, 450_000);
    (true, title, text, String::new())
}

/// Byte-cap a string without splitting a UTF-8 char (a mid-char slice
/// panics — the old `&text[..450000]` truncation could crash on CJK pages).
fn truncate_on_char_boundary(mut s: String, max_bytes: usize) -> String {
    if s.len() <= max_bytes {
        return s;
    }
    let mut cut = max_bytes;
    while cut > 0 && !s.is_char_boundary(cut) {
        cut -= 1;
    }
    s.truncate(cut);
    s.push_str("...");
    s
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

    collapse_whitespace(&result)
}

/// Collapse runs of spaces/tabs to one, 3+ newlines to two, and normalize
/// non-breaking spaces. Shared by the HTML and PDF extraction paths.
fn collapse_whitespace(s: &str) -> String {
    let mut cleaned = String::with_capacity(s.len());
    let mut prev_newline_count = 0;
    let mut prev_space = false;
    for ch in s.chars() {
        match ch {
            '\n' | '\r' => {
                prev_newline_count += 1;
                if prev_newline_count <= 2 {
                    cleaned.push('\n');
                }
                prev_space = false;
            }
            ' ' | '\t' | '\u{a0}' => {
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

#[cfg(test)]
mod tests {
    use super::*;

    /// Minimal one-page PDF with `text` in Helvetica; xref offsets computed
    /// so lopdf accepts the table. Keep `text` free of parentheses.
    fn tiny_pdf(text: &str) -> Vec<u8> {
        let stream = format!("BT /F1 12 Tf 72 720 Td ({text}) Tj ET");
        let objs: Vec<String> = vec![
            "<< /Type /Catalog /Pages 2 0 R >>".into(),
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>".into(),
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R \
             /Resources << /Font << /F1 5 0 R >> >> >>"
                .into(),
            format!("<< /Length {} >>\nstream\n{stream}\nendstream", stream.len()),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>".into(),
        ];
        let mut out = String::from("%PDF-1.4\n");
        let mut offsets = Vec::new();
        for (i, body) in objs.iter().enumerate() {
            offsets.push(out.len());
            out.push_str(&format!("{} 0 obj\n{}\nendobj\n", i + 1, body));
        }
        let xref_pos = out.len();
        out.push_str(&format!("xref\n0 {}\n0000000000 65535 f \n", objs.len() + 1));
        for off in &offsets {
            out.push_str(&format!("{:010} 00000 n \n", off));
        }
        out.push_str(&format!(
            "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n",
            objs.len() + 1,
            xref_pos
        ));
        out.into_bytes()
    }

    #[test]
    fn pdf_sniffing() {
        assert!(looks_like_pdf(&tiny_pdf("x")));
        // Magic past the start (spec-tolerated junk prefix).
        let mut padded = b"\xef\xbb\xbfjunk".to_vec();
        padded.extend_from_slice(&tiny_pdf("x"));
        assert!(looks_like_pdf(&padded));
        assert!(!looks_like_pdf(b"<html><body>hello</body></html>"));
        assert!(!looks_like_pdf(b""));
    }

    #[test]
    fn pdf_text_extraction() {
        let pdf = tiny_pdf("Net revenue 1234.5 million");
        let text = extract_pdf_text(&pdf).expect("extraction");
        assert!(text.contains("Net revenue 1234.5 million"), "got: {text:?}");
    }

    #[test]
    fn malformed_pdf_is_an_error_not_a_panic() {
        let garbage = b"%PDF-1.4\nnot actually a pdf at all";
        assert!(extract_pdf_text(garbage).is_err());
    }

    /// Live-network check against the real IR PDF from the MANH agent trace.
    /// Ignored by default; run manually: `cargo test web_fetcher -- --ignored`
    #[tokio::test]
    #[ignore]
    async fn live_ir_pdf_fetch() {
        let (out, ok) = fetch_and_extract_pages(vec![
            "https://ir.manh.com/static-files/e5fe0394-4403-4c81-a718-443f2bf4e9b0".to_string(),
        ])
        .await;
        assert_eq!(ok.len(), 1, "fetch failed:\n{out}");
        assert!(out.contains("Total revenue"), "income statement text missing:\n{out}");
    }

    #[test]
    fn truncate_respects_char_boundaries() {
        // 2-byte chars with an odd byte cap: the cut lands mid-char and must
        // back up instead of panicking.
        let s: String = "é".repeat(1000);
        let out = truncate_on_char_boundary(s, 999);
        assert!(out.ends_with("..."));
        assert!(out.len() <= 1002);
        let short = truncate_on_char_boundary("abc".to_string(), 999);
        assert_eq!(short, "abc");
    }
}
