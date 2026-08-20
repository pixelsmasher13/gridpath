//! SEC EDGAR filing lookup — deterministic primary-source document URLs.
//!
//! Two public endpoints, no auth:
//!   - https://www.sec.gov/files/company_tickers.json       (ticker → CIK map)
//!   - https://data.sec.gov/submissions/CIK##########.json  (recent filings)
//!
//! Motivation (benchmark round 3, 2026-08): the agent's build-quality gap
//! traced to sourcing — guessed deep URLs 404'd and aggregator summaries
//! carry restated/rounded figures on mixed bases. This tool removes both
//! failure modes: the model names a company, gets exact filing URLs, and
//! `fetch_web`s the authoritative document.
//!
//! SEC's fair-access policy asks for a descriptive User-Agent and modest
//! volume; the ticker map (~1 MB) is cached for the process lifetime.

use log::info;
use once_cell::sync::Lazy;
use serde_json::{json, Value};
use std::sync::Mutex;
use std::time::Duration;

#[derive(Clone)]
struct TickerEntry {
    ticker: String,
    title: String,
    cik: u64,
}

static TICKER_CACHE: Lazy<Mutex<Option<Vec<TickerEntry>>>> = Lazy::new(|| Mutex::new(None));

fn client() -> reqwest::Client {
    reqwest::Client::builder()
        .timeout(Duration::from_secs(10))
        .user_agent("GridPath spreadsheet agent (support@gridpath.dev)")
        .build()
        .unwrap_or_default()
}

async fn load_tickers(client: &reqwest::Client) -> Result<Vec<TickerEntry>, String> {
    if let Ok(cache) = TICKER_CACHE.lock() {
        if let Some(v) = cache.as_ref() {
            return Ok(v.clone());
        }
    }
    let resp = client
        .get("https://www.sec.gov/files/company_tickers.json")
        .send()
        .await
        .map_err(|e| format!("ticker map fetch failed: {e}"))?;
    let body: Value = resp
        .json()
        .await
        .map_err(|e| format!("ticker map parse failed: {e}"))?;
    let mut entries = Vec::new();
    if let Some(map) = body.as_object() {
        for v in map.values() {
            let (Some(ticker), Some(title), Some(cik)) = (
                v.get("ticker").and_then(|x| x.as_str()),
                v.get("title").and_then(|x| x.as_str()),
                v.get("cik_str").and_then(|x| x.as_u64()),
            ) else {
                continue;
            };
            entries.push(TickerEntry {
                ticker: ticker.to_string(),
                title: title.to_string(),
                cik,
            });
        }
    }
    if entries.is_empty() {
        return Err("ticker map was empty or malformed".to_string());
    }
    if let Ok(mut cache) = TICKER_CACHE.lock() {
        *cache = Some(entries.clone());
    }
    Ok(entries)
}

fn resolve(entries: &[TickerEntry], query: &str) -> Result<TickerEntry, Value> {
    let q = query.trim().to_lowercase();
    // Pass 1: exact ticker match.
    if let Some(e) = entries.iter().find(|e| e.ticker.to_lowercase() == q) {
        return Ok(e.clone());
    }
    // Pass 2: company-title substring; prefer the shortest matching title
    // (so "apple" → "Apple Inc." over "Apple Hospitality REIT Inc").
    let mut matches: Vec<&TickerEntry> =
        entries.iter().filter(|e| e.title.to_lowercase().contains(&q)).collect();
    matches.sort_by_key(|e| e.title.len());
    match matches.len() {
        0 => Err(json!({
            "error": format!("no SEC registrant matched {query:?} — try the exact ticker symbol"),
        })),
        1 => Ok(matches[0].clone()),
        _ => {
            // Ambiguous: pick the shortest-title match; the caller attaches
            // the runners-up as `other_matches` so a wrong pick is one
            // obvious correction away.
            let best = matches[0].clone();
            info!(
                "edgar_lookup: {:?} matched {} candidates, picked {}",
                query,
                matches.len(),
                best.ticker
            );
            Ok(best)
        }
    }
}

fn alternatives(entries: &[TickerEntry], query: &str, chosen: &TickerEntry) -> Vec<Value> {
    let q = query.trim().to_lowercase();
    entries
        .iter()
        .filter(|e| e.title.to_lowercase().contains(&q) && e.cik != chosen.cik)
        .take(5)
        .map(|e| json!({ "ticker": e.ticker, "company": e.title }))
        .collect()
}

/// Look up a company's recent SEC filings and return exact document URLs.
pub async fn lookup(company: &str, forms: &[String], count: usize) -> Value {
    let client = client();
    let entries = match load_tickers(&client).await {
        Ok(e) => e,
        Err(e) => return json!({ "error": e }),
    };
    let entry = match resolve(&entries, company) {
        Ok(e) => e,
        Err(err) => return err,
    };
    let wanted: Vec<String> = if forms.is_empty() {
        vec!["10-K".to_string(), "10-Q".to_string()]
    } else {
        forms.iter().map(|f| f.trim().to_uppercase()).collect()
    };

    let sub_url = format!("https://data.sec.gov/submissions/CIK{:010}.json", entry.cik);
    let sub: Value = match client.get(&sub_url).send().await {
        Ok(r) => match r.json().await {
            Ok(v) => v,
            Err(e) => return json!({ "error": format!("submissions parse failed: {e}") }),
        },
        Err(e) => return json!({ "error": format!("submissions fetch failed: {e}") }),
    };

    let recent = &sub["filings"]["recent"];
    let get_arr = |k: &str| recent.get(k).and_then(|v| v.as_array()).cloned().unwrap_or_default();
    let form_a = get_arr("form");
    let acc_a = get_arr("accessionNumber");
    let doc_a = get_arr("primaryDocument");
    let filed_a = get_arr("filingDate");
    let period_a = get_arr("reportDate");

    let mut filings = Vec::new();
    for i in 0..form_a.len() {
        if filings.len() >= count {
            break;
        }
        let form = form_a[i].as_str().unwrap_or("");
        if !wanted.iter().any(|w| w == &form.to_uppercase()) {
            continue;
        }
        let acc = acc_a.get(i).and_then(|v| v.as_str()).unwrap_or("");
        let doc = doc_a.get(i).and_then(|v| v.as_str()).unwrap_or("");
        if acc.is_empty() || doc.is_empty() {
            continue;
        }
        let acc_nodash: String = acc.chars().filter(|c| *c != '-').collect();
        filings.push(json!({
            "form": form,
            "filed": filed_a.get(i).and_then(|v| v.as_str()).unwrap_or(""),
            "period": period_a.get(i).and_then(|v| v.as_str()).unwrap_or(""),
            "url": format!(
                "https://www.sec.gov/Archives/edgar/data/{}/{}/{}",
                entry.cik, acc_nodash, doc
            ),
        }));
    }

    if filings.is_empty() {
        return json!({
            "company": entry.title,
            "ticker": entry.ticker,
            "cik": entry.cik,
            "error": format!("no recent filings matched forms {:?}", wanted),
        });
    }

    let alts = alternatives(&entries, company, &entry);
    let mut out = json!({
        "company": entry.title,
        "ticker": entry.ticker,
        "cik": entry.cik,
        "filings": filings,
        "note": "Primary-source documents — fetch_web the URL(s) you need. These are the authoritative reported figures; prefer them over aggregator summaries (rule 22).",
    });
    if !alts.is_empty() {
        if let Some(o) = out.as_object_mut() {
            o.insert("other_matches".to_string(), Value::Array(alts));
        }
    }
    out
}
