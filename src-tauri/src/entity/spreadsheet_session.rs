use rusqlite_from_row::FromRow;
use serde_derive::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug, FromRow, Clone)]
pub struct SpreadsheetSession {
    pub id: String,
    pub name: String,
    pub workbook_path: String,
    pub created_at: String,
    pub updated_at: String,
    pub last_opened_at: String,
    pub archived: i32,
    /// Lifetime LLM token totals across every agent batch in this session.
    /// Persisted so the Usage tab survives app restarts.
    #[serde(default)]
    pub total_input_tokens: i64,
    #[serde(default)]
    pub total_output_tokens: i64,
    #[serde(default)]
    pub total_cache_read_tokens: i64,
    #[serde(default)]
    pub total_cache_creation_tokens: i64,
}

#[derive(Serialize, Deserialize, Debug, FromRow, Clone)]
pub struct SpreadsheetSessionMessage {
    pub id: i64,
    pub session_id: String,
    pub role: String,
    pub payload: String,
    pub created_at: String,
}
