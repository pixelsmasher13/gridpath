use rusqlite::{Connection, params};
use serde::{Serialize, Deserialize};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ClarificationEntry {
    pub id: i64,
    pub automation_id: Option<i64>,
    pub question: String,
    pub answer: Option<String>,
    pub status: String, // "pending" or "answered"
    pub created_at: String,
    pub answered_at: Option<String>,
}

pub fn init_table(db: &Connection) {
    db.execute_batch(
        "CREATE TABLE IF NOT EXISTS pending_clarifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            automation_id INTEGER,
            question TEXT NOT NULL,
            answer TEXT,
            status TEXT NOT NULL DEFAULT 'pending',
            created_at DATETIME DEFAULT (datetime('now')),
            answered_at DATETIME
        );",
    ).expect("Failed to create pending_clarifications table");
}

pub fn insert_question(db: &Connection, automation_id: Option<i64>, question: &str) -> rusqlite::Result<i64> {
    db.execute(
        "INSERT INTO pending_clarifications (automation_id, question) VALUES (?, ?)",
        params![automation_id, question],
    )?;
    Ok(db.last_insert_rowid())
}

pub fn update_answer(db: &Connection, id: i64, answer: &str) -> rusqlite::Result<()> {
    db.execute(
        "UPDATE pending_clarifications SET answer = ?, status = 'answered', answered_at = datetime('now') WHERE id = ?",
        params![answer, id],
    )?;
    Ok(())
}

pub fn get_by_automation(db: &Connection, automation_id: i64) -> rusqlite::Result<Vec<ClarificationEntry>> {
    let mut stmt = db.prepare(
        "SELECT id, automation_id, question, answer, status, created_at, answered_at FROM pending_clarifications WHERE automation_id = ? ORDER BY created_at ASC",
    )?;
    let rows = stmt.query_map(params![automation_id], |row| {
        Ok(ClarificationEntry {
            id: row.get(0)?,
            automation_id: row.get(1)?,
            question: row.get(2)?,
            answer: row.get(3)?,
            status: row.get(4)?,
            created_at: row.get(5)?,
            answered_at: row.get(6)?,
        })
    })?;
    Ok(rows.filter_map(Result::ok).collect())
}

pub fn get_answered_by_automation(db: &Connection, automation_id: i64) -> rusqlite::Result<Vec<ClarificationEntry>> {
    let mut stmt = db.prepare(
        "SELECT id, automation_id, question, answer, status, created_at, answered_at FROM pending_clarifications WHERE automation_id = ? AND status = 'answered' ORDER BY created_at ASC",
    )?;
    let rows = stmt.query_map(params![automation_id], |row| {
        Ok(ClarificationEntry {
            id: row.get(0)?,
            automation_id: row.get(1)?,
            question: row.get(2)?,
            answer: row.get(3)?,
            status: row.get(4)?,
            created_at: row.get(5)?,
            answered_at: row.get(6)?,
        })
    })?;
    Ok(rows.filter_map(Result::ok).collect())
} 