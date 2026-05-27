use rusqlite::{params, Connection, named_params};
use rusqlite_from_row::FromRow;

use crate::entity::spreadsheet_session::{SpreadsheetSession, SpreadsheetSessionMessage};

/// Insert a new session, or update the name + last_opened_at if `id` already exists.
/// We use INSERT OR REPLACE to keep the call site idempotent — the frontend can
/// fire create_session on every Open without worrying about duplicates.
pub fn upsert_session(
    db: &Connection,
    id: &str,
    name: &str,
    workbook_path: &str,
) -> Result<(), rusqlite::Error> {
    db.execute(
        "INSERT INTO spreadsheet_sessions (id, name, workbook_path, created_at, updated_at, last_opened_at)
         VALUES (@id, @name, @path, datetime('now'), datetime('now'), datetime('now'))
         ON CONFLICT(id) DO UPDATE SET
            name = excluded.name,
            workbook_path = excluded.workbook_path,
            updated_at = datetime('now'),
            last_opened_at = datetime('now')",
        named_params! {
            "@id": id,
            "@name": name,
            "@path": workbook_path,
        },
    )?;
    Ok(())
}

pub fn rename_session(db: &Connection, id: &str, name: &str) -> Result<(), rusqlite::Error> {
    db.execute(
        "UPDATE spreadsheet_sessions SET name = @name, updated_at = datetime('now') WHERE id = @id",
        named_params! { "@id": id, "@name": name },
    )?;
    Ok(())
}

/// Add a batch's token counts to the running per-session totals. Called
/// from the frontend's `done` handler so cumulative usage survives the
/// React state being torn down on tab close / app restart.
pub fn add_session_tokens(
    db: &Connection,
    id: &str,
    input: i64,
    output: i64,
    cache_read: i64,
    cache_creation: i64,
) -> Result<(), rusqlite::Error> {
    db.execute(
        "UPDATE spreadsheet_sessions SET
            total_input_tokens          = total_input_tokens          + @in,
            total_output_tokens         = total_output_tokens         + @out,
            total_cache_read_tokens     = total_cache_read_tokens     + @cr,
            total_cache_creation_tokens = total_cache_creation_tokens + @cc,
            updated_at                  = datetime('now')
         WHERE id = @id",
        named_params! {
            "@id": id,
            "@in": input,
            "@out": output,
            "@cr": cache_read,
            "@cc": cache_creation,
        },
    )?;
    Ok(())
}

pub fn touch_session(db: &Connection, id: &str) -> Result<(), rusqlite::Error> {
    db.execute(
        "UPDATE spreadsheet_sessions SET updated_at = datetime('now') WHERE id = @id",
        named_params! { "@id": id },
    )?;
    Ok(())
}

pub fn archive_session(db: &Connection, id: &str) -> Result<(), rusqlite::Error> {
    db.execute(
        "UPDATE spreadsheet_sessions SET archived = 1, updated_at = datetime('now') WHERE id = @id",
        named_params! { "@id": id },
    )?;
    Ok(())
}

pub fn delete_session(db: &Connection, id: &str) -> Result<(), rusqlite::Error> {
    // Messages cascade via FK.
    db.execute(
        "DELETE FROM spreadsheet_sessions WHERE id = @id",
        named_params! { "@id": id },
    )?;
    Ok(())
}

/// Recent sessions for the sidebar. Newest updated first; archived excluded.
pub fn list_sessions(db: &Connection, limit: i64) -> Result<Vec<SpreadsheetSession>, rusqlite::Error> {
    let mut stmt = db.prepare(
        "SELECT * FROM spreadsheet_sessions
         WHERE archived = 0
         ORDER BY datetime(updated_at) DESC
         LIMIT ?1",
    )?;
    let rows = stmt.query_map(params![limit], SpreadsheetSession::try_from_row)?;
    let mut out = Vec::new();
    for r in rows {
        out.push(r?);
    }
    Ok(out)
}

#[allow(dead_code)] // Reserved for "open session by id" — sessions are currently loaded by listing.
pub fn get_session(db: &Connection, id: &str) -> Result<Option<SpreadsheetSession>, rusqlite::Error> {
    let mut stmt = db.prepare("SELECT * FROM spreadsheet_sessions WHERE id = @id")?;
    let mut rows = stmt.query(named_params! { "@id": id })?;
    if let Some(row) = rows.next()? {
        return Ok(Some(SpreadsheetSession::try_from_row(row)?));
    }
    Ok(None)
}

/// Append a message. Returns the new row id.
pub fn append_message(
    db: &Connection,
    session_id: &str,
    role: &str,
    payload: &str,
) -> Result<i64, rusqlite::Error> {
    db.execute(
        "INSERT INTO spreadsheet_session_messages (session_id, role, payload) VALUES (@sid, @role, @payload)",
        named_params! {
            "@sid": session_id,
            "@role": role,
            "@payload": payload,
        },
    )?;
    // Touching is cheap and keeps list ordering fresh.
    let _ = touch_session(db, session_id);
    Ok(db.last_insert_rowid())
}

pub fn get_messages(db: &Connection, session_id: &str) -> Result<Vec<SpreadsheetSessionMessage>, rusqlite::Error> {
    let mut stmt = db.prepare(
        "SELECT * FROM spreadsheet_session_messages WHERE session_id = @sid ORDER BY id ASC",
    )?;
    let rows = stmt.query_map(named_params! { "@sid": session_id }, SpreadsheetSessionMessage::try_from_row)?;
    let mut out = Vec::new();
    for r in rows {
        out.push(r?);
    }
    Ok(out)
}
