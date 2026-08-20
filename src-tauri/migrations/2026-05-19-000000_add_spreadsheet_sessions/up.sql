-- Spreadsheet workspace sessions: one per "agent + workbook" pair the user opens.
-- The workbook itself stays on disk (path is canonical); this table only stores
-- session metadata + the in-app name. Closing/reopening the app restores tabs
-- from here and replays messages.
CREATE TABLE IF NOT EXISTS spreadsheet_sessions (
    id TEXT PRIMARY KEY,                -- uuid v4 from the frontend
    name TEXT NOT NULL,                 -- human-readable session name (auto-gen on first prompt)
    workbook_path TEXT NOT NULL,        -- absolute path to the xlsx file
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now')),
    last_opened_at TEXT NOT NULL DEFAULT (datetime('now')),
    archived INTEGER NOT NULL DEFAULT 0 -- 0/1; archived sessions hidden from the default list
);

CREATE INDEX IF NOT EXISTS idx_spreadsheet_sessions_workbook
    ON spreadsheet_sessions(workbook_path);

CREATE INDEX IF NOT EXISTS idx_spreadsheet_sessions_updated
    ON spreadsheet_sessions(updated_at DESC);

-- Append-only conversation log for each session: user prompts, agent text deltas
-- (compacted to one row per turn), and accepted/rejected batches.
-- The `payload` column is a JSON blob keyed by `role`:
--   role = "user"           -> { "prompt": "..." }
--   role = "agent_text"     -> { "text": "..." }
--   role = "agent_batch"    -> { "batch": { id, prompt, justification, mutations, status, created_at } }
-- Keeping it as one table with a discriminator avoids a 4-table join when we
-- load a session's history.
CREATE TABLE IF NOT EXISTS spreadsheet_session_messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL,
    role TEXT NOT NULL,                 -- "user" | "agent_text" | "agent_batch"
    payload TEXT NOT NULL,              -- JSON
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (session_id) REFERENCES spreadsheet_sessions(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_spreadsheet_session_messages_session
    ON spreadsheet_session_messages(session_id, id);
