-- Long-term memory log: every MEMORY_SAVE entry is appended here per run, lossless.
-- Working-memory stores (engine String, memory_manager) cap at ~50K and discard oldest;
-- this table preserves the full record for completion-message generation and continuation.
CREATE TABLE IF NOT EXISTS automation_run_memory_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    execution_run_id INTEGER NOT NULL,
    content TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (execution_run_id) REFERENCES automation_execution_runs(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_run_memory_log_run ON automation_run_memory_log(execution_run_id);
CREATE INDEX IF NOT EXISTS idx_run_memory_log_run_created ON automation_run_memory_log(execution_run_id, created_at);
