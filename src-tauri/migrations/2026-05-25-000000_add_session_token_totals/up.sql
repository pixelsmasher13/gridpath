-- Per-session lifetime LLM token totals so the Usage tab shows cumulative
-- spend across app restarts. Previously these counts lived only in React
-- state on the open tab and were lost on close. Each batch's `done` event
-- now bumps these via session_add_tokens.
ALTER TABLE spreadsheet_sessions ADD COLUMN total_input_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE spreadsheet_sessions ADD COLUMN total_output_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE spreadsheet_sessions ADD COLUMN total_cache_read_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE spreadsheet_sessions ADD COLUMN total_cache_creation_tokens INTEGER NOT NULL DEFAULT 0;
