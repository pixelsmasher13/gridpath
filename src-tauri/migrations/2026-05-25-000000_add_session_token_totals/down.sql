-- SQLite supports ALTER TABLE DROP COLUMN since 3.35 (2021). All Tauri
-- targets ship a newer version, so this is safe.
ALTER TABLE spreadsheet_sessions DROP COLUMN total_input_tokens;
ALTER TABLE spreadsheet_sessions DROP COLUMN total_output_tokens;
ALTER TABLE spreadsheet_sessions DROP COLUMN total_cache_read_tokens;
ALTER TABLE spreadsheet_sessions DROP COLUMN total_cache_creation_tokens;
