ALTER TABLE usage_sessions ADD COLUMN total_input_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE usage_sessions ADD COLUMN total_output_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE usage_sessions ADD COLUMN cache_read_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE usage_sessions ADD COLUMN cache_creation_tokens INTEGER NOT NULL DEFAULT 0;
ALTER TABLE usage_sessions ADD COLUMN api_calls INTEGER NOT NULL DEFAULT 0;
ALTER TABLE usage_sessions ADD COLUMN provider TEXT NOT NULL DEFAULT '';
