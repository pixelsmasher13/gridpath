ALTER TABLE usage_sessions DROP COLUMN total_input_tokens;
ALTER TABLE usage_sessions DROP COLUMN total_output_tokens;
ALTER TABLE usage_sessions DROP COLUMN cache_read_tokens;
ALTER TABLE usage_sessions DROP COLUMN cache_creation_tokens;
ALTER TABLE usage_sessions DROP COLUMN api_calls;
ALTER TABLE usage_sessions DROP COLUMN provider;
