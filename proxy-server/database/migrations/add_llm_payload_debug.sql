-- Migration: Debug table for capturing raw LLM input payloads per turn
-- Enable by setting LOG_LLM_PAYLOAD=1 in the proxy-server environment.
-- Disable in normal prod to avoid storage bloat (each row can be 30-100KB).

CREATE TABLE IF NOT EXISTS llm_payload_debug (
  id SERIAL PRIMARY KEY,
  execution_run_id INTEGER NOT NULL,
  step_number INTEGER NOT NULL,
  model VARCHAR(64),
  user_prompt TEXT NOT NULL,
  user_prompt_chars INTEGER NOT NULL,
  est_memory_tokens INTEGER,
  est_history_tokens INTEGER,
  est_elements_tokens INTEGER,
  est_skill_tokens INTEGER,
  est_user_total_tokens INTEGER,
  elements_count INTEGER,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_llm_payload_debug_run_step
  ON llm_payload_debug (execution_run_id, step_number);

-- Lock down PostgREST (Supabase anon/authenticated keys). The proxy-server uses
-- a direct pg Pool with DATABASE_URL, which bypasses RLS, so this is purely a
-- defense-in-depth seal against accidental API exposure of debug payloads.
ALTER TABLE llm_payload_debug ENABLE ROW LEVEL SECURITY;
