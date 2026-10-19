-- Migration: Add nl_description column to cloud_automations table
-- This column stores the natural language description/script generated from user prompts
-- Used by the execution engine to guide automation steps

ALTER TABLE cloud_automations
ADD COLUMN IF NOT EXISTS nl_description TEXT;

-- Add comment for documentation
COMMENT ON COLUMN cloud_automations.nl_description IS 'Natural language step-by-step script generated from user prompt, used by LLM during execution';
