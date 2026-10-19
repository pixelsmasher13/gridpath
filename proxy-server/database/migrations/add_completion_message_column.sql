-- Migration: Add completion_message column to cloud_execution_runs
-- Date: 2025-01-24
-- Description: Adds completion_message column to store LLM-generated closing message summarizing workflow results

-- Add completion_message column if it doesn't exist
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'completion_message'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN completion_message TEXT;

        COMMENT ON COLUMN cloud_execution_runs.completion_message IS 'LLM-generated closing message summarizing what was accomplished during the workflow execution';
    END IF;
END $$;
