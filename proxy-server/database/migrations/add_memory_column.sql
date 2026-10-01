-- Migration: Add memory column to cloud_execution_runs
-- Date: 2025-01-12
-- Description: Adds memory column to store MEMORY_SAVE data from automation executions

-- Add memory column if it doesn't exist
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'memory'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN memory TEXT;

        COMMENT ON COLUMN cloud_execution_runs.memory IS 'Data collected during execution via MEMORY_SAVE commands';
    END IF;
END $$;
