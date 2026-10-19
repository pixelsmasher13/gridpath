-- Migration: Add long_term_memory column to cloud_execution_runs
-- Date: 2026-04-27
-- Description:
--   Splits MEMORY_SAVE storage into two tiers so long executions stop dropping data:
--     • memory            — bounded working set shown on every step (~45K cap, unchanged)
--     • long_term_memory  — append-only archive of raw entries displaced from working set,
--                           sent ONLY to the completion prompt (never per-step)
--   When working memory crosses the soft cap, the oldest entries are appended verbatim to
--   long_term_memory and replaced in working memory by a short manifest summary line.

DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'long_term_memory'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN long_term_memory TEXT;

        COMMENT ON COLUMN cloud_execution_runs.long_term_memory IS
            'Append-only archive of raw MEMORY_SAVE entries displaced from working memory. Used by completion prompt; never sent in per-step decision calls.';
    END IF;
END $$;
