-- Migration: Add takeover columns to cloud_execution_runs
-- Date: 2025-12-08
-- Description: Adds columns to support user takeover functionality for login/CAPTCHA/2FA

-- Add takeover_instructions column if it doesn't exist
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'takeover_instructions'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN takeover_instructions TEXT;

        COMMENT ON COLUMN cloud_execution_runs.takeover_instructions IS 'Instructions displayed to user during takeover (login, CAPTCHA, 2FA)';
    END IF;
END $$;

-- Add takeover_reasoning column if it doesn't exist
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'takeover_reasoning'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN takeover_reasoning TEXT;

        COMMENT ON COLUMN cloud_execution_runs.takeover_reasoning IS 'AI reasoning for why takeover is needed';
    END IF;
END $$;

-- Add takeover_started_at column if it doesn't exist
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.columns
        WHERE table_name = 'cloud_execution_runs'
        AND column_name = 'takeover_started_at'
    ) THEN
        ALTER TABLE cloud_execution_runs
        ADD COLUMN takeover_started_at TIMESTAMP WITH TIME ZONE;

        COMMENT ON COLUMN cloud_execution_runs.takeover_started_at IS 'When takeover was requested (for timeout tracking)';
    END IF;
END $$;

-- Update status CHECK constraint to include 'waiting_for_user'
-- Note: PostgreSQL doesn't have a CHECK constraint on status in the original schema,
-- so we just need to ensure the new status value is accepted (no constraint to modify)
