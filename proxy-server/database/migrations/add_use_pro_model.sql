-- Migration: Add use_pro_model column to user_settings table
-- When true, uses gemini-pro model at 1.5x minute consumption rate

ALTER TABLE user_settings
ADD COLUMN IF NOT EXISTS use_pro_model BOOLEAN DEFAULT FALSE;

COMMENT ON COLUMN user_settings.use_pro_model IS 'When true, uses gemini-pro model at 1.5x minute consumption rate';
