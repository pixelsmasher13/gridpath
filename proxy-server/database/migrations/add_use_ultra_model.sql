-- Migration: Add use_ultra_model column to user_settings table
-- When true, the cloud execution path routes LLM calls to Claude Opus 4.7
-- and bills minutes at 3x the base rate.

ALTER TABLE user_settings
ADD COLUMN IF NOT EXISTS use_ultra_model BOOLEAN DEFAULT FALSE;

COMMENT ON COLUMN user_settings.use_ultra_model IS 'When true, uses Claude Opus 4.7 (Ultra tier) at 3x minute consumption rate. Mutually exclusive with use_pro_model.';
