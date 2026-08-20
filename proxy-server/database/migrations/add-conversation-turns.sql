-- Migration: Add conversation turn tracking for chat-like UI
-- This enables grouping steps by conversation turn and distinguishing user prompts from agent actions

-- Add turn_number to track which conversation turn a step belongs to
-- Turn 0 = initial execution, Turn 1+ = continuations
ALTER TABLE cloud_execution_steps 
ADD COLUMN IF NOT EXISTS turn_number INTEGER DEFAULT 0;

-- Add step_type to distinguish between different step types:
-- 'agent_action' = normal execution step (default)
-- 'user_prompt' = user's follow-up/continuation prompt
-- 'completion' = AI's completion message for this turn
ALTER TABLE cloud_execution_steps 
ADD COLUMN IF NOT EXISTS step_type VARCHAR(50) DEFAULT 'agent_action';

-- Add updated_at for tracking
ALTER TABLE cloud_execution_steps 
ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP;

-- Add current_turn to execution runs to track which turn we're on
ALTER TABLE cloud_execution_runs 
ADD COLUMN IF NOT EXISTS current_turn INTEGER DEFAULT 0;

-- Index for efficient querying by turn
CREATE INDEX IF NOT EXISTS idx_execution_steps_turn 
ON cloud_execution_steps(execution_run_id, turn_number);

-- Index for finding completion messages
CREATE INDEX IF NOT EXISTS idx_execution_steps_type 
ON cloud_execution_steps(execution_run_id, step_type);

-- Comment for documentation
COMMENT ON COLUMN cloud_execution_steps.turn_number IS 'Conversation turn: 0=initial, 1+=continuations';
COMMENT ON COLUMN cloud_execution_steps.step_type IS 'Type: agent_action, user_prompt, or completion';
COMMENT ON COLUMN cloud_execution_runs.current_turn IS 'Current conversation turn number';
