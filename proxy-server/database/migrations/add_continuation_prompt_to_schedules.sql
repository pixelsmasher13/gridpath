-- Add custom continuation prompt to task_schedules
-- Allows users to override the default "Continue this task" prompt for scheduled continuations

ALTER TABLE task_schedules
ADD COLUMN IF NOT EXISTS continuation_prompt TEXT DEFAULT NULL;
