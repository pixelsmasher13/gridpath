-- Add persistent_run_id to task_schedules
-- Allows schedules to continue from a specific run instead of fresh each time

ALTER TABLE task_schedules
ADD COLUMN IF NOT EXISTS persistent_run_id INTEGER REFERENCES cloud_execution_runs(id) ON DELETE SET NULL;
