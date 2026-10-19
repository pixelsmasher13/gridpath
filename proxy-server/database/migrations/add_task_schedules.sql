-- Task Schedules for recurring automation runs
-- Run this migration to add scheduling support

-- Task Schedules table
CREATE TABLE IF NOT EXISTS task_schedules (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  automation_id INTEGER NOT NULL REFERENCES cloud_automations(id) ON DELETE CASCADE,

  -- Recurrence pattern
  recurrence_type VARCHAR(50) NOT NULL,  -- 'daily', 'weekdays', 'weekly'
  recurrence_days INTEGER[],              -- For weekly: [1,3,5] = Mon,Wed,Fri (0=Sun, 6=Sat)
  execution_hour INTEGER NOT NULL,        -- 0-23 in user's timezone
  execution_minute INTEGER DEFAULT 0,     -- 0-59
  timezone VARCHAR(100) DEFAULT 'UTC',

  -- Random offset within the hour (prevents thundering herd)
  offset_minutes INTEGER NOT NULL,        -- 0-59, randomly assigned at creation

  -- State
  is_active BOOLEAN DEFAULT true,
  next_run_at TIMESTAMP WITH TIME ZONE,   -- Pre-calculated next run time (includes offset)
  last_run_at TIMESTAMP WITH TIME ZONE,

  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for efficient queries
CREATE INDEX idx_task_schedules_user ON task_schedules(user_id);
CREATE INDEX idx_task_schedules_next_run ON task_schedules(next_run_at) WHERE is_active = true;
CREATE INDEX idx_task_schedules_automation ON task_schedules(automation_id);

-- Add schedule_id to execution runs to track scheduled executions
ALTER TABLE cloud_execution_runs
  ADD COLUMN IF NOT EXISTS schedule_id INTEGER REFERENCES task_schedules(id) ON DELETE SET NULL;

-- Update trigger for updated_at
CREATE TRIGGER update_task_schedules_updated_at
  BEFORE UPDATE ON task_schedules
  FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();
