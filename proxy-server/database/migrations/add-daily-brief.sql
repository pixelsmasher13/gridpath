-- Daily Brief Feature
-- Run this migration to add daily brief support

-- Add include_in_daily_brief flag to task_schedules
ALTER TABLE task_schedules
ADD COLUMN IF NOT EXISTS include_in_daily_brief BOOLEAN DEFAULT false;

-- Daily Brief Items table - stores items as tasks complete
CREATE TABLE IF NOT EXISTS daily_brief_items (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  brief_date DATE NOT NULL,
  execution_run_id INTEGER REFERENCES cloud_execution_runs(id) ON DELETE CASCADE,
  schedule_id INTEGER REFERENCES task_schedules(id) ON DELETE SET NULL,
  
  -- Task info (denormalized for fast reads)
  task_name VARCHAR(255) NOT NULL,
  task_objective TEXT,
  completion_message TEXT NOT NULL,
  completed_at TIMESTAMP WITH TIME ZONE NOT NULL,
  
  -- User interaction
  is_read BOOLEAN DEFAULT false,
  is_pinned BOOLEAN DEFAULT false,
  user_notes TEXT,
  
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for efficient queries
CREATE INDEX IF NOT EXISTS idx_daily_brief_user_date ON daily_brief_items(user_id, brief_date DESC);
CREATE INDEX IF NOT EXISTS idx_daily_brief_unread ON daily_brief_items(user_id, is_read) WHERE is_read = false;

-- Daily Brief Settings (add to user_settings if not using separate table)
-- If user_settings table exists, run these ALTER statements:
-- ALTER TABLE user_settings ADD COLUMN IF NOT EXISTS daily_brief_enabled BOOLEAN DEFAULT true;

-- Generated Podcasts table - stores generated podcast audio
CREATE TABLE IF NOT EXISTS daily_brief_podcasts (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  brief_date DATE NOT NULL,
  
  -- Audio file info
  audio_url TEXT,  -- URL to stored audio file
  duration_seconds INTEGER,
  
  -- Generation info
  script TEXT,  -- The podcast script that was generated
  voice_config JSONB,  -- ElevenLabs voice settings used
  
  -- Status
  status VARCHAR(50) DEFAULT 'pending',  -- pending, generating, completed, failed
  error_message TEXT,
  
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  completed_at TIMESTAMP WITH TIME ZONE,
  
  UNIQUE(user_id, brief_date)
);

CREATE INDEX IF NOT EXISTS idx_daily_brief_podcasts_user ON daily_brief_podcasts(user_id, brief_date DESC);
