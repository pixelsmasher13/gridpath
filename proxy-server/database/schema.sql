-- Usage tracking database schema for PostgreSQL (recommended for Vercel)
-- You can also adapt this for MySQL if preferred

-- Create usage_sessions table
CREATE TABLE IF NOT EXISTS usage_sessions (
  id VARCHAR(255) PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  automation_id INTEGER NOT NULL,
  automation_name VARCHAR(255) NOT NULL,
  started_at TIMESTAMP WITH TIME ZONE NOT NULL,
  ended_at TIMESTAMP WITH TIME ZONE,
  total_seconds INTEGER DEFAULT 0,
  billed_minutes INTEGER DEFAULT 0,
  status VARCHAR(50) DEFAULT 'active' CHECK (status IN ('active', 'completed', 'failed')),
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Create indexes for performance
CREATE INDEX IF NOT EXISTS idx_usage_sessions_user_id ON usage_sessions(user_id);
CREATE INDEX IF NOT EXISTS idx_usage_sessions_started_at ON usage_sessions(started_at);
CREATE INDEX IF NOT EXISTS idx_usage_sessions_status ON usage_sessions(status);
CREATE INDEX IF NOT EXISTS idx_usage_sessions_user_started ON usage_sessions(user_id, started_at DESC);

-- Create usage_aggregates table for faster queries
CREATE TABLE IF NOT EXISTS usage_aggregates (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  date DATE NOT NULL,
  automation_id INTEGER,
  automation_name VARCHAR(255),
  total_minutes INTEGER DEFAULT 0,
  total_sessions INTEGER DEFAULT 0,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  
  CONSTRAINT unique_user_date_automation UNIQUE (user_id, date, automation_id)
);

-- Create indexes for aggregates
CREATE INDEX IF NOT EXISTS idx_usage_aggregates_user_date ON usage_aggregates(user_id, date DESC);
CREATE INDEX IF NOT EXISTS idx_usage_aggregates_date ON usage_aggregates(date);

-- Function to update the updated_at timestamp
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ language 'plpgsql';

-- Create triggers for updated_at
DROP TRIGGER IF EXISTS update_usage_sessions_updated_at ON usage_sessions;
CREATE TRIGGER update_usage_sessions_updated_at BEFORE UPDATE
    ON usage_sessions FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

DROP TRIGGER IF EXISTS update_usage_aggregates_updated_at ON usage_aggregates;
CREATE TRIGGER update_usage_aggregates_updated_at BEFORE UPDATE
    ON usage_aggregates FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- Function to aggregate usage data (can be called periodically)
CREATE OR REPLACE FUNCTION aggregate_usage_data(target_date DATE DEFAULT CURRENT_DATE)
RETURNS void AS $$
BEGIN
    INSERT INTO usage_aggregates (user_id, date, automation_id, automation_name, total_minutes, total_sessions)
    SELECT 
        user_id,
        DATE(started_at) as usage_date,
        automation_id,
        automation_name,
        SUM(billed_minutes) as total_minutes,
        COUNT(*) as total_sessions
    FROM usage_sessions
    WHERE DATE(started_at) = target_date
        AND status IN ('completed', 'failed')
    GROUP BY user_id, DATE(started_at), automation_id, automation_name
    ON CONFLICT (user_id, date, automation_id) 
    DO UPDATE SET
        total_minutes = EXCLUDED.total_minutes,
        total_sessions = EXCLUDED.total_sessions,
        updated_at = CURRENT_TIMESTAMP;
END;
$$ LANGUAGE plpgsql;

-- Create shared_automations table for sharing automation scripts
CREATE TABLE IF NOT EXISTS shared_automations (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  share_name VARCHAR(255) NOT NULL,
  automation_name VARCHAR(255) NOT NULL,
  objective TEXT,
  nl_description TEXT,
  generalized_script TEXT NOT NULL,
  category VARCHAR(50) DEFAULT 'general' CHECK (category IN ('general', 'research', 'finance', 'productivity', 'development', 'data', 'communication', 'entertainment', 'utilities')),
  visibility VARCHAR(50) DEFAULT 'private' CHECK (visibility IN ('private', 'public')),
  share_token VARCHAR(255) UNIQUE DEFAULT gen_random_uuid()::text,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Create indexes for shared automations
CREATE INDEX IF NOT EXISTS idx_shared_automations_user_id ON shared_automations(user_id);
CREATE INDEX IF NOT EXISTS idx_shared_automations_visibility ON shared_automations(visibility);
CREATE INDEX IF NOT EXISTS idx_shared_automations_category ON shared_automations(category);
CREATE INDEX IF NOT EXISTS idx_shared_automations_share_token ON shared_automations(share_token);
CREATE INDEX IF NOT EXISTS idx_shared_automations_created_at ON shared_automations(created_at DESC);

-- Add trigger for updated_at on shared_automations
DROP TRIGGER IF EXISTS update_shared_automations_updated_at ON shared_automations;
CREATE TRIGGER update_shared_automations_updated_at BEFORE UPDATE
    ON shared_automations FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- Create automation_ratings table for user ratings
CREATE TABLE IF NOT EXISTS automation_ratings (
  id SERIAL PRIMARY KEY,
  automation_id INTEGER NOT NULL REFERENCES shared_automations(id) ON DELETE CASCADE,
  user_id VARCHAR(255) NOT NULL,
  rating INTEGER NOT NULL CHECK (rating >= 1 AND rating <= 5),
  review TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  
  CONSTRAINT unique_user_automation_rating UNIQUE (automation_id, user_id)
);

-- Create indexes for ratings
CREATE INDEX IF NOT EXISTS idx_automation_ratings_automation_id ON automation_ratings(automation_id);
CREATE INDEX IF NOT EXISTS idx_automation_ratings_user_id ON automation_ratings(user_id);

-- Add trigger for updated_at on automation_ratings
DROP TRIGGER IF EXISTS update_automation_ratings_updated_at ON automation_ratings;
CREATE TRIGGER update_automation_ratings_updated_at BEFORE UPDATE
    ON automation_ratings FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

-- Add computed fields to shared_automations for ratings
ALTER TABLE shared_automations 
ADD COLUMN IF NOT EXISTS average_rating DECIMAL(3,2) DEFAULT 0,
ADD COLUMN IF NOT EXISTS total_ratings INTEGER DEFAULT 0,
ADD COLUMN IF NOT EXISTS total_uses INTEGER DEFAULT 0;

-- Function to update automation rating statistics
CREATE OR REPLACE FUNCTION update_automation_rating_stats()
RETURNS TRIGGER AS $$
BEGIN
    UPDATE shared_automations
    SET average_rating = (
        SELECT AVG(rating)::DECIMAL(3,2)
        FROM automation_ratings
        WHERE automation_id = COALESCE(NEW.automation_id, OLD.automation_id)
    ),
    total_ratings = (
        SELECT COUNT(*)
        FROM automation_ratings
        WHERE automation_id = COALESCE(NEW.automation_id, OLD.automation_id)
    )
    WHERE id = COALESCE(NEW.automation_id, OLD.automation_id);
    
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Create triggers to update rating stats
DROP TRIGGER IF EXISTS update_rating_stats_on_insert ON automation_ratings;
CREATE TRIGGER update_rating_stats_on_insert
    AFTER INSERT ON automation_ratings
    FOR EACH ROW EXECUTE FUNCTION update_automation_rating_stats();

DROP TRIGGER IF EXISTS update_rating_stats_on_update ON automation_ratings;
CREATE TRIGGER update_rating_stats_on_update
    AFTER UPDATE ON automation_ratings
    FOR EACH ROW EXECUTE FUNCTION update_automation_rating_stats();

DROP TRIGGER IF EXISTS update_rating_stats_on_delete ON automation_ratings;
CREATE TRIGGER update_rating_stats_on_delete
    AFTER DELETE ON automation_ratings
    FOR EACH ROW EXECUTE FUNCTION update_automation_rating_stats();

-- Migration: Add category column to existing tables (safe to run multiple times)
DO $$ 
BEGIN
    IF NOT EXISTS (SELECT 1 FROM information_schema.columns 
                  WHERE table_name = 'shared_automations' 
                  AND column_name = 'category') THEN
        ALTER TABLE shared_automations 
        ADD COLUMN category VARCHAR(50) DEFAULT 'general' 
        CHECK (category IN ('general', 'research', 'finance', 'productivity', 'development', 'data', 'communication', 'entertainment', 'utilities'));
    END IF;
END $$;