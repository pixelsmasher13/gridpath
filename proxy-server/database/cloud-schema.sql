-- Cloud VM Management Schema for PostgreSQL
-- Supports VM pool, sessions, and real-time communication

-- VM Pool Management
CREATE TABLE IF NOT EXISTS vm_pool (
  instance_id VARCHAR(255) PRIMARY KEY,
  ip_address VARCHAR(50) NOT NULL,
  private_ip VARCHAR(50),
  hostname VARCHAR(255),  -- DNS hostname for SSL access (e.g., abc123.vm.linefox.ai)
  state VARCHAR(50) NOT NULL DEFAULT 'pending', -- pending, running, stopping, stopped, terminated
  health_status VARCHAR(50) DEFAULT 'unknown', -- healthy, unhealthy, unknown
  pool_status VARCHAR(50) DEFAULT 'available', -- available, assigned, maintenance, terminating
  assigned_user_id VARCHAR(255),
  assigned_session_id VARCHAR(255),
  last_health_check TIMESTAMP WITH TIME ZONE,
  last_used_at TIMESTAMP WITH TIME ZONE,
  launch_time TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Indexes for pool queries
CREATE INDEX idx_vm_pool_status ON vm_pool(pool_status, state);
CREATE INDEX idx_vm_pool_assigned ON vm_pool(assigned_user_id);
CREATE INDEX idx_vm_pool_health ON vm_pool(health_status);

-- Recording Sessions
CREATE TABLE IF NOT EXISTS cloud_recording_sessions (
  id SERIAL PRIMARY KEY,
  session_id VARCHAR(255) UNIQUE NOT NULL,
  user_id VARCHAR(255) NOT NULL,
  vm_instance_id VARCHAR(255),
  automation_name VARCHAR(255) NOT NULL,
  objective TEXT,
  status VARCHAR(50) DEFAULT 'initializing', -- initializing, connected, recording, processing, completed, failed
  websocket_connected BOOLEAN DEFAULT FALSE,
  vnc_port INTEGER DEFAULT 5900,
  vnc_password VARCHAR(255),
  started_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  ended_at TIMESTAMP WITH TIME ZONE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,

  FOREIGN KEY (vm_instance_id) REFERENCES vm_pool(instance_id) ON DELETE SET NULL
);

-- Activity Logs (raw events captured from NVDA on VM during recording)
CREATE TABLE IF NOT EXISTS cloud_activity_logs (
  id SERIAL PRIMARY KEY,
  session_id VARCHAR(255) NOT NULL,
  user_id VARCHAR(255) NOT NULL,
  timestamp TIMESTAMP WITH TIME ZONE NOT NULL,
  window_title TEXT,
  url TEXT,  -- Current page URL in Chrome
  detected_action JSONB NOT NULL,  -- Raw NVDA event data
  action_type VARCHAR(100),  -- click, type, navigate, scroll, etc.
  element_text TEXT,  -- Text of the element interacted with
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,

  FOREIGN KEY (session_id) REFERENCES cloud_recording_sessions(session_id) ON DELETE CASCADE
);

CREATE INDEX idx_activity_logs_session ON cloud_activity_logs(session_id);
CREATE INDEX idx_activity_logs_timestamp ON cloud_activity_logs(session_id, timestamp);

-- Automation Definitions (aligned with desktop, but for VM execution)
CREATE TABLE IF NOT EXISTS cloud_automations (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  name VARCHAR(255) NOT NULL,
  objective TEXT NOT NULL,
  generalized_script JSONB,  -- Generalized/parameterized version for VM execution
  run_count INTEGER DEFAULT 0,
  last_run_at TIMESTAMP WITH TIME ZONE,
  is_active BOOLEAN DEFAULT TRUE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

-- Execution Runs (tracks automation runs on VMs)
CREATE TABLE IF NOT EXISTS cloud_execution_runs (
  id SERIAL PRIMARY KEY,
  job_id VARCHAR(255) UNIQUE NOT NULL,
  automation_id INTEGER NOT NULL,
  user_id VARCHAR(255) NOT NULL,
  vm_instance_id VARCHAR(255),
  vm_ip VARCHAR(45),  -- Store VM IP directly for reliable DCV reconnection
  vm_hostname VARCHAR(255),  -- Store hostname for SSL connections
  status VARCHAR(50) DEFAULT 'queued', -- queued, starting, running, completed, failed, stopped
  additional_instructions TEXT,  -- User can provide extra instructions for NVDA on VM
  error_message TEXT,
  memory TEXT,  -- Data collected during execution via MEMORY_SAVE commands
  started_at TIMESTAMP WITH TIME ZONE,
  completed_at TIMESTAMP WITH TIME ZONE,
  duration_seconds INTEGER,                            -- Active duration (sum across turns); replaces old wall-clock value
  total_active_seconds INTEGER DEFAULT 0,              -- Authoritative active-time accumulator (mirrors duration_seconds)
  current_turn_started_at TIMESTAMP WITH TIME ZONE,    -- Start of the currently in-progress turn; NULL between turns
  last_activity_at TIMESTAMP WITH TIME ZONE GENERATED ALWAYS AS (
    GREATEST(started_at, current_turn_started_at, completed_at)
  ) STORED,                                            -- Stable sort key; updates only on real start/continue/end boundaries
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,

  FOREIGN KEY (automation_id) REFERENCES cloud_automations(id) ON DELETE CASCADE,
  FOREIGN KEY (vm_instance_id) REFERENCES vm_pool(instance_id) ON DELETE SET NULL
);

-- Execution Steps (NVDA's step-by-step execution on VM)
CREATE TABLE IF NOT EXISTS cloud_execution_steps (
  id SERIAL PRIMARY KEY,
  execution_run_id INTEGER NOT NULL,
  step_number INTEGER NOT NULL,
  timestamp TIMESTAMP WITH TIME ZONE NOT NULL,
  explanation TEXT,  -- NVDA's explanation of what it's doing on the VM
  next_step TEXT,  -- NVDA's plan for the next action on the VM
  status VARCHAR(50) DEFAULT 'executing', -- pending, executing, completed, error
  error_message TEXT,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,

  FOREIGN KEY (execution_run_id) REFERENCES cloud_execution_runs(id) ON DELETE CASCADE
);

-- Note: VMs are controlled via direct HTTP calls to their private IPs
-- No command queue table needed - using SQS for job queue instead

-- Note: WebSocket connections not used in current architecture
-- VNC provides visual control, HTTP provides API communication

-- Chrome Profiles Storage
CREATE TABLE IF NOT EXISTS chrome_profiles (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  profile_name VARCHAR(255) DEFAULT 'default',
  s3_key VARCHAR(500),
  size_bytes BIGINT,
  last_sync TIMESTAMP WITH TIME ZONE,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  updated_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,

  CONSTRAINT unique_user_profile UNIQUE (user_id, profile_name)
);

-- VM Pre-warm Cache (for paid users)
CREATE TABLE IF NOT EXISTS vm_prewarm_cache (
  user_id VARCHAR(255) PRIMARY KEY,
  instance_id VARCHAR(255) NOT NULL,
  reserved_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
  expires_at TIMESTAMP WITH TIME ZONE NOT NULL,

  FOREIGN KEY (instance_id) REFERENCES vm_pool(instance_id) ON DELETE CASCADE
);

-- Note: VM pool configuration is managed via environment variables
-- See .env.cloud for configuration settings:
-- - MAX_POOL_SIZE: Maximum VMs allowed
-- - AMI_ID, INSTANCE_TYPE, SUBNET_ID, SECURITY_GROUP_ID: AWS resources
-- Other settings (idle timeout, health check interval) are hardcoded in vm-pool-manager.ts

-- Usage Logs (for billing)
CREATE TABLE IF NOT EXISTS cloud_usage_logs (
  id SERIAL PRIMARY KEY,
  user_id VARCHAR(255) NOT NULL,
  session_id VARCHAR(255),
  session_type VARCHAR(50), -- recording, execution
  vm_instance_id VARCHAR(255),
  started_at TIMESTAMP WITH TIME ZONE NOT NULL,
  ended_at TIMESTAMP WITH TIME ZONE,
  duration_seconds INTEGER,
  billed_minutes INTEGER,
  created_at TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX idx_usage_logs_user ON cloud_usage_logs(user_id, started_at DESC);

-- Update triggers
CREATE OR REPLACE FUNCTION update_updated_at_column()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = CURRENT_TIMESTAMP;
    RETURN NEW;
END;
$$ language 'plpgsql';

CREATE TRIGGER update_vm_pool_updated_at BEFORE UPDATE ON vm_pool
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_recording_sessions_updated_at BEFORE UPDATE ON cloud_recording_sessions
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_automations_updated_at BEFORE UPDATE ON cloud_automations
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_execution_runs_updated_at BEFORE UPDATE ON cloud_execution_runs
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();

CREATE TRIGGER update_chrome_profiles_updated_at BEFORE UPDATE ON chrome_profiles
    FOR EACH ROW EXECUTE FUNCTION update_updated_at_column();