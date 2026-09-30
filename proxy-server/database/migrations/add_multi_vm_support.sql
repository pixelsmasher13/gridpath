-- Multi-VM support: allow users to run multiple tasks on separate VMs concurrently

-- Track which execution is currently using a VM (prevents two tasks on the same Chrome)
ALTER TABLE vm_pool ADD COLUMN IF NOT EXISTS active_execution_run_id INTEGER;

-- Index for quick lookup of idle VMs assigned to a user
CREATE INDEX IF NOT EXISTS idx_vm_pool_active_execution ON vm_pool(assigned_user_id, active_execution_run_id)
  WHERE active_execution_run_id IS NULL;

-- Concurrency limit per user (default 8 for all users)
ALTER TABLE user_settings ADD COLUMN IF NOT EXISTS max_concurrent_vms INTEGER NOT NULL DEFAULT 8;

-- Backfill existing rows to 8
UPDATE user_settings SET max_concurrent_vms = 8 WHERE max_concurrent_vms < 8;
