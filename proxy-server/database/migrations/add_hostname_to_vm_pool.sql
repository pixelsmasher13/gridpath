-- Add hostname column to vm_pool for SSL access via DNS
-- Hostname format: {instance_id}.vm.linefox.ai

ALTER TABLE vm_pool ADD COLUMN IF NOT EXISTS hostname VARCHAR(255);

-- Optional: Add comment for documentation
COMMENT ON COLUMN vm_pool.hostname IS 'DNS hostname for SSL access (e.g., abc123.vm.linefox.ai)';
