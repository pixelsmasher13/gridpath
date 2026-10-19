-- Cloud execution mode: per-automation routing between local and cloud (proxy/EC2 VM).
-- 'local'  → existing local engine (accessibility-driven on user's machine)
-- 'cloud'  → proxy /api/cloud/execute/start, runs on user's sticky VM, viewed via DCV
ALTER TABLE automations ADD COLUMN execution_target TEXT NOT NULL DEFAULT 'local';

-- When mirrored to Supabase cloud_automations, store the remote id here so subsequent
-- runs reference the same cloud automation (and therefore the same VM browser profile).
ALTER TABLE automations ADD COLUMN cloud_automation_id INTEGER;

-- Whether to persist the cloud Chrome profile after a run. Default ON so that
-- once the user logs into LinkedIn / Reddit / etc on the VM, subsequent runs reuse it.
ALTER TABLE automations ADD COLUMN save_browser_profile INTEGER NOT NULL DEFAULT 1;

-- Mirror of cloud execution runs initiated from the desktop. Lets the desktop show
-- cloud history offline and link a delegated sub-run back to its parent local run.
CREATE TABLE IF NOT EXISTS cloud_execution_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    automation_id INTEGER NOT NULL,
    proxy_execution_run_id TEXT NOT NULL,
    vm_ip TEXT,
    hostname TEXT,
    status TEXT NOT NULL,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    completion_message TEXT,
    extracted_data TEXT,
    parent_local_run_id INTEGER,
    last_seen_step_index INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_cloud_runs_automation ON cloud_execution_runs(automation_id);
CREATE INDEX IF NOT EXISTS idx_cloud_runs_proxy_id ON cloud_execution_runs(proxy_execution_run_id);
CREATE INDEX IF NOT EXISTS idx_cloud_runs_parent ON cloud_execution_runs(parent_local_run_id);
