DROP INDEX IF EXISTS idx_cloud_runs_parent;
DROP INDEX IF EXISTS idx_cloud_runs_proxy_id;
DROP INDEX IF EXISTS idx_cloud_runs_automation;
DROP TABLE IF EXISTS cloud_execution_runs;

-- SQLite doesn't support DROP COLUMN before 3.35; downgrade is no-op.
-- The new columns will simply be unused on older builds.
