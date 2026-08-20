-- Migration: Per-turn duration & billing tracking for cloud_execution_runs
--
-- Background:
--   Previously, `duration_seconds` was computed as `(completed_at - started_at)`,
--   i.e. wall time from the very first start to the latest completion. For runs
--   that were continued days later, this counted dormant time as "duration",
--   producing values like 17 days for a task that did 5 minutes of real work.
--
--   Billing also keyed on a per-run session id (`cloud_execution_${runId}`), so the
--   process_usage_session_credits dedupe path silently skipped every continuation
--   after turn 0. Going forward we use a per-turn session id and only bill the
--   active seconds within that turn.
--
-- This migration adds:
--   * total_active_seconds      - running sum of active time across all turns
--   * current_turn_started_at   - wall-clock start of the currently active turn
--                                 (NULL when no turn is in progress / between
--                                 continuations). Used to compute the active
--                                 delta when a turn ends.

ALTER TABLE cloud_execution_runs
  ADD COLUMN IF NOT EXISTS total_active_seconds INTEGER DEFAULT 0;

ALTER TABLE cloud_execution_runs
  ADD COLUMN IF NOT EXISTS current_turn_started_at TIMESTAMP WITH TIME ZONE;

-- Best-effort backfill so historical rows aren't reset to 0:
-- use the existing (wall-clock-inflated) duration_seconds as a fallback.
-- Going forward, processBilling will overwrite duration_seconds with the
-- accurate total_active_seconds value.
UPDATE cloud_execution_runs
SET total_active_seconds = COALESCE(duration_seconds, 0)
WHERE total_active_seconds = 0
  AND duration_seconds IS NOT NULL;

-- Stable sort key for the dashboard. Sorting by updated_at was too noisy
-- (every memory write, status flip, vm reassignment, etc. trips the
-- update_updated_at_column trigger and reorders the list). Sorting by
-- started_at buried continued tasks. last_activity_at = the latest of:
--   * started_at              (initial creation)
--   * current_turn_started_at (start of the in-flight turn, if any)
--   * completed_at            (end of the most recent turn)
-- so a run only moves up when something genuinely happens to it. Postgres's
-- GREATEST ignores NULL args, so completed runs use completed_at, in-flight
-- runs use current_turn_started_at, etc.
ALTER TABLE cloud_execution_runs
  ADD COLUMN IF NOT EXISTS last_activity_at TIMESTAMP WITH TIME ZONE
  GENERATED ALWAYS AS (
    GREATEST(started_at, current_turn_started_at, completed_at)
  ) STORED;

CREATE INDEX IF NOT EXISTS idx_execution_runs_last_activity
  ON cloud_execution_runs (user_id, last_activity_at DESC);

COMMENT ON COLUMN cloud_execution_runs.total_active_seconds IS
  'Sum of active execution time across all turns (seconds). Replaces wall-clock duration_seconds going forward.';

COMMENT ON COLUMN cloud_execution_runs.current_turn_started_at IS
  'Wall-clock start of the currently active turn. NULL when no turn is in progress (e.g. between continuations or after completion).';

COMMENT ON COLUMN cloud_execution_runs.last_activity_at IS
  'Generated stable sort key: GREATEST(started_at, current_turn_started_at, completed_at). Updates only on real boundaries (start, continuation, turn end), not on memory/status churn.';
