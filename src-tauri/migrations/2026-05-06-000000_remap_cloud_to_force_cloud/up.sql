-- Phase 3 architecture flip: the local planner is now always the orchestrator,
-- and cloud is reachable as a per-action tool (DELEGATE_TO_CLOUD). The legacy
-- 'cloud' value used to mean "skip the planner entirely" — that semantic now
-- lives under 'force_cloud'. Anything else falls through to the planner.
--
-- Existing rows tagged 'cloud' presumably wanted the skip-planner behavior, so
-- map them to 'force_cloud' to preserve user intent.
UPDATE automations SET execution_target = 'force_cloud' WHERE execution_target = 'cloud';
