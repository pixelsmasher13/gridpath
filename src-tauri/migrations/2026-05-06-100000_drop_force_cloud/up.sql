-- The local planner is now always the orchestrator and chooses cloud delegation
-- per-action. There is no longer a "skip the planner" execution mode.
-- Map any leftover 'force_cloud' rows back to 'local' so they run through the
-- planner like everything else.
UPDATE automations SET execution_target = 'local' WHERE execution_target = 'force_cloud';
