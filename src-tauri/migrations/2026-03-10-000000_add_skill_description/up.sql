-- No-op: the `description` column is already created by the initial schema
-- migration (2024-10-10-232810_add_tables, user_skills definition). This
-- migration is kept as a no-op so Diesel's migration tracker on existing
-- installs (where it was historically applied) stays consistent.
SELECT 1;
