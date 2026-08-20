-- Reference workbooks attached to a session as read-only agent context
-- (other xlsx files, e.g. analyst models to compare against). Stored as a
-- JSON array of absolute paths so restoring a session restores its chips.
ALTER TABLE spreadsheet_sessions ADD COLUMN reference_paths TEXT NOT NULL DEFAULT '[]';
