-- Add selected_role column to user_settings for role-based workflow suggestions
ALTER TABLE user_settings
ADD COLUMN IF NOT EXISTS selected_role VARCHAR(50);
