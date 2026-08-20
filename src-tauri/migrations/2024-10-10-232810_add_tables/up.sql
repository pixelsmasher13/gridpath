-- Consolidated schema: all tables with full column definitions

CREATE TABLE IF NOT EXISTS activity_logs (
    timestamp TEXT NOT NULL DEFAULT '',
    detected_actions TEXT NOT NULL DEFAULT '',
    element_tree_dump TEXT NOT NULL DEFAULT '',
    editing_mode TEXT,
    ocr_text TEXT NOT NULL DEFAULT '',
    full_activity_text TEXT,
    original_ocr_text TEXT,
    os_details TEXT NOT NULL DEFAULT '',
    user_id TEXT NOT NULL DEFAULT '',
    window_title TEXT NOT NULL DEFAULT '',
    window_app_name TEXT NOT NULL DEFAULT '',
    similarity_percentage_to_previous_ocr_text TEXT NOT NULL DEFAULT '',
    interval_length INTEGER,
    keypress_count INTEGER,
    action_insights TEXT,
    automation_id INTEGER
);

CREATE INDEX IF NOT EXISTS idx_activity_logs_automation_id ON activity_logs(automation_id);

CREATE TABLE IF NOT EXISTS settings (
    setting_key TEXT NOT NULL PRIMARY KEY DEFAULT '',
    setting_value TEXT NOT NULL DEFAULT ''
);

INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES
    ('api_choice', 'proxy'),
    ('interval', '10'),
    ('is_dev_mode', 'false'),
    ('auto_start', 'false'),
    ('api_key_claude', ''),
    ('api_key_open_ai', ''),
    ('api_key_grok', ''),
    ('api_key_gemini', ''),
    ('model_claude',  'claude-sonnet-4-5-20250929'),
    ('model_openai',  'gpt-5'),
    ('model_grok',    'grok-3'),
    ('model_gemini',  'gemini-2.5-flash');

CREATE TABLE IF NOT EXISTS keypress_logs (timestamp TEXT NOT NULL DEFAULT '');

CREATE TABLE IF NOT EXISTS activity_full_text (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dateofentry TEXT NOT NULL DEFAULT '',
    window_title TEXT NOT NULL DEFAULT '',
    window_app_name TEXT NOT NULL DEFAULT '',
    original_full_text TEXT NOT NULL DEFAULT '',
    edited_full_text TEXT NOT NULL DEFAULT '',
    save_count INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS chats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chat_id INTEGER NOT NULL DEFAULT '',
    role TEXT NOT NULL DEFAULT '',
    content TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    FOREIGN KEY (chat_id) REFERENCES chats (id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS permissions (
    app_path TEXT NOT NULL PRIMARY KEY DEFAULT '',
    app_name TEXT NOT NULL DEFAULT '',
    icon_path TEXT NOT NULL DEFAULT '',
    allow BOOLEAN NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS automations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    objective TEXT NOT NULL,
    raw_script TEXT NOT NULL,
    generalized_script TEXT NOT NULL,
    nl_description TEXT,
    additional_instructions TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL DEFAULT '',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS projects_activities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    activity_id INTEGER,
    full_document_text TEXT NOT NULL DEFAULT '',
    document_name TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS task_extracted_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    execution_run_id INTEGER NOT NULL,
    automation_id INTEGER NOT NULL,
    record_name TEXT NOT NULL,
    record_type TEXT NOT NULL,
    source_context TEXT,
    data TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    FOREIGN KEY (execution_run_id) REFERENCES automation_execution_runs(id) ON DELETE CASCADE,
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_task_extracted_data_type ON task_extracted_data(record_type);
CREATE INDEX IF NOT EXISTS idx_task_extracted_data_automation ON task_extracted_data(automation_id);
CREATE INDEX IF NOT EXISTS idx_task_extracted_data_created ON task_extracted_data(created_at DESC);

CREATE TABLE IF NOT EXISTS user_skills (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL,
    skill_type  TEXT    NOT NULL DEFAULT 'site',
    domains     TEXT,
    triggers    TEXT,
    description TEXT    NOT NULL DEFAULT '',
    content     TEXT    NOT NULL,
    is_active   INTEGER NOT NULL DEFAULT 1,
    is_default  INTEGER NOT NULL DEFAULT 0,
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_user_skills_type ON user_skills(skill_type);
CREATE INDEX IF NOT EXISTS idx_user_skills_active ON user_skills(is_active);

CREATE TABLE IF NOT EXISTS automation_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    automation_id INTEGER NOT NULL,
    timestamp TEXT NOT NULL,
    action_type TEXT NOT NULL,
    element_role TEXT,
    element_value TEXT,
    element_description TEXT,
    window_title TEXT,
    window_app_name TEXT,
    importance_score INTEGER NOT NULL DEFAULT 5,
    raw_event_data TEXT,
    sequence_number INTEGER NOT NULL,
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_automation_logs_automation_id ON automation_logs(automation_id);
CREATE INDEX IF NOT EXISTS idx_automation_logs_sequence ON automation_logs(automation_id, sequence_number);

CREATE TABLE IF NOT EXISTS user_auth (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL UNIQUE,
    email TEXT,
    auth_token TEXT NOT NULL,
    token_expires_at TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS usage_sessions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL,
    automation_id INTEGER NOT NULL,
    started_at TEXT NOT NULL,
    ended_at TEXT,
    total_seconds INTEGER DEFAULT 0,
    billed_minutes INTEGER DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'active',
    created_at TEXT NOT NULL,
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_usage_sessions_status ON usage_sessions(status);
CREATE INDEX IF NOT EXISTS idx_usage_sessions_user_id ON usage_sessions(user_id);
CREATE INDEX IF NOT EXISTS idx_usage_sessions_started_at ON usage_sessions(started_at);

CREATE TABLE IF NOT EXISTS automation_execution_runs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    automation_id INTEGER NOT NULL,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    status TEXT NOT NULL DEFAULT 'running',
    additional_instructions TEXT,
    error_message TEXT,
    clipboard TEXT,
    completion_message TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS automation_execution_steps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    execution_run_id INTEGER NOT NULL,
    step_number INTEGER NOT NULL,
    timestamp TEXT NOT NULL,
    explanation TEXT,
    next_step TEXT,
    status TEXT NOT NULL DEFAULT 'executing',
    error_message TEXT,
    step_type TEXT NOT NULL DEFAULT 'action',
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (execution_run_id) REFERENCES automation_execution_runs(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS automation_schedules (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    automation_id INTEGER NOT NULL,
    recurrence_type TEXT NOT NULL DEFAULT 'daily',
    recurrence_days TEXT,
    execution_hour INTEGER NOT NULL DEFAULT 9,
    execution_minute INTEGER NOT NULL DEFAULT 0,
    timezone TEXT NOT NULL DEFAULT 'UTC',
    is_active INTEGER NOT NULL DEFAULT 1,
    next_run_at TEXT,
    last_run_at TEXT,
    persistent_run_id INTEGER DEFAULT NULL,
    continuation_prompt TEXT DEFAULT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (automation_id) REFERENCES automations(id) ON DELETE CASCADE
);
