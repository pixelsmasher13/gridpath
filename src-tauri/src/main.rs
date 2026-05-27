// Prevents an extra console window from opening on Windows release builds.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::env;
use std::fs::remove_file;
use std::path::PathBuf;

use lazy_static::lazy_static;
use log::{error, info};
use rusqlite::Connection;
use tauri::menu::{Menu, MenuItem};
use tauri::tray::{TrayIconBuilder, TrayIconEvent};
use tauri::{AppHandle, Emitter, Manager, State, Wry};

use crate::bootstrap::fix_path_env;
use crate::configuration::database;
use crate::configuration::state::{AppState, ServiceAccess};
use crate::entity::setting::Setting;
use crate::repository::settings_repository::{get_setting, insert_or_update_setting};

mod auth;
mod bootstrap;
mod configuration;
mod engine;
mod entity;
mod repository;

lazy_static! {
    /// Path to the single-instance lock file, captured at startup so the
    /// window-destroyed / exit-requested handlers can clean it up even after
    /// the original `main()` scope has unwound into the Tauri runtime.
    static ref LOCK_FILE_PATH: std::sync::Mutex<Option<PathBuf>> = std::sync::Mutex::new(None);
}

// =============================================================================
// Single-instance lock
// =============================================================================

fn check_single_instance() -> Result<PathBuf, String> {
    let lock_dir = dirs::data_local_dir()
        .ok_or_else(|| "Could not determine local data directory".to_string())?
        .join("GridPath");
    std::fs::create_dir_all(&lock_dir).map_err(|e| format!("create lock dir: {}", e))?;
    let lock_path = lock_dir.join("app.lock");

    // Stale-lock detection: if the file exists but the PID inside it is no
    // longer running, the previous instance crashed without cleanup — claim
    // the lock for ourselves rather than refusing to start.
    if lock_path.exists() {
        if let Ok(contents) = std::fs::read_to_string(&lock_path) {
            if let Ok(pid) = contents.trim().parse::<u32>() {
                if !is_pid_running(pid) {
                    let _ = remove_file(&lock_path);
                } else {
                    return Err(format!(
                        "GridPath is already running (PID {}). Close it first.",
                        pid
                    ));
                }
            }
        }
    }

    std::fs::write(&lock_path, std::process::id().to_string())
        .map_err(|e| format!("write lock file: {}", e))?;
    Ok(lock_path)
}

#[cfg(target_os = "windows")]
fn is_pid_running(pid: u32) -> bool {
    use sysinfo::{Pid, ProcessesToUpdate, System};
    let mut sys = System::new();
    sys.refresh_processes(ProcessesToUpdate::All, true);
    sys.process(Pid::from_u32(pid)).is_some()
}

#[cfg(unix)]
fn is_pid_running(pid: u32) -> bool {
    use nix::sys::signal::kill;
    use nix::unistd::Pid;
    kill(Pid::from_raw(pid as i32), None).is_ok()
}

fn cleanup_lock_file() {
    if let Ok(guard) = LOCK_FILE_PATH.lock() {
        if let Some(path) = guard.as_ref() {
            let _ = remove_file(path);
        }
    }
}

// =============================================================================
// Updater commands
// =============================================================================

async fn check_for_update_on_startup(app: AppHandle<Wry>) {
    use tauri_plugin_updater::UpdaterExt;
    info!("Checking for updates...");
    match app.updater() {
        Ok(updater) => match updater.check().await {
            Ok(Some(update)) => {
                info!("Update available: {}", update.version);
                let _ = app.emit(
                    "update-available",
                    serde_json::json!({
                        "version": update.version,
                        "date": update.date.map(|d| d.to_string()),
                        "body": update.body
                    }),
                );
            }
            Ok(None) => info!("App is up to date"),
            Err(e) => error!("Update check failed: {}", e),
        },
        Err(e) => error!("Failed to get updater: {}", e),
    }
}

#[tauri::command]
async fn check_for_updates(app: AppHandle) -> Result<Option<serde_json::Value>, String> {
    use tauri_plugin_updater::UpdaterExt;
    let updater = app.updater().map_err(|e| format!("updater: {}", e))?;
    match updater.check().await {
        Ok(Some(update)) => Ok(Some(serde_json::json!({
            "version": update.version,
            "date": update.date.map(|d| d.to_string()),
            "body": update.body
        }))),
        Ok(None) => Ok(None),
        Err(e) => Err(format!("Update check failed: {}", e)),
    }
}

#[tauri::command]
async fn install_update(app: AppHandle) -> Result<(), String> {
    use tauri_plugin_updater::UpdaterExt;
    let updater = app.updater().map_err(|e| format!("updater: {}", e))?;
    match updater.check().await {
        Ok(Some(update)) => update
            .download_and_install(|_, _| {}, || {})
            .await
            .map_err(|e| format!("install: {}", e)),
        Ok(None) => Err("No update available".to_string()),
        Err(e) => Err(format!("Update check failed: {}", e)),
    }
}

/// Restart the app in-place. Called by the frontend after install_update
/// completes so the new binary takes effect immediately.
#[tauri::command]
fn relaunch_app(app: AppHandle) {
    info!("🔄 Relaunching app to apply update");
    app.restart();
}

// =============================================================================
// Main
// =============================================================================

#[tokio::main]
async fn main() {
    match check_single_instance() {
        Ok(path) => {
            if let Ok(mut guard) = LOCK_FILE_PATH.lock() {
                *guard = Some(path);
            }
        }
        Err(e) => {
            eprintln!("{}", e);
            std::process::exit(1);
        }
    }

    fix_path_env::fix_all_vars().expect("Failed to load env");

    tauri::Builder::default()
        .plugin(tauri_plugin_oauth::init())
        .plugin(
            tauri_plugin_log::Builder::default()
                .level(log::LevelFilter::Info)
                .build(),
        )
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .invoke_handler(tauri::generate_handler![
            // Updater
            check_for_updates,
            install_update,
            relaunch_app,
            // ChatGPT Plus/Pro (OpenAI Codex) OAuth
            crate::auth::openai_codex_oauth::openai_codex_login,
            crate::auth::openai_codex_oauth::openai_codex_logout,
            crate::auth::openai_codex_oauth::openai_codex_status,
            // Spreadsheet workspace file I/O
            crate::engine::workbook::commands::read_workbook_file,
            crate::engine::workbook::commands::write_workbook_file,
            crate::engine::workbook::commands::append_change_batch,
            crate::engine::workbook::commands::read_change_log,
            // Spreadsheet agent
            crate::engine::spreadsheet_agent::commands::spreadsheet_agent_turn,
            crate::engine::spreadsheet_agent::commands::spreadsheet_agent_stop,
            crate::engine::spreadsheet_agent::commands::spreadsheet_tool_result,
            crate::engine::spreadsheet_agent::commands::generate_session_title,
            // Spreadsheet session persistence
            crate::engine::workbook::session_commands::spreadsheet_session_upsert,
            crate::engine::workbook::session_commands::spreadsheet_session_rename,
            crate::engine::workbook::session_commands::spreadsheet_session_archive,
            crate::engine::workbook::session_commands::spreadsheet_session_delete,
            crate::engine::workbook::session_commands::spreadsheet_session_list,
            crate::engine::workbook::session_commands::spreadsheet_session_get_messages,
            crate::engine::workbook::session_commands::spreadsheet_session_append_message,
            crate::engine::workbook::session_commands::spreadsheet_session_add_tokens,
            // Spreadsheet workspace settings
            crate::engine::workbook::settings_commands::ssws_get_setting,
            crate::engine::workbook::settings_commands::ssws_set_setting,
            crate::engine::workbook::settings_commands::ssws_get_model,
            crate::engine::workbook::settings_commands::ssws_set_model,
        ])
        .manage(AppState { db: Default::default() })
        .on_window_event(|window, event| match event {
            tauri::WindowEvent::CloseRequested { api, .. } => {
                // Closing the main window hides it to tray instead of quitting,
                // so an in-flight agent turn isn't killed when the user clicks X.
                api.prevent_close();
                let _ = window.hide();
            }
            tauri::WindowEvent::Destroyed => {
                cleanup_lock_file();
            }
            _ => {}
        })
        .setup(move |app| {
            // Tray: single quit menu item; left-click toggles window visibility.
            let quit = MenuItem::with_id(app, "quit", "Quit", true, None::<&str>)?;
            let menu = Menu::with_items(app, &[&quit])?;
            let tray_icon = app.default_window_icon().unwrap().clone();
            let tray = TrayIconBuilder::with_id("gridpath-main-tray")
                .menu(&menu)
                .icon(tray_icon)
                .show_menu_on_left_click(false)
                .on_tray_icon_event(|tray, event| {
                    if let TrayIconEvent::Click { button, button_state, .. } = event {
                        use tauri::tray::{MouseButton, MouseButtonState};
                        if button == MouseButton::Left && button_state == MouseButtonState::Up {
                            if let Some(window) = tray.app_handle().get_webview_window("main") {
                                let visible = window.is_visible().unwrap_or(false);
                                let minimized = window.is_minimized().unwrap_or(false);
                                let focused = window.is_focused().unwrap_or(false);
                                if !visible || minimized || !focused {
                                    let _ = window.show();
                                    if minimized {
                                        let _ = window.unminimize();
                                    }
                                    let _ = window.set_focus();
                                } else {
                                    let _ = window.hide();
                                }
                            }
                        }
                    }
                })
                .on_menu_event(|app, event| {
                    if event.id().as_ref() == "quit" {
                        cleanup_lock_file();
                        app.exit(0);
                    }
                })
                .build(app)?;
            app.manage(tray);

            // `--minimized` lets autostarted launches come up to tray only.
            let args: Vec<String> = env::args().collect();
            let start_minimized = args.contains(&"--minimized".to_string());
            if let Some(window) = app.get_webview_window("main") {
                if start_minimized {
                    let _ = window.hide();
                } else {
                    let _ = window.show();
                }
            }

            let app_handle = app.handle();

            // DB connect + diesel migrations.
            let db: Connection = database::initialize_database(app_handle)
                .expect("Database initialization failed");
            {
                let state: State<AppState> = app_handle.state();
                *state.db.lock().unwrap() = Some(db);
            }

            // Hydrate the in-memory model overrides from settings so the
            // agent picks the user's chosen model on the very first turn
            // rather than the built-in default.
            let m_claude = app_handle
                .db(|db| get_setting(db, "model_claude").map(|s| s.setting_value).unwrap_or_default());
            let mut m_openai_codex = app_handle.db(|db| {
                get_setting(db, "model_openai_codex").map(|s| s.setting_value).unwrap_or_default()
            });

            // One-shot migration: ChatGPT subscriptions cannot use `gpt-5-codex` or
            // `gpt-5` (the Codex /responses endpoint rejects them with HTTP 400).
            // Older builds shipped these as defaults and persisted them — clear any
            // known-bad stored value so the new default kicks in.
            if matches!(m_openai_codex.as_str(), "gpt-5-codex" | "gpt-5") {
                info!(
                    "Migrating invalid ChatGPT subscription model '{}' -> default '{}'",
                    m_openai_codex,
                    crate::engine::provider_config::DEFAULT_OPENAI_CODEX_MODEL
                );
                let _ = app_handle.db(|db| {
                    insert_or_update_setting(
                        db,
                        Setting {
                            setting_key: "model_openai_codex".to_string(),
                            setting_value: String::new(),
                        },
                    )
                });
                m_openai_codex.clear();
            }

            crate::engine::provider_config::set_model("claude", &m_claude);
            crate::engine::provider_config::set_model("openai-codex", &m_openai_codex);

            let codex_effort = app_handle.db(|db| {
                get_setting(db, "openai_codex_reasoning_effort")
                    .map(|s| s.setting_value)
                    .unwrap_or_default()
            });
            crate::engine::provider_config::set_openai_codex_reasoning_effort(&codex_effort);

            // Updater check, non-blocking.
            let app_for_updater = app.handle().clone();
            tauri::async_runtime::spawn(async move {
                check_for_update_on_startup(app_for_updater).await;
            });

            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("error building tauri application")
        .run(|_app_handle, event| {
            if let tauri::RunEvent::ExitRequested { api, .. } = event {
                cleanup_lock_file();
                // Stay alive in the tray on macOS/Linux; only Windows exits
                // when the window closes (tray icon stays visible regardless).
                #[cfg(not(target_os = "windows"))]
                api.prevent_exit();
                #[cfg(target_os = "windows")]
                let _ = api;
            }
        });
}
