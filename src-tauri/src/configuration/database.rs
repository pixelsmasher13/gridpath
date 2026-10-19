use std::fs;

use diesel::sqlite::SqliteConnection;
use diesel::Connection;
use diesel_migrations::{embed_migrations, EmbeddedMigrations, MigrationHarness};
use log::info;
use tauri::{AppHandle, Manager};


const MIGRATIONS: EmbeddedMigrations = embed_migrations!();

pub fn initialize_database(
    app_handle: &AppHandle,
) -> Result<rusqlite::Connection, Box<dyn std::error::Error>> {
    let app_dir = app_handle
        .path()
        .app_data_dir()
        .expect("The app data directory should exist.");
    fs::create_dir_all(&app_dir).expect("The app data directory should be created.");
    let sqlite_path = app_dir.join("gridpath.sqlite");
    info!("SQLITE_PATH: {}", sqlite_path.display());
    let db = rusqlite::Connection::open(sqlite_path.clone())?;
    let user_pragma = db.prepare("PRAGMA user_version")?;
    drop(user_pragma);
    let mut connection_diesel =
        SqliteConnection::establish(sqlite_path.display().to_string().as_str())
            .unwrap_or_else(|_| panic!("Error connecting to {}", "database"));
    connection_diesel
        .run_pending_migrations(MIGRATIONS)
        .unwrap();
    
    Ok(db)
}




