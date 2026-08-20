use log::info;
use rusqlite::{params, Connection, Result};
use rusqlite::OptionalExtension;
use crate::entity::user_auth::{UserAuth, NewUserAuth};

pub fn create_or_update_user_auth(db: &Connection, auth: NewUserAuth) -> Result<UserAuth> {
    let now = chrono::Utc::now().to_rfc3339();
    
    // Try to update existing
    let updated = db.execute(
        "UPDATE user_auth SET auth_token = ?1, token_expires_at = ?2, email = ?3, updated_at = ?4 WHERE user_id = ?5",
        params![auth.auth_token, auth.token_expires_at, auth.email, now, auth.user_id],
    )?;
    
    if updated > 0 {
        // Return the updated auth
        get_user_auth_by_user_id(db, &auth.user_id)?.ok_or_else(|| {
            rusqlite::Error::QueryReturnedNoRows
        })
    } else {
        // Insert new
        db.execute(
            "INSERT INTO user_auth (user_id, email, auth_token, token_expires_at, created_at, updated_at) 
             VALUES (?1, ?2, ?3, ?4, ?5, ?6)",
            params![auth.user_id, auth.email, auth.auth_token, auth.token_expires_at, now, now],
        )?;
        
        let id = db.last_insert_rowid();
        
        Ok(UserAuth {
            id,
            user_id: auth.user_id,
            email: auth.email,
            auth_token: auth.auth_token,
            token_expires_at: auth.token_expires_at,
            created_at: now.clone(),
            updated_at: now,
        })
    }
}

pub fn get_user_auth_by_user_id(db: &Connection, user_id: &str) -> Result<Option<UserAuth>> {
    let mut stmt = db.prepare(
        "SELECT id, user_id, email, auth_token, token_expires_at, created_at, updated_at 
         FROM user_auth WHERE user_id = ?1"
    )?;
    
    let auth = stmt.query_row(params![user_id], |row| {
        Ok(UserAuth {
            id: row.get(0)?,
            user_id: row.get(1)?,
            email: row.get(2)?,
            auth_token: row.get(3)?,
            token_expires_at: row.get(4)?,
            created_at: row.get(5)?,
            updated_at: row.get(6)?,
        })
    }).optional()?;
    
    Ok(auth)
}

#[allow(dead_code)] // Reserved for proxy-JWT refresh flows; current callers always read the full UserAuth row.
pub fn get_valid_auth_token(db: &Connection, user_id: &str) -> Result<Option<String>> {
    let now = chrono::Utc::now().to_rfc3339();
    
    let mut stmt = db.prepare(
        "SELECT auth_token FROM user_auth 
         WHERE user_id = ?1 AND token_expires_at > ?2"
    )?;
    
    let token = stmt.query_row(params![user_id, now], |row| {
        row.get::<_, String>(0)
    }).optional()?;
    
    Ok(token)
}

pub fn clear_user_auth(db: &Connection, user_id: &str) -> Result<()> {
    db.execute("DELETE FROM user_auth WHERE user_id = ?1", params![user_id])?;
    info!("Cleared auth for user: {}", user_id);
    Ok(())
} 