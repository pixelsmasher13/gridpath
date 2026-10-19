use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct UserAuth {
    pub id: i64,
    pub user_id: String,
    pub email: Option<String>,
    pub auth_token: String,
    pub token_expires_at: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NewUserAuth {
    pub user_id: String,
    pub email: Option<String>,
    pub auth_token: String,
    pub token_expires_at: String,
} 