// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
use tauri::Manager;
#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[cfg(not(any(target_os = "android", target_os = "ios")))]
#[tauri::command]
async fn zoho_exchange_token(
    client_id: String,
    client_secret: String,
    code: String,
) -> Result<String, String> {
    let client = reqwest::Client::new();
    let params = [
        ("grant_type", "authorization_code"),
        ("client_id", &client_id),
        ("client_secret", &client_secret),
        ("code", &code),
    ];
    let resp = client
        .post("https://accounts.zoho.eu/oauth/v2/token")
        .form(&params)
        .send()
        .await
        .map_err(|e| e.to_string())?;
    let body = resp.text().await.map_err(|e| e.to_string())?;
    Ok(body)
}

/// Hilfsfunktion: Access Token via Refresh Token holen.
#[cfg(not(any(target_os = "android", target_os = "ios")))]
async fn get_zoho_token(
    client: &reqwest::Client,
    client_id: &str,
    client_secret: &str,
    refresh_token: &str,
) -> Result<String, String> {
    let params = [
        ("grant_type", "refresh_token"),
        ("client_id", client_id),
        ("client_secret", client_secret),
        ("refresh_token", refresh_token),
    ];
    let resp = client
        .post("https://accounts.zoho.eu/oauth/v2/token")
        .form(&params)
        .send()
        .await
        .map_err(|e| e.to_string())?;
    let body: serde_json::Value = resp.json().await.map_err(|e| e.to_string())?;
    body["access_token"]
        .as_str()
        .ok_or_else(|| format!("Token-Refresh fehlgeschlagen: {}", body))
        .map(|s| s.to_string())
}

/// Einzelner GET-Request an Zoho CRM.
#[cfg(not(any(target_os = "android", target_os = "ios")))]
#[tauri::command]
async fn zoho_api_get(
    client_id: String,
    client_secret: String,
    refresh_token: String,
    path: String,
) -> Result<String, String> {
    let client = reqwest::Client::new();
    let token = get_zoho_token(&client, &client_id, &client_secret, &refresh_token).await?;
    let url = format!("https://www.zohoapis.eu/crm/v8/{}", path);
    let resp = client
        .get(&url)
        .header("Authorization", format!("Zoho-oauthtoken {}", token))
        .send()
        .await
        .map_err(|e| e.to_string())?;
    resp.text().await.map_err(|e| e.to_string())
}

/// Lädt ALLE Accounts paginiert (200 pro Seite) und gibt ein JSON-Array zurück.
#[cfg(not(any(target_os = "android", target_os = "ios")))]
#[tauri::command]
async fn zoho_get_all_accounts(
    client_id: String,
    client_secret: String,
    refresh_token: String,
) -> Result<String, String> {
    let client = reqwest::Client::new();
    let token = get_zoho_token(&client, &client_id, &client_secret, &refresh_token).await?;

    let mut all_accounts: Vec<serde_json::Value> = Vec::new();
    let mut page = 1u32;

    loop {
        let url = format!(
            "https://www.zohoapis.eu/crm/v8/Accounts?fields=id,Account_Name&per_page=200&page={}",
            page
        );
        let resp = client
            .get(&url)
            .header("Authorization", format!("Zoho-oauthtoken {}", token))
            .send()
            .await
            .map_err(|e| e.to_string())?;

        if resp.status() == 204 {
            break; // keine weiteren Daten
        }

        let body: serde_json::Value = resp.json().await.map_err(|e| e.to_string())?;

        if let Some(records) = body["data"].as_array() {
            let count = records.len();
            all_accounts.extend(records.iter().cloned());
            // Zoho gibt mehr_records: true/false im info-Objekt
            let more = body["info"]["more_records"].as_bool().unwrap_or(false);
            if !more || count < 200 {
                break;
            }
            page += 1;
        } else {
            // Fehler oder leere Antwort
            break;
        }
    }

    serde_json::to_string(&all_accounts).map_err(|e| e.to_string())
}

/// Sendet einen POST/PUT zu einem Zoho CRM Endpunkt.
#[cfg(not(any(target_os = "android", target_os = "ios")))]
#[tauri::command]
async fn zoho_api_post(
    client_id: String,
    client_secret: String,
    refresh_token: String,
    path: String,
    body: String,
) -> Result<String, String> {
    let client = reqwest::Client::new();

    // Access Token holen
    let token_params = [
        ("grant_type", "refresh_token"),
        ("client_id", &client_id),
        ("client_secret", &client_secret),
        ("refresh_token", &refresh_token),
    ];
    let token_resp = client
        .post("https://accounts.zoho.eu/oauth/v2/token")
        .form(&token_params)
        .send()
        .await
        .map_err(|e| e.to_string())?;
    let token_body: serde_json::Value = token_resp.json().await.map_err(|e| e.to_string())?;
    let access_token = token_body["access_token"]
        .as_str()
        .ok_or_else(|| format!("Token-Refresh fehlgeschlagen: {}", token_body))?
        .to_string();

    // API-Request
    let url = format!("https://www.zohoapis.eu/crm/v8/{}", path);
    let resp = client
        .post(&url)
        .header("Authorization", format!("Zoho-oauthtoken {}", access_token))
        .header("Content-Type", "application/json")
        .body(body)
        .send()
        .await
        .map_err(|e| e.to_string())?;
    let resp_body = resp.text().await.map_err(|e| e.to_string())?;
    Ok(resp_body)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .setup(|app| {
            #[cfg(debug_assertions)]
            app.get_webview_window("main").unwrap().open_devtools();
            Ok(())
        })
        .plugin(tauri_plugin_window_state::Builder::new().build())
        .plugin(tauri_plugin_store::Builder::new().build())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_http::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![greet, zoho_exchange_token, zoho_api_get, zoho_api_post, zoho_get_all_accounts])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
