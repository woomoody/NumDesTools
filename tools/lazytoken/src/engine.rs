//! 平台 JWT token 管理引擎：存储、解码、验证、剪贴板复制。
//! token 列表从外置 JSON 读（不进 git），见 `load_tokens`。

use base64::Engine as _;
use chrono::{DateTime, Local, Utc};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::io;
use std::path::PathBuf;

/// token 定义
#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct TokenDef {
    pub label: String,
    pub token: String,
    #[serde(default)]
    pub notes: String,
}

/// JWT payload 解码后的信息
#[derive(Debug, Clone)]
pub struct JwtInfo {
    pub user_id: String,
    pub exp_str: String,
    pub expired: bool,
    pub expires_in: String,
}

/// 从 base64 URL-safe 解码 JWT payload
fn decode_jwt_payload(token: &str) -> Option<HashMap<String, serde_json::Value>> {
    let parts: Vec<&str> = token.split('.').collect();
    if parts.len() != 3 {
        return None;
    }

    let engine = base64::engine::general_purpose::URL_SAFE_NO_PAD;
    let bytes = engine.decode(parts[1]).ok()?;
    let payload_str = String::from_utf8(bytes).ok()?;
    serde_json::from_str(&payload_str).ok()
}

/// 解析 JWT 返回可读信息
pub fn parse_jwt(token: &str) -> Option<JwtInfo> {
    let payload = decode_jwt_payload(token)?;

    let user_id = payload
        .get("user_id")
        .and_then(|v| v.as_str())
        .unwrap_or("?")
        .to_string();

    let exp = payload.get("exp").and_then(|v| v.as_i64()).unwrap_or(0);

    let now = Utc::now().timestamp();
    let expired = now > exp;

    let dt = DateTime::from_timestamp(exp, 0)
        .map(|d| {
            d.with_timezone(&Local)
                .format("%Y-%m-%d %H:%M:%S")
                .to_string()
        })
        .unwrap_or_else(|| "未知".to_string());

    let diff = exp - now;
    let expires_in = if expired {
        "已过期".to_string()
    } else if diff < 3600 {
        format!("{} 分钟", diff / 60 + 1)
    } else if diff < 86400 {
        format!("{} 小时", diff / 3600 + 1)
    } else {
        format!("{} 天", diff / 86400 + 1)
    };

    Some(JwtInfo {
        user_id,
        exp_str: dt,
        expired,
        expires_in,
    })
}

/// 从外置 JSON 读 token 列表。路径：`%USERPROFILE%\lazytoken.tokens.json`
pub fn load_tokens() -> Vec<TokenDef> {
    let path = tokens_config_path();
    load_tokens_from(&path)
}

pub fn load_tokens_from(path: &std::path::Path) -> Vec<TokenDef> {
    let content = match fs::read_to_string(path) {
        Ok(c) => c,
        Err(_) => return Vec::new(),
    };
    serde_json::from_str(&content).unwrap_or_default()
}

/// 保存 token 列表
pub fn save_tokens(tokens: &[TokenDef]) -> io::Result<()> {
    let path = tokens_config_path();
    let tmp_path = path.with_extension("json.tmp");
    let content = serde_json::to_vec_pretty(tokens).map_err(io::Error::other)?;
    fs::write(&tmp_path, content)?;
    if path.exists() {
        fs::remove_file(&path)?;
    }
    fs::rename(tmp_path, path)
}

fn tokens_config_path() -> PathBuf {
    let home = std::env::var("USERPROFILE").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(home).join("lazytoken.tokens.json")
}

/// 获取 token 的简短显示（遮罩中间部分）
pub fn mask_token(token: &str) -> String {
    if token.len() < 20 {
        return token.to_string();
    }
    let prefix = &token[..12];
    let suffix = &token[token.len() - 6..];
    format!("{}...{}", prefix, suffix)
}

/// 复制到剪贴板（Windows）
pub fn copy_to_clipboard(text: &str) -> io::Result<()> {
    let output = std::process::Command::new("pwsh")
        .args([
            "-NoProfile",
            "-Command",
            &format!("Set-Clipboard -Value '{}'", text.replace('\'', "''")),
        ])
        .output()?;
    if output.status.success() {
        Ok(())
    } else {
        Err(io::Error::other("复制到剪贴板失败"))
    }
}

/// 获取第一个可用 token 的字符串
pub fn get_token<'a>(tokens: &'a [TokenDef], name: Option<&str>) -> Option<&'a str> {
    match name {
        Some(n) => tokens
            .iter()
            .find(|t| t.label == *n)
            .map(|t| t.token.as_str()),
        None => tokens.first().map(|t| t.token.as_str()),
    }
}

/// 代理 API 调用：向 platform 的 /api/<path> 发送请求
pub fn proxy_api(
    token: &str,
    method: &str,
    path: &str,
    body: Option<&str>,
) -> Result<String, String> {
    let agent = ureq::AgentBuilder::new()
        .timeout(std::time::Duration::from_secs(300))
        .build();
    let url = format!("https://ai.solotopiax.com{}", path);
    let req = agent
        .request(method, &url)
        .set("Authorization", &format!("Bearer {}", token))
        .set("Content-Type", "application/json");

    let response = match body {
        Some(b) => req.send_string(b),
        None => req.call(),
    };

    match response {
        Ok(resp) => {
            let status = resp.status();
            let body = resp.into_string().unwrap_or_default();
            if (200..300).contains(&status) {
                Ok(body)
            } else {
                Err(format!("HTTP {}: {}", status, body))
            }
        }
        Err(e) => Err(format!("网络错误: {}", e)),
    }
}

/// 测试 token 是否有效（调用平台 API）
pub fn test_token(token: &str) -> Result<String, String> {
    let agent = ureq::AgentBuilder::new()
        .timeout(std::time::Duration::from_secs(10))
        .build();

    match agent
        .get("https://ai.solotopiax.com/api/menus")
        .set("Authorization", &format!("Bearer {}", token))
        .call()
    {
        Ok(response) => {
            let status = response.status();
            if status == 200 {
                Ok("✅ 有效".to_string())
            } else {
                Err(format!("❌ HTTP {}", status))
            }
        }
        Err(e) => Err(format!("❌ 网络错误: {}", e)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_jwt() {
        // This is a test token from the user
        let token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3ODYyNjc0NDMsInVzZXJfaWQiOiIxNDkifQ.Flczma_CNZ9dzEFm1aDClt2xgVM3qpGguOCLDSwnuSg";
        let info = parse_jwt(token).unwrap();
        assert_eq!(info.user_id, "149");
        assert!(!info.expired);
    }

    #[test]
    fn test_mask_token() {
        let token = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3ODYyNjc0NDMsInVzZXJfaWQiOiIxNDkifQ.Flczma_CNZ9dzEFm1aDClt2xgVM3qpGguOCLDSwnuSg";
        let masked = mask_token(token);
        assert!(masked.contains("..."));
        assert!(masked.starts_with("eyJhbGciOiJ"));
    }

    #[test]
    fn test_load_tokens_from_missing() {
        let dir = std::env::temp_dir();
        let missing = dir.join("notexist.json");
        assert!(load_tokens_from(&missing).is_empty());
    }

    #[test]
    fn test_load_tokens_from_bad_json() {
        let dir = std::env::temp_dir();
        let bad = dir.join("bad.json");
        fs::write(&bad, "not json").unwrap();
        assert!(load_tokens_from(&bad).is_empty());
        fs::remove_file(&bad).ok();
    }
}
