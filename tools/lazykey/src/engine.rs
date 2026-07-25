//! LiteLLM key 切换引擎：配置、spend 查询和文件替换，不碰终端。
//! 原理：key 字符串全局唯一，字面 replace 即可，不用 per-file regex。
//! key 列表从外置 JSON 读（不进 git），见 `load_keys`。

use chrono::Local;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::time::Duration;

const KEY_INFO_URL: &str = "https://litellm.solotopia.net/key/info";

/// key 定义：label（菜单显示）、alias（直切短名）、key 字符串
#[derive(Debug, Clone, Deserialize)]
pub struct KeyDef {
    pub label: String,
    pub alias: String,
    pub key: String,
    #[serde(default)]
    pub admin_key: bool,
    /// 外部 key（如 DeepSeek 官方 API key，非 LiteLLM 管理），跳过 spend 查询
    #[serde(default)]
    pub external: bool,
}

#[derive(Debug, Clone, Deserialize)]
pub struct KeySpend {
    pub key_alias: String,
    pub spend: f64,
}

#[derive(Deserialize)]
struct KeyInfoResponse {
    info: KeySpend,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SpendSnapshot {
    #[serde(default)]
    pub snapshot_month: String,
    #[serde(default)]
    pub snapshots: HashMap<String, f64>,
}

/// Current local calendar month in `YYYY-MM` format.
pub fn current_month() -> String {
    Local::now().format("%Y-%m").to_string()
}

/// Loads the local spend baseline. Missing or invalid files produce an empty snapshot.
pub fn load_snapshot() -> SpendSnapshot {
    load_snapshot_from(&snapshot_path())
}

fn load_snapshot_from(path: &Path) -> SpendSnapshot {
    fs::read_to_string(path)
        .ok()
        .and_then(|content| serde_json::from_str(&content).ok())
        .unwrap_or_default()
}

/// Saves the spend baseline through a temporary file before replacing the destination.
pub fn save_snapshot(snapshot: &SpendSnapshot) -> io::Result<()> {
    save_snapshot_to(&snapshot_path(), snapshot)
}

fn save_snapshot_to(path: &Path, snapshot: &SpendSnapshot) -> io::Result<()> {
    let temporary_path = path.with_extension("json.tmp");
    let content = serde_json::to_vec_pretty(snapshot).map_err(io::Error::other)?;
    fs::write(&temporary_path, content)?;
    if path.exists() {
        fs::remove_file(path)?;
    }
    fs::rename(temporary_path, path)
}

/// Converts lifetime accumulated spend into spend since this month's baseline.
pub fn compute_period_spend(current_spend: &HashMap<String, f64>) -> HashMap<String, f64> {
    if current_spend.is_empty() {
        return HashMap::new();
    }

    compute_period_spend_from(
        current_spend,
        load_snapshot(),
        &current_month(),
        save_snapshot,
    )
}

#[cfg(test)]
fn compute_period_spend_at(
    current_spend: &HashMap<String, f64>,
    path: &Path,
    month: &str,
) -> HashMap<String, f64> {
    if current_spend.is_empty() {
        return HashMap::new();
    }

    compute_period_spend_from(current_spend, load_snapshot_from(path), month, |snapshot| {
        save_snapshot_to(path, snapshot)
    })
}

fn compute_period_spend_from(
    current_spend: &HashMap<String, f64>,
    snapshot: SpendSnapshot,
    month: &str,
    save: impl FnOnce(&SpendSnapshot) -> io::Result<()>,
) -> HashMap<String, f64> {
    match snapshot.snapshot_month.as_str().cmp(month) {
        std::cmp::Ordering::Greater => HashMap::new(),
        std::cmp::Ordering::Equal => current_spend
            .iter()
            .filter_map(|(key, accumulated)| {
                snapshot
                    .snapshots
                    .get(key)
                    .map(|baseline| (key.clone(), (accumulated - baseline).max(0.0)))
            })
            .collect(),
        std::cmp::Ordering::Less => {
            let refreshed = SpendSnapshot {
                snapshot_month: month.to_string(),
                snapshots: current_spend.clone(),
            };
            if save(&refreshed).is_err() {
                return HashMap::new();
            }
            current_spend.keys().map(|key| (key.clone(), 0.0)).collect()
        }
    }
}

fn snapshot_path() -> PathBuf {
    let home = std::env::var("USERPROFILE").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(home).join("lazykey.spend-snapshot.json")
}

/// Fetches each key's accumulated LiteLLM spend in USD.
/// Failed requests are represented as `0.0` so network issues never stop the TUI.
pub fn fetch_spend(admin_key: &str, keys: &[KeyDef]) -> HashMap<String, f64> {
    let agent = ureq::AgentBuilder::new()
        .timeout(Duration::from_secs(5))
        .build();
    let authorization = format!("Bearer {admin_key}");

    let mut any_success = false;
    let spend_map = keys
        .iter()
        .filter(|k| !k.external) // 跳过外部 key（非 LiteLLM 管理，无 spend 数据）
        .map(|key| {
            let spend = agent
                .get(KEY_INFO_URL)
                .query("key", &key.key)
                .set("Authorization", &authorization)
                .call()
                .ok()
                .and_then(|response| response.into_string().ok())
                .and_then(|body| parse_key_spend(&body));
            any_success |= spend.is_some();
            (key.key.clone(), spend)
        })
        .map(|(key, spend)| (key, spend.unwrap_or(0.0)))
        .collect();

    if any_success {
        spend_map
    } else {
        HashMap::new()
    }
}

fn parse_key_spend(body: &str) -> Option<f64> {
    let response: KeyInfoResponse = serde_json::from_str(body).ok()?;
    let _key_alias = response.info.key_alias;
    Some(response.info.spend)
}

/// 从外置 JSON 读 key 列表。路径：`%USERPROFILE%\lazykey.keys.json`（不进 git，纯粹本地配置）。
/// 文件不存在/解析失败返回空列表（调用方要处理"没有 key 可用"的情况）。
pub fn load_keys() -> Vec<KeyDef> {
    let path = keys_config_path();
    load_keys_from(&path)
}

/// 测试用：从指定路径读
pub fn load_keys_from(path: &Path) -> Vec<KeyDef> {
    let content = match fs::read_to_string(path) {
        Ok(c) => c,
        Err(_) => return Vec::new(),
    };
    serde_json::from_str(&content).unwrap_or_default()
}

fn keys_config_path() -> PathBuf {
    let home = std::env::var("USERPROFILE").unwrap_or_else(|_| ".".to_string());
    PathBuf::from(home).join("lazykey.keys.json")
}

/// 全局承载 key 的 4 个固定文件 + 自动扫描 home_root 下 CC* 项目的 .claude\settings.json
pub fn get_key_target_files(home_root: &Path) -> Vec<PathBuf> {
    let mut files = vec![
        home_root.join(".claude").join("settings.json"),
        home_root
            .join("AppData")
            .join("Roaming")
            .join("Code")
            .join("User")
            .join("settings.json"),
        home_root.join("Documents").join("NumDesGlobalKey.json"),
        home_root
            .join("Documents")
            .join("LazyGit")
            .join("ai_commit.ps1"),
        // opencode 配置（.config/opencode/opencode.jsonc 的 apiKey）
        home_root
            .join(".config")
            .join("opencode")
            .join("opencode.jsonc"),
        // hermes 配置（AppData/Local/hermes/config.yaml 的 api_key）
        home_root
            .join("AppData")
            .join("Local")
            .join("hermes")
            .join("config.yaml"),
    ];

    if let Ok(entries) = fs::read_dir(home_root) {
        for entry in entries.flatten() {
            let path = entry.path();
            if !path.is_dir() {
                continue;
            }
            if let Some(name) = path.file_name().and_then(|n| n.to_str()) {
                if name.starts_with("CC") {
                    let settings = path.join(".claude").join("settings.json");
                    if settings.exists() {
                        files.push(settings);
                    }
                }
            }
        }
    }

    files.retain(|f| f.exists());
    files
}

/// 探测文件当前含哪把已知 key；都没命中或文件不存在返回 None
pub fn find_file_key(file: &Path, key_values: &[&str]) -> Option<String> {
    let content = fs::read_to_string(file).ok()?;
    for k in key_values {
        if content.contains(k) {
            return Some(k.to_string());
        }
    }
    None
}

/// 一键直切：对每个文件自动探旧 key → 换成 new_key；已是目标/无已知 key/不存在 跳过
/// 返回 (changed 已切文件, skipped 跳过原因列表)
pub fn switch_files_to_key(
    new_key: &str,
    files: &[PathBuf],
    key_values: &[&str],
) -> (Vec<PathBuf>, Vec<String>) {
    let mut changed = Vec::new();
    let mut skipped = Vec::new();
    for f in files {
        let content = match fs::read_to_string(f) {
            Ok(c) => c,
            Err(_) => {
                skipped.push(format!("{} (不存在)  {}", label_for_path(f), f.display()));
                continue;
            }
        };
        if content.contains(new_key) {
            skipped.push(format!("{} (已是目标)  {}", label_for_path(f), f.display()));
            continue;
        }
        let old_key = match find_file_key(f, key_values) {
            Some(k) => k,
            None => {
                skipped.push(format!(
                    "{} (无已知 key)  {}",
                    label_for_path(f),
                    f.display()
                ));
                continue;
            }
        };
        let next = content.replace(&old_key, new_key);
        if next != content && fs::write(f, next).is_ok() {
            changed.push(f.clone());
        }
    }
    (changed, skipped)
}

/// 直切命令行别名 → key；未知名返回 None，完整 sk- 透传
pub fn resolve_key_alias(name: &str, keys: &[KeyDef]) -> Option<String> {
    for def in keys {
        if def.alias.eq_ignore_ascii_case(name) {
            return Some(def.key.clone());
        }
    }
    if name.starts_with("sk-") {
        Some(name.to_string())
    } else {
        None
    }
}

/// label_of：由 key 找显示名
pub fn label_of<'a>(key: &str, keys: &'a [KeyDef]) -> Option<&'a str> {
    keys.iter().find(|d| d.key == key).map(|d| d.label.as_str())
}

/// 所有 key 的字符串列表（探测用）
pub fn all_key_values(keys: &[KeyDef]) -> Vec<&str> {
    keys.iter().map(|d| d.key.as_str()).collect()
}

/// 路径 → 形象短名（TUI/输出显示用，替代长路径）
pub fn label_for_path(path: &Path) -> String {
    let s = path.to_string_lossy().replace('/', "\\");
    if s.ends_with("\\.claude\\settings.json") {
        // .claude 的 parent 是 home → 全局，否则 CC-<项目名>
        let claude_parent = path.parent().and_then(|p| p.parent());
        let home = std::env::var("USERPROFILE").unwrap_or_default();
        let parent_str = claude_parent.map(|p| p.to_string_lossy().replace('/', "\\"));
        if parent_str.as_deref() == Some(&home) {
            "CC全局".to_string()
        } else {
            let name = claude_parent
                .and_then(|p| p.file_name())
                .map(|n| n.to_string_lossy().to_string())
                .unwrap_or_default();
            format!("CC-{}", name)
        }
    } else if s.contains("\\Code\\User\\settings.json") {
        "VSCode".to_string()
    } else if s.ends_with("\\NumDesGlobalKey.json") {
        "全局key库".to_string()
    } else if s.ends_with("\\ai_commit.ps1") {
        "LazyGit提交".to_string()
    } else if s.ends_with("\\opencode.jsonc") {
        "opencode".to_string()
    } else if s.ends_with("\\hermes\\config.yaml") {
        "hermes".to_string()
    } else {
        path.file_name()
            .map(|n| n.to_string_lossy().to_string())
            .unwrap_or_else(|| s)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::env;

    fn fixture_dir() -> PathBuf {
        let dir = env::temp_dir().join(format!(
            "lazykey-test-{}-{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        fs::create_dir_all(&dir).unwrap();
        dir
    }

    fn write_file(dir: &Path, name: &str, content: &str) -> PathBuf {
        let p = dir.join(name);
        if let Some(parent) = p.parent() {
            fs::create_dir_all(parent).unwrap();
        }
        fs::write(&p, content).unwrap();
        p
    }

    #[test]
    fn test_snapshot_stale_month_refreshes() {
        let dir = fixture_dir();
        let path = dir.join("snapshot.json");
        let old = SpendSnapshot {
            snapshot_month: "2026-06".to_string(),
            snapshots: HashMap::from([("sk-test".to_string(), 100.0)]),
        };
        save_snapshot_to(&path, &old).unwrap();
        let current = HashMap::from([("sk-test".to_string(), 150.0)]);

        let period = compute_period_spend_at(&current, &path, "2026-07");

        assert_eq!(period.get("sk-test"), Some(&0.0));
        let refreshed = load_snapshot_from(&path);
        assert_eq!(refreshed.snapshot_month, "2026-07");
        assert_eq!(refreshed.snapshots.get("sk-test"), Some(&150.0));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_snapshot_current_month_subtracts() {
        let dir = fixture_dir();
        let path = dir.join("snapshot.json");
        let baseline = SpendSnapshot {
            snapshot_month: "2026-07".to_string(),
            snapshots: HashMap::from([("sk-test".to_string(), 100.0)]),
        };
        save_snapshot_to(&path, &baseline).unwrap();
        let current = HashMap::from([("sk-test".to_string(), 150.0)]);

        let period = compute_period_spend_at(&current, &path, "2026-07");

        assert_eq!(period.get("sk-test"), Some(&50.0));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_missing_snapshot_creates_baseline() {
        let dir = fixture_dir();
        let path = dir.join("snapshot.json");
        let current = HashMap::from([("sk-test".to_string(), 150.0)]);

        let period = compute_period_spend_at(&current, &path, "2026-07");

        assert_eq!(period.get("sk-test"), Some(&0.0));
        let created = load_snapshot_from(&path);
        assert_eq!(created.snapshot_month, "2026-07");
        assert_eq!(created.snapshots, current);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_corrupt_snapshot_graceful() {
        let dir = fixture_dir();
        let path = write_file(&dir, "snapshot.json", "not json");

        let snapshot = load_snapshot_from(&path);

        assert!(snapshot.snapshot_month.is_empty());
        assert!(snapshot.snapshots.is_empty());
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn test_empty_current_spend_returns_empty() {
        let dir = fixture_dir();
        let path = dir.join("snapshot.json");

        let period = compute_period_spend_at(&HashMap::new(), &path, "2026-07");

        assert!(period.is_empty());
        assert!(!path.exists());
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_keys_from_reads_json() {
        let dir = fixture_dir();
        let json = r#"[
            { "label": "cent(自己)", "alias": "cent", "key": "sk-testCent111" },
            { "label": "休眠(借用)", "alias": "sleep", "key": "sk-testSleep222" }
        ]"#;
        let p = write_file(&dir, "keys.json", json);
        let keys = load_keys_from(&p);
        assert_eq!(keys.len(), 2);
        assert_eq!(keys[0].alias, "cent");
        assert!(!keys[0].admin_key);
        assert_eq!(keys[1].key, "sk-testSleep222");
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_keys_from_reads_admin_key_marker() {
        let dir = fixture_dir();
        let json = r#"[
            { "label": "cent", "alias": "cent", "key": "sk-admin", "admin_key": true }
        ]"#;
        let p = write_file(&dir, "keys.json", json);

        let keys = load_keys_from(&p);

        assert!(keys[0].admin_key);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_keys_from_reads_external_key() {
        let dir = fixture_dir();
        let json = r#"[
            { "label": "外部key", "alias": "ext", "key": "sk-ext", "external": true }
        ]"#;
        let p = write_file(&dir, "keys.json", json);

        let keys = load_keys_from(&p);

        assert!(keys[0].external);
        assert!(!keys[0].admin_key);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_keys_from_defaults_external_false() {
        let dir = fixture_dir();
        let json = r#"[
            { "label": "cent", "alias": "cent", "key": "sk-cent" }
        ]"#;
        let p = write_file(&dir, "keys.json", json);

        let keys = load_keys_from(&p);

        assert!(!keys[0].external);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn parse_key_spend_reads_nested_info() {
        let json = r#"{"key":"sk-user","info":{"key_alias":"cent","spend":123.45}}"#;

        let spend = parse_key_spend(json);

        assert_eq!(spend, Some(123.45));
    }

    #[test]
    fn parse_key_spend_rejects_invalid_response() {
        assert_eq!(parse_key_spend("not json"), None);
        assert_eq!(parse_key_spend(r#"{"info":{}}"#), None);
    }

    #[test]
    fn load_keys_from_returns_empty_for_missing_or_bad() {
        let dir = fixture_dir();
        let missing = dir.join("notexist.json");
        assert!(load_keys_from(&missing).is_empty());
        let bad = write_file(&dir, "bad.json", "not json");
        assert!(load_keys_from(&bad).is_empty());
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn find_file_key_detects_known() {
        let dir = fixture_dir();
        let keys = vec!["sk-testCent111", "sk-testSleep222"];
        let f = write_file(&dir, "a.json", "\"sk-testCent111\"");
        assert_eq!(find_file_key(&f, &keys), Some("sk-testCent111".to_string()));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn find_file_key_returns_none_for_unknown() {
        let dir = fixture_dir();
        let keys = vec!["sk-testCent111"];
        let f = write_file(&dir, "b.json", "\"sk-unknown999\"");
        assert_eq!(find_file_key(&f, &keys), None);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn switch_files_to_key_switches_only_known() {
        let dir = fixture_dir();
        let keys = vec!["sk-testCent111", "sk-testSleep222"];
        let f1 = write_file(&dir, "a.json", "\"sk-testCent111\"");
        let f2 = write_file(&dir, "b.json", "\"sk-unknown999\"");
        let (changed, skipped) =
            switch_files_to_key("sk-testSleep222", &[f1.clone(), f2.clone()], &keys);
        assert_eq!(changed, vec![f1.clone()]);
        assert_eq!(skipped.len(), 1);
        assert!(fs::read_to_string(&f1).unwrap().contains("sk-testSleep222"));
        assert!(fs::read_to_string(&f2).unwrap().contains("sk-unknown999"));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn switch_files_to_key_skips_already_target() {
        let dir = fixture_dir();
        let keys = vec!["sk-testCent111"];
        let f = write_file(&dir, "a.json", "\"sk-testCent111\"");
        let (changed, skipped) = switch_files_to_key("sk-testCent111", &[f], &keys);
        assert!(changed.is_empty());
        assert_eq!(skipped.len(), 1);
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn resolve_key_alias_maps_and_passthrough() {
        let keys = vec![
            KeyDef {
                label: "cent".into(),
                alias: "cent".into(),
                key: "sk-c1".into(),
                admin_key: false,
                external: false,
            },
            KeyDef {
                label: "sleep".into(),
                alias: "sleep".into(),
                key: "sk-s2".into(),
                admin_key: false,
                external: false,
            },
        ];
        assert_eq!(resolve_key_alias("cent", &keys), Some("sk-c1".to_string()));
        assert_eq!(
            resolve_key_alias("sk-custom", &keys),
            Some("sk-custom".to_string())
        );
        assert_eq!(resolve_key_alias("notexist", &keys), None);
    }

    #[test]
    fn get_key_target_files_scans_cc_projects() {
        let dir = fixture_dir();
        let home = dir.join("home");
        fs::create_dir_all(home.join(".claude")).unwrap();
        fs::create_dir_all(
            home.join("AppData")
                .join("Roaming")
                .join("Code")
                .join("User"),
        )
        .unwrap();
        fs::create_dir_all(home.join("Documents").join("LazyGit")).unwrap();
        fs::write(home.join(".claude").join("settings.json"), "{}").unwrap();
        fs::write(
            home.join("AppData")
                .join("Roaming")
                .join("Code")
                .join("User")
                .join("settings.json"),
            "{}",
        )
        .unwrap();
        fs::write(home.join("Documents").join("NumDesGlobalKey.json"), "{}").unwrap();
        fs::write(
            home.join("Documents").join("LazyGit").join("ai_commit.ps1"),
            "",
        )
        .unwrap();
        fs::create_dir_all(home.join("CCglm").join(".claude")).unwrap();
        fs::write(
            home.join("CCglm").join(".claude").join("settings.json"),
            "{}",
        )
        .unwrap();
        fs::create_dir_all(home.join("CCKimi").join(".claude")).unwrap();
        fs::write(
            home.join("CCKimi").join(".claude").join("settings.json"),
            "{}",
        )
        .unwrap();
        fs::create_dir_all(home.join("CCNoSettings")).unwrap();

        let targets = get_key_target_files(&home);
        assert_eq!(targets.len(), 6);
        assert!(targets
            .iter()
            .any(|t| t.to_string_lossy().contains("CCglm")));
        assert!(targets
            .iter()
            .any(|t| t.to_string_lossy().contains("CCKimi")));
        assert!(!targets
            .iter()
            .any(|t| t.to_string_lossy().contains("CCNoSettings")));
        fs::remove_dir_all(&dir).ok();
    }
}
