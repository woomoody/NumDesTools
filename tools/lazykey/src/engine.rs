//! LiteLLM key 切换引擎：纯函数，不碰终端，可单测。
//! 原理：key 字符串全局唯一，字面 replace 即可，不用 per-file regex。
//! key 列表从外置 JSON 读（不进 git），见 `load_keys`。

use serde::Deserialize;
use std::fs;
use std::path::{Path, PathBuf};

/// key 定义：label（菜单显示）、alias（直切短名）、key 字符串
#[derive(Debug, Clone, Deserialize)]
pub struct KeyDef {
    pub label: String,
    pub alias: String,
    pub key: String,
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
                skipped.push(format!("{} (不存在)", f.display()));
                continue;
            }
        };
        if content.contains(new_key) {
            skipped.push(format!("{} (已是目标)", f.display()));
            continue;
        }
        let old_key = match find_file_key(f, key_values) {
            Some(k) => k,
            None => {
                skipped.push(format!("{} (无已知 key)", f.display()));
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
        assert_eq!(keys[1].key, "sk-testSleep222");
        fs::remove_dir_all(&dir).ok();
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
            KeyDef { label: "cent".into(), alias: "cent".into(), key: "sk-c1".into() },
            KeyDef { label: "sleep".into(), alias: "sleep".into(), key: "sk-s2".into() },
        ];
        assert_eq!(resolve_key_alias("cent", &keys), Some("sk-c1".to_string()));
        assert_eq!(resolve_key_alias("sk-custom", &keys), Some("sk-custom".to_string()));
        assert_eq!(resolve_key_alias("notexist", &keys), None);
    }

    #[test]
    fn get_key_target_files_scans_cc_projects() {
        let dir = fixture_dir();
        let home = dir.join("home");
        fs::create_dir_all(home.join(".claude")).unwrap();
        fs::create_dir_all(home.join("AppData").join("Roaming").join("Code").join("User")).unwrap();
        fs::create_dir_all(home.join("Documents").join("LazyGit")).unwrap();
        fs::write(home.join(".claude").join("settings.json"), "{}").unwrap();
        fs::write(
            home.join("AppData").join("Roaming").join("Code").join("User").join("settings.json"),
            "{}",
        )
        .unwrap();
        fs::write(home.join("Documents").join("NumDesGlobalKey.json"), "{}").unwrap();
        fs::write(home.join("Documents").join("LazyGit").join("ai_commit.ps1"), "").unwrap();
        fs::create_dir_all(home.join("CCglm").join(".claude")).unwrap();
        fs::write(home.join("CCglm").join(".claude").join("settings.json"), "{}").unwrap();
        fs::create_dir_all(home.join("CCKimi").join(".claude")).unwrap();
        fs::write(home.join("CCKimi").join(".claude").join("settings.json"), "{}").unwrap();
        fs::create_dir_all(home.join("CCNoSettings")).unwrap();

        let targets = get_key_target_files(&home);
        assert_eq!(targets.len(), 6);
        assert!(targets.iter().any(|t| t.to_string_lossy().contains("CCglm")));
        assert!(targets.iter().any(|t| t.to_string_lossy().contains("CCKimi")));
        assert!(!targets.iter().any(|t| t.to_string_lossy().contains("CCNoSettings")));
        fs::remove_dir_all(&dir).ok();
    }
}
