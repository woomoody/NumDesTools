mod engine;
mod tui;

use std::env;
use std::io;
use std::path::PathBuf;

fn main() -> io::Result<()> {
    let keys = engine::load_keys();
    if keys.is_empty() {
        eprintln!("没找到 key 配置文件：%USERPROFILE%\\lazykey.keys.json");
        eprintln!("格式：[{{\"label\":\"...\",\"alias\":\"...\",\"key\":\"sk-...\"}}, ...]");
        std::process::exit(1);
    }

    let args: Vec<String> = env::args().skip(1).collect();
    let key_name = args.first().filter(|a| !a.starts_with('-'));

    if let Some(name) = key_name {
        return run_direct(name, &keys);
    }

    tui::run_interactive(&keys)
}

/// 直切：lazykey cent / lazykey sleep / lazykey sk-...
fn run_direct(name: &str, keys: &[engine::KeyDef]) -> io::Result<()> {
    let new_key = match engine::resolve_key_alias(name, keys) {
        Some(k) => k,
        None => {
            eprintln!(
                "未知 key 别名: {}（可选：{}，或完整 sk-...）",
                name,
                keys.iter().map(|k| k.alias.as_str()).collect::<Vec<_>>().join(" / ")
            );
            std::process::exit(1);
        }
    };
    let label = engine::label_of(&new_key, keys).unwrap_or("(自定义)");
    let home = PathBuf::from(env::var("USERPROFILE").unwrap_or_else(|_| ".".to_string()));
    let files = engine::get_key_target_files(&home);
    let all = engine::all_key_values(keys);
    let (changed, skipped) = engine::switch_files_to_key(&new_key, &files, &all);

    for f in &changed {
        println!("✓ 已切：{}  {}", engine::label_for_path(f), f.display());
    }
    for s in &skipped {
        println!("- 跳过：{}", s);
    }
    println!("完成：{} 个文件切到 {}", changed.len(), label);
    Ok(())
}
