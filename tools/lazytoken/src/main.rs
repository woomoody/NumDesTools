mod engine;
mod proxy;
mod tui;

use std::env;
use std::io;

fn main() -> io::Result<()> {
    let mut tokens = engine::load_tokens();

    let args: Vec<String> = env::args().skip(1).collect();

    if let Some(cmd) = args.first() {
        match cmd.as_str() {
            "list" | "ls" => cmd_list(&tokens),
            "copy" | "cp" => cmd_copy(&tokens, args.get(1)),
            "test" => cmd_test(&tokens, args.get(1)),
            "add" => cmd_add(&mut tokens, args.get(1), args.get(2), args.get(3)),
            "rm" | "delete" => cmd_delete(&mut tokens, args.get(1)),
            "token" | "get" => cmd_token(&tokens, args.get(1)),
            "api" => cmd_api(
                &tokens,
                args.get(1),
                args.get(2),
                args.get(3).map(|s| s.as_str()),
            ),
            "serve" => cmd_serve(&tokens),
            "help" | "--help" | "-h" => cmd_help(),
            _ => {
                // Try to use as label to copy
                if let Some(_) = tokens.iter().find(|t| t.label == *cmd) {
                    cmd_copy(&tokens, Some(cmd))
                } else {
                    eprintln!("未知命令: {} （使用 lazytoken help 查看帮助）", cmd);
                    std::process::exit(1);
                }
            }
        }
    } else {
        tui::run_interactive(&mut tokens)
    }
}

fn cmd_list(tokens: &[engine::TokenDef]) -> io::Result<()> {
    if tokens.is_empty() {
        println!("暂无 token。使用 lazytoken add 添加。");
        return Ok(());
    }

    println!("{:─^80}", " lazytoken 列表 ");
    println!(
        "{:<4} {:<12} {:<8} {:<30} {:<18} {:<10}",
        "#", "名称", "用户", "Token", "过期时间", "剩余"
    );
    println!("{}", "─".repeat(80));

    for (i, t) in tokens.iter().enumerate() {
        let info = engine::parse_jwt(&t.token);
        let status = match &info {
            Some(i) if i.expired => "❌",
            Some(_) => "✅",
            None => "⚠️",
        };
        let uid = info.as_ref().map(|i| i.user_id.as_str()).unwrap_or("?");
        let exp = info.as_ref().map(|i| i.exp_str.as_str()).unwrap_or("?");
        let remain = info.as_ref().map(|i| i.expires_in.as_str()).unwrap_or("?");
        println!(
            "{:<4} {:<12} {:<8} {:<30} {:<18} {:<10}",
            format!("{}{}", status, i + 1),
            t.label,
            format!("#{}", uid),
            engine::mask_token(&t.token),
            exp,
            remain
        );
    }
    println!("{}", "─".repeat(80));
    Ok(())
}

fn cmd_copy(tokens: &[engine::TokenDef], name: Option<&String>) -> io::Result<()> {
    let token = match name {
        Some(n) => tokens.iter().find(|t| t.label == *n).map(|t| &t.token),
        None => tokens.first().map(|t| &t.token),
    };

    match token {
        Some(t) => {
            engine::copy_to_clipboard(t)?;
            let label = tokens
                .iter()
                .find(|x| x.token == *t)
                .map(|x| x.label.as_str())
                .unwrap_or("(未命名)");
            println!("✅ 已复制 token 到剪贴板: {}", label);
            Ok(())
        }
        None => {
            eprintln!("没找到 token。使用 lazytoken add 添加。");
            std::process::exit(1);
        }
    }
}

fn cmd_test(tokens: &[engine::TokenDef], name: Option<&String>) -> io::Result<()> {
    let token = match name {
        Some(n) => tokens.iter().find(|t| t.label == *n).map(|t| &t.token),
        None => tokens.first().map(|t| &t.token),
    };

    match token {
        Some(t) => {
            println!("正在测试 token...");
            match engine::test_token(t) {
                Ok(msg) => println!("{}", msg),
                Err(msg) => println!("{}", msg),
            }
            Ok(())
        }
        None => {
            eprintln!("没找到 token。使用 lazytoken add 添加。");
            std::process::exit(1);
        }
    }
}

fn cmd_add(
    tokens: &mut Vec<engine::TokenDef>,
    label: Option<&String>,
    token: Option<&String>,
    notes: Option<&String>,
) -> io::Result<()> {
    let label = label.map(|s| s.as_str()).unwrap_or("");
    let token = token.map(|s| s.as_str()).unwrap_or("");

    if label.is_empty() || token.is_empty() {
        eprintln!("用法: lazytoken add <名称> <token> [备注]");
        eprintln!("示例: lazytoken add 平台主账号 eyJhbG...");
        std::process::exit(1);
    }

    tokens.push(engine::TokenDef {
        label: label.to_string(),
        token: token.to_string(),
        notes: notes.map(|s| s.to_string()).unwrap_or_default(),
    });
    engine::save_tokens(tokens)?;
    println!("✅ 已添加 token: {}", label);
    Ok(())
}

fn cmd_delete(tokens: &mut Vec<engine::TokenDef>, name: Option<&String>) -> io::Result<()> {
    let name = match name {
        Some(n) => n,
        None => {
            eprintln!("用法: lazytoken rm <名称>");
            std::process::exit(1);
        }
    };

    let idx = tokens.iter().position(|t| t.label == *name);
    match idx {
        Some(i) => {
            let removed = tokens.remove(i);
            engine::save_tokens(tokens)?;
            println!("✅ 已删除: {}", removed.label);
            Ok(())
        }
        None => {
            eprintln!("没找到 token: {}", name);
            std::process::exit(1);
        }
    }
}

/// 输出纯 token 字符串（供其他工具/脚本使用）
fn cmd_token(tokens: &[engine::TokenDef], name: Option<&String>) -> io::Result<()> {
    let token = match name {
        Some(n) => tokens.iter().find(|t| t.label == *n).map(|t| &t.token),
        None => tokens.first().map(|t| &t.token),
    };
    match token {
        Some(t) => {
            print!("{}", t);
            Ok(())
        }
        None => {
            eprintln!("没找到 token。使用 lazytoken add 添加。");
            std::process::exit(1);
        }
    }
}

/// 代理 API 调用：lazytoken api <方法> <路径> [请求体]
/// 示例: lazytoken api POST /api/chat '{"session_id":1047,"model":"gpt-5.4",...}'
fn cmd_api(
    tokens: &[engine::TokenDef],
    method: Option<&String>,
    path: Option<&String>,
    body: Option<&str>,
) -> io::Result<()> {
    let method = match method {
        Some(m) => m.to_uppercase(),
        None => {
            eprintln!("用法: lazytoken api <GET|POST|PUT|DELETE> <路径> [请求体]");
            std::process::exit(1);
        }
    };
    let path = match path {
        Some(p) => {
            if !p.starts_with('/') {
                format!("/api/{}", p)
            } else {
                p.clone()
            }
        }
        None => {
            eprintln!("用法: lazytoken api <GET|POST|PUT|DELETE> <路径> [请求体]");
            std::process::exit(1);
        }
    };

    let token = match engine::get_token(tokens, None) {
        Some(t) => t,
        None => {
            eprintln!("没找到 token。使用 lazytoken add 添加。");
            std::process::exit(1);
        }
    };

    let body_ref = body.filter(|s| !s.is_empty());

    match engine::proxy_api(token, &method, &path, body_ref) {
        Ok(resp) => {
            println!("{}", resp);
            Ok(())
        }
        Err(e) => {
            eprintln!("{}", e);
            std::process::exit(1);
        }
    }
}

fn cmd_help() -> io::Result<()> {
    println!("lazytoken - 平台 JWT token 管理工具");
    println!("");
    println!("用法:");
    println!("  lazytoken              TUI 交互模式");
    println!("  lazytoken list         列出所有 token");
    println!("  lazytoken copy [名称]   复制 token 到剪贴板");
    println!("  lazytoken test [名称]   测试 token 有效性");
    println!("  lazytoken add <名称> <token> [备注]  添加 token");
    println!("  lazytoken rm <名称>     删除 token");
    println!("  lazytoken <名称>        快捷复制（省略命令）");
    println!("  lazytoken token [名称]   输出 token 到 stdout（供脚本/工具使用）");
    println!("  lazytoken api <方法> <路径> [请求体]  代理 API 调用");
    println!("  lazytoken serve         启动本地 OpenAI 兼容文本代理");
    println!("");
    println!("API 代理示例:");
    println!("  lazytoken api GET /api/menus");
    println!("  lazytoken api GET /api/chat/sessions?user_id=149");
    println!("  lazytoken api POST /api/chat \"{{\\\"session_id\\\":1047,\\\"model\\\":\\\"gpt-5.4\\\"}}\"");
    println!("");
    println!("配置文件: %USERPROFILE%\\lazytoken.tokens.json");
    Ok(())
}

fn cmd_serve(tokens: &[engine::TokenDef]) -> io::Result<()> {
    proxy::serve(tokens)
}
