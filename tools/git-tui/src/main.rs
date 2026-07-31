mod history;

use std::{env, io, path::PathBuf, process::Command};

fn main() {
    if let Err(error) = run() {
        eprintln!("git-tui: {error}");
        std::process::exit(1);
    }
}

fn run() -> io::Result<()> {
    configure_utf8_console();
    let mut args = env::args_os().skip(1);
    match args
        .next()
        .and_then(|value| value.into_string().ok())
        .as_deref()
    {
        Some("history") => {
            let file = args.next().map(PathBuf::from).ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "用法: git-tui history <file>")
            })?;
            if args.next().is_some() {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "history 只接受一个文件路径",
                ));
            }
            history::run(&file)
        }
        Some(command) => Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("未知子命令: {command}\n用法: git-tui history <file>"),
        )),
        None => Err(io::Error::new(
            io::ErrorKind::InvalidInput,
            "用法: git-tui history <file>",
        )),
    }
}

fn configure_utf8_console() {
    env::set_var("LESSCHARSET", "utf-8");
    #[cfg(windows)]
    {
        let _ = Command::new("cmd")
            .args(["/d", "/c", "chcp 65001 > nul"])
            .status();
    }
}
