mod model;
mod tui;

use std::env;
use std::fs;
use std::io;
use std::path::PathBuf;

fn main() -> io::Result<()> {
    let args: Vec<String> = env::args().skip(1).collect();
    let diff_path = args.first().map(PathBuf::from).ok_or_else(|| {
        io::Error::new(io::ErrorKind::InvalidInput, "用法: conflict-tui <diff.json>")
    })?;

    let diff_json = fs::read_to_string(&diff_path)?;
    let diff: model::FileDiffDto = serde_json::from_str(&diff_json)
        .map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e))?;

    let result = tui::run_interactive(&diff)?;

    // 退出码：0=确认（result.json 已写）/ 1=放弃 / 2=错误
    match result {
        Some(selections) => {
            let result_path = diff_path.with_file_name(
                diff_path
                    .file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or("diff")
                    .replace("diff.json", "result.json"),
            );
            let json = serde_json::to_string(&selections)?;
            fs::write(&result_path, json)?;
            std::process::exit(0);
        }
        None => std::process::exit(1),
    }
}
