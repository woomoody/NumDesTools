use std::{
    io::{self, Read, Write},
    path::{Path, PathBuf},
    process::{Command, Stdio},
};

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct Commit {
    pub(super) sha: String,
    pub(super) date: String,
    pub(super) author: String,
    pub(super) message: String,
}

pub(super) fn resolve_file(file: &Path) -> io::Result<(PathBuf, PathBuf)> {
    let absolute = if file.is_absolute() {
        file.to_path_buf()
    } else {
        std::env::current_dir()?.join(file)
    }
    .canonicalize()?;
    let parent = absolute
        .parent()
        .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "文件路径没有父目录"))?;
    let output = Command::new("git")
        .current_dir(parent)
        .args(["rev-parse", "--show-toplevel"])
        .output()?;
    if !output.status.success() {
        return Err(command_error("无法定位 Git 仓库", &output.stderr));
    }
    let repo_root = PathBuf::from(String::from_utf8_lossy(&output.stdout).trim()).canonicalize()?;
    let relative = absolute
        .strip_prefix(&repo_root)
        .map(Path::to_path_buf)
        .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "文件不在当前 Git 仓库内"))?;
    Ok((repo_root, relative))
}

pub(super) fn load_commits(repo_root: &Path, file: &Path, author: &str) -> io::Result<Vec<Commit>> {
    let recent = run_git_log(repo_root, file, author, Some("1 month ago"), None)?;
    if recent.is_empty() {
        run_git_log(repo_root, file, author, None, Some(30))
    } else {
        Ok(recent)
    }
}

fn run_git_log(
    repo_root: &Path,
    file: &Path,
    author: &str,
    since: Option<&str>,
    limit: Option<usize>,
) -> io::Result<Vec<Commit>> {
    let mut command = Command::new("git");
    command.current_dir(repo_root).args([
        "log",
        "--follow",
        "--format=%h|%ad|%an|%s",
        "--date=format:%Y-%m-%d %H:%M",
    ]);
    if let Some(value) = since {
        command.arg(format!("--since={value}"));
    }
    if let Some(value) = limit {
        command.arg(format!("-{value}"));
    }
    if !author.is_empty() {
        command.arg(format!("--author={author}"));
    }
    let output = command.arg("--").arg(file).output()?;
    if !output.status.success() {
        return Err(command_error("git log 失败", &output.stderr));
    }
    Ok(String::from_utf8_lossy(&output.stdout)
        .lines()
        .filter_map(parse_log_line)
        .collect())
}

pub(super) fn pipe_diff_to_delta(repo_root: &Path, file: &Path, sha: &str) -> io::Result<()> {
    let mut command = Command::new("git");
    command
        .current_dir(repo_root)
        .args(["diff", sha, "--"])
        .arg(file);
    pipe_git_to_delta_and_less(&mut command)
}

pub(super) fn pipe_diff_two_to_delta(
    repo_root: &Path,
    file: &Path,
    sha1: &str,
    sha2: &str,
) -> io::Result<()> {
    let mut command = Command::new("git");
    command
        .current_dir(repo_root)
        .args(["diff", sha1, sha2, "--"])
        .arg(file);
    pipe_git_to_delta_and_less(&mut command)
}

fn pipe_git_to_delta_and_less(git_command: &mut Command) -> io::Result<()> {
    // git diff → delta → 临时文件（管道完整执行完才进 less，避免 less 提前 q 退出导致管道死锁）
    let tmp = std::env::temp_dir().join(format!("git-tui-diff-{}.txt", std::process::id()));
    {
        let mut git = git_command.stdout(Stdio::piped()).spawn()?;
        let git_stdout = git
            .stdout
            .take()
            .ok_or_else(|| io::Error::other("无法读取 git diff 输出"))?;
        let mut delta_status = Command::new("delta")
            .args(["--paging", "never"])
            .stdin(Stdio::from(git_stdout))
            .stdout(Stdio::piped())
            .spawn()
            .map_err(|error| {
                io::Error::new(
                    error.kind(),
                    format!("无法启动 delta（请确认已安装并在 PATH 中）: {error}"),
                )
            })?;
        let delta_stdout = delta_status
            .stdout
            .take()
            .ok_or_else(|| io::Error::other("无法读取 delta 输出"))?;
        let mut delta_child = delta_status;
        let mut delta_out = delta_stdout;
        let mut buf = Vec::new();
        delta_out.read_to_end(&mut buf)?;
        delta_child.wait()?;
        git.wait()?;

        // diff 输出为空（xlsx 二进制文件无 textconv 或无改动）→ 给提示而不是空屏
        if buf.is_empty() {
            std::fs::write(
                &tmp,
                "（无文本差异——xlsx 二进制文件可能需要 textconv 配置，或该提交无改动）\n",
            )?;
        } else {
            std::fs::write(&tmp, &buf)?;
        }
    }

    // less 读临时文件（less 退出后删临时文件）
    let less_path = find_less();
    let status = match less_path {
        Some(path) => Command::new(path).arg("-R").arg(&tmp).status(),
        None => Ok(std::process::ExitStatus::default()),
    };
    let _ = std::fs::remove_file(&tmp);

    match status {
        Ok(s) if s.success() => Ok(()),
        Ok(s) => Err(io::Error::other(format!("less 失败: {s}"))),
        Err(e) => Err(io::Error::new(e.kind(), format!("无法启动 less: {e}"))),
    }
}

/// 查找 less.exe：先 PATH，再 Git for Windows 常见路径。
fn find_less() -> Option<std::path::PathBuf> {
    if Command::new("less")
        .arg("--version")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .is_ok()
    {
        return Some(std::path::PathBuf::from("less"));
    }
    let candidates = [
        r"C:\Program Files\Git\usr\bin\less.exe",
        r"C:\Program Files (x86)\Git\usr\bin\less.exe",
    ];
    candidates
        .iter()
        .find(|p| std::path::Path::new(p).exists())
        .map(|p| p.into())
}

pub(super) fn parse_log_line(line: &str) -> Option<Commit> {
    let mut fields = line.splitn(4, '|');
    Some(Commit {
        sha: fields.next()?.to_string(),
        date: fields.next()?.to_string(),
        author: fields.next()?.to_string(),
        message: fields.next()?.to_string(),
    })
}

pub(super) fn filter_by_message(commits: &[Commit], query: &str) -> Vec<usize> {
    let needle = query.to_lowercase();
    commits
        .iter()
        .enumerate()
        .filter(|(_, commit)| needle.is_empty() || commit.message.to_lowercase().contains(&needle))
        .map(|(index, _)| index)
        .collect()
}

fn command_error(context: &str, stderr: &[u8]) -> io::Error {
    io::Error::other(format!(
        "{context}: {}",
        String::from_utf8_lossy(stderr).trim()
    ))
}

#[cfg(test)]
mod tests {
    use super::{filter_by_message, parse_log_line};

    #[test]
    fn parses_message_containing_separator_when_git_log_line_is_valid() {
        let line = "abc1234|2026-07-30 14:00|张三|fix: 保留 | 分隔符";
        let commit = parse_log_line(line).expect("valid git log line");
        assert_eq!(commit.sha, "abc1234");
        assert_eq!(commit.date, "2026-07-30 14:00");
        assert_eq!(commit.author, "张三");
        assert_eq!(commit.message, "fix: 保留 | 分隔符");
    }

    #[test]
    fn rejects_line_when_git_log_fields_are_missing() {
        assert!(parse_log_line("abc1234|2026-07-30 14:00|张三").is_none());
    }

    #[test]
    fn filters_messages_case_insensitively_when_query_is_present() {
        let commits = [
            parse_log_line("abc1234|2026-07-30 14:00|张三|Fix: Balloon").expect("fixture"),
            parse_log_line("def5678|2026-07-29 10:00|李四|feat: map").expect("fixture"),
        ];
        let matches = filter_by_message(&commits, "balloon");
        assert_eq!(matches.len(), 1);
        assert_eq!(commits[matches[0]].sha, "abc1234");
    }
}
