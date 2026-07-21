//! ratatui 全屏 TUI：key 面板 → 文件面板 → 结果页。
//! 键盘：↑↓ 移动 · 空格 勾选 · a 全选/全不选 · Enter 确认 · Esc 返回/退出。

use crate::engine;
use crossterm::{
    event::{self, Event, KeyCode, KeyEventKind},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table, TableState},
    Terminal,
};
use std::io;

enum Stage {
    PickKey,
    CustomInput,
    PickFiles,
    ShowResult,
    Quit,
}

pub fn run_interactive(keys: &[engine::KeyDef]) -> io::Result<()> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let result = run_core(&mut terminal, keys);

    disable_raw_mode()?;
    execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
    terminal.show_cursor()?;
    result
}

fn run_core(terminal: &mut Terminal<CrosstermBackend<io::Stdout>>, keys: &[engine::KeyDef]) -> io::Result<()> {
    let mut stage = Stage::PickKey;
    let mut key_sel = 0usize;
    let key_count = keys.len() + 1; // + 自定义
    let mut file_sel = 0usize;
    let mut files: Vec<std::path::PathBuf> = Vec::new();
    let mut cur_labels: Vec<String> = Vec::new();
    let mut checked: Vec<bool> = Vec::new();
    let mut new_key = String::new();
    let mut new_label = String::new();
    let mut result_lines: Vec<(String, String)> = Vec::new(); // (status, file)
    let mut result_summary = String::new();
    let mut custom_input = String::new(); // 自定义 key 输入缓冲

    while !matches!(stage, Stage::Quit) {
        terminal.draw(|f| {
            let size = f.size();
            match stage {
                Stage::PickKey => {
                    let mut state = TableState::default();
                    state.select(Some(key_sel));
                    let rows: Vec<Row> = keys
                        .iter()
                        .map(|d| {
                            Row::new(vec![
                                Cell::from(d.label.clone()),
                                Cell::from(d.key.clone()),
                            ])
                        })
                        .chain(std::iter::once(Row::new(vec![
                            Cell::from("自定义输入"),
                            Cell::from("sk-..."),
                        ])))
                        .collect();
                    let table = Table::new(
                        rows,
                        [Constraint::Percentage(55), Constraint::Percentage(45)],
                    )
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title(" LiteLLM Key 切换 · 第 1 步 · 选目标 key ")
                            .border_style(Style::default().fg(Color::Gray)),
                    )
                    .highlight_style(
                        Style::default()
                            .fg(Color::Black)
                            .bg(Color::Cyan)
                            .add_modifier(Modifier::BOLD),
                    )
                    .highlight_symbol("▶ ");
                    f.render_stateful_widget(table, size, &mut state);
                    render_footer(f, "↑↓ 移动 · Enter 确认 · Esc 退出");
                }
                Stage::PickFiles => {
                    let mut state = TableState::default();
                    state.select(Some(file_sel));
                    let rows: Vec<Row> = files
                        .iter()
                        .enumerate()
                        .map(|(i, path)| {
                            let mark = if checked[i] { "[x]" } else { "[ ]" };
                            let cur = if cur_labels[i] == new_label {
                                Span::styled(cur_labels[i].clone(), Style::default().fg(Color::DarkGray))
                            } else {
                                Span::styled(cur_labels[i].clone(), Style::default().fg(Color::Yellow))
                            };
                            Row::new(vec![
                                Cell::from(mark),
                                Cell::from(path.display().to_string()),
                                Cell::from(cur),
                            ])
                        })
                        .collect();
                    let table = Table::new(
                        rows,
                        [
                            Constraint::Length(4),
                            Constraint::Percentage(65),
                            Constraint::Percentage(30),
                        ],
                    )
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title(format!(" 第 2 步 · 空格勾选 → {} ", new_label))
                            .border_style(Style::default().fg(Color::Gray)),
                    )
                    .highlight_style(
                        Style::default()
                            .fg(Color::Black)
                            .bg(Color::Cyan)
                            .add_modifier(Modifier::BOLD),
                    )
                    .highlight_symbol("▶ ");
                    f.render_stateful_widget(table, size, &mut state);
                    let pending = checked.iter().filter(|c| **c).count();
                    render_footer(
                        f,
                        &format!(
                            "↑↓ 移动 · 空格 勾选 · a 全选/全不选 · Enter 执行 · Esc 返回   已选 {}/{}",
                            pending,
                            files.len()
                        ),
                    );
                }
                Stage::ShowResult => {
                    let rows: Vec<Row> = result_lines
                        .iter()
                        .map(|(status, file)| {
                            let style = if status.starts_with('✓') {
                                Style::default().fg(Color::Green)
                            } else {
                                Style::default().fg(Color::DarkGray)
                            };
                            Row::new(vec![
                                Cell::from(Span::styled(status.clone(), style)),
                                Cell::from(file.clone()),
                            ])
                        })
                        .collect();
                    let table = Table::new(
                        rows,
                        [Constraint::Length(8), Constraint::Percentage(90)],
                    )
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title(" 执行结果 ")
                            .border_style(Style::default().fg(Color::Gray)),
                    );
                    let area = Layout::default()
                        .direction(Direction::Vertical)
                        .constraints([Constraint::Min(3), Constraint::Length(3)])
                        .split(size);
                    f.render_widget(table, area[0]);
                    let summary = Paragraph::new(vec![
                        Line::from(result_summary.clone()),
                        Line::from("Enter 再切一轮 · Esc 退出"),
                    ])
                    .style(Style::default().fg(Color::Gray));
                    f.render_widget(summary, area[1]);
                }
                Stage::CustomInput => {
                    let area = Layout::default()
                        .direction(Direction::Vertical)
                        .constraints([Constraint::Length(3), Constraint::Min(1)])
                        .split(size);
                    let input_block = Paragraph::new(custom_input.as_str())
                        .block(
                            Block::default()
                                .borders(Borders::ALL)
                                .title(" 自定义输入完整 key (sk-...) · 支持粘贴 ")
                                .border_style(Style::default().fg(Color::Cyan)),
                        )
                        .style(Style::default().fg(Color::White));
                    f.render_widget(input_block, area[0]);
                    // 光标定位到输入末尾
                    f.set_cursor(
                        area[0].x + custom_input.len() as u16 + 1,
                        area[0].y + 1,
                    );
                    render_footer(f, "Enter 确认 · Esc 返回 · 支持 Ctrl+V / 右键粘贴");
                }
                Stage::Quit => {}
            }
        })?;

        match event::read()? {
            Event::Paste(text) => {
                // bracketed paste（Windows Terminal / 现代终端默认开启）
                if matches!(stage, Stage::CustomInput) {
                    custom_input.push_str(&text.replace(['\r', '\n'], ""));
                }
                continue;
            }
            Event::Key(key) => {
                if key.kind != KeyEventKind::Press {
                    continue;
                }
                handle_key(
                    key,
                    &mut stage,
                    &mut key_sel,
                    key_count,
                    &mut file_sel,
                    &mut files,
                    &mut cur_labels,
                    &mut checked,
                    &mut new_key,
                    &mut new_label,
                    &mut result_lines,
                    &mut result_summary,
                    &mut custom_input,
                    keys,
                );
            }
            _ => {}
        }
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn handle_key(
    key: crossterm::event::KeyEvent,
    stage: &mut Stage,
    key_sel: &mut usize,
    key_count: usize,
    file_sel: &mut usize,
    files: &mut Vec<std::path::PathBuf>,
    cur_labels: &mut Vec<String>,
    checked: &mut Vec<bool>,
    new_key: &mut String,
    new_label: &mut String,
    result_lines: &mut Vec<(String, String)>,
    result_summary: &mut String,
    custom_input: &mut String,
    keys: &[engine::KeyDef],
) {
    // Ctrl+C 全局退出（raw mode 下 Ctrl+C 不发 SIGINT，发键盘事件，要显式处理）
    if key.modifiers.contains(event::KeyModifiers::CONTROL)
        && matches!(key.code, KeyCode::Char('c'))
    {
        *stage = Stage::Quit;
        return;
    }

    match stage {
        Stage::PickKey => match key.code {
            KeyCode::Up => *key_sel = (*key_sel + key_count - 1) % key_count,
            KeyCode::Down => *key_sel = (*key_sel + 1) % key_count,
            KeyCode::Esc => *stage = Stage::Quit,
            KeyCode::Enter => {
                if *key_sel == keys.len() {
                    custom_input.clear();
                    *stage = Stage::CustomInput;
                } else {
                    *new_key = keys[*key_sel].key.clone();
                    *new_label = keys[*key_sel].label.clone();
                    load_files(files, cur_labels, checked, keys);
                    *file_sel = 0;
                    *stage = Stage::PickFiles;
                }
            }
            _ => {}
        },
        Stage::CustomInput => match key.code {
            KeyCode::Esc => *stage = Stage::PickKey,
            KeyCode::Enter => {
                let trimmed = custom_input.trim();
                if trimmed.starts_with("sk-") && trimmed.len() > 10 {
                    *new_key = trimmed.to_string();
                    *new_label = "(自定义)".to_string();
                    load_files(files, cur_labels, checked, keys);
                    *file_sel = 0;
                    *stage = Stage::PickFiles;
                }
                // 非法输入：留在 CustomInput 继续编辑
            }
            KeyCode::Backspace => {
                custom_input.pop();
            }
            KeyCode::Char(c) => {
                // Ctrl+V 兜底（部分终端不发 bracketed paste，发 Ctrl+V）
                if key.modifiers.contains(event::KeyModifiers::CONTROL) && c == 'v' {
                    if let Ok(clip) = read_clipboard() {
                        custom_input.push_str(&clip.replace(['\r', '\n'], ""));
                    }
                } else if !key.modifiers.contains(event::KeyModifiers::CONTROL) {
                    custom_input.push(c);
                }
            }
            _ => {}
        },
        Stage::PickFiles => match key.code {
            KeyCode::Up => *file_sel = (*file_sel + files.len().max(1) - 1) % files.len().max(1),
            KeyCode::Down => *file_sel = (*file_sel + 1) % files.len().max(1),
            KeyCode::Char(' ') => checked[*file_sel] = !checked[*file_sel],
            KeyCode::Char('a') | KeyCode::Char('A') => {
                let any_unchecked = checked.iter().any(|c| !*c);
                checked.fill(any_unchecked);
            }
            KeyCode::Esc => *stage = Stage::PickKey,
            KeyCode::Enter => {
                let chosen: Vec<_> = files
                    .iter()
                    .zip(checked.iter())
                    .filter(|(_, c)| **c)
                    .map(|(f, _)| f.clone())
                    .collect();
                if !chosen.is_empty() {
                    let all: Vec<&str> = engine::all_key_values(keys);
                    let (changed, skipped) = engine::switch_files_to_key(new_key, &chosen, &all);
                    *result_lines = changed
                        .iter()
                        .map(|f| ("✓ 已切".to_string(), f.display().to_string()))
                        .chain(skipped.iter().map(|s| ("- 跳过".to_string(), s.clone())))
                        .collect();
                    *result_summary = format!("完成：{} 个文件切到 {}", changed.len(), new_label);
                    *stage = Stage::ShowResult;
                }
            }
            _ => {}
        },
        Stage::ShowResult => match key.code {
            KeyCode::Esc => *stage = Stage::Quit,
            KeyCode::Enter => {
                *stage = Stage::PickKey;
                *key_sel = 0;
            }
            _ => {}
        },
        Stage::Quit => {}
    }
}

fn load_files(
    files: &mut Vec<std::path::PathBuf>,
    cur_labels: &mut Vec<String>,
    checked: &mut Vec<bool>,
    keys: &[engine::KeyDef],
) {
    let home = dirs_home();
    *files = engine::get_key_target_files(&home);
    let all: Vec<&str> = engine::all_key_values(keys);
    *cur_labels = files
        .iter()
        .map(|f| {
            engine::find_file_key(f, &all)
                .and_then(|k| engine::label_of(&k, keys).map(|s| s.to_string()))
                .unwrap_or_else(|| "(无已知 key)".to_string())
        })
        .collect();
    *checked = vec![true; files.len()];
}

/// 读系统剪贴板（Windows）：PowerShell Get-Clipboard，比引 clipboard crate 少一层依赖
fn read_clipboard() -> std::io::Result<String> {
    let out = std::process::Command::new("pwsh")
        .args(["-NoProfile", "-Command", "Get-Clipboard"])
        .output()?;
    Ok(String::from_utf8_lossy(&out.stdout).trim_end().to_string())
}

fn render_footer(f: &mut ratatui::Frame, text: &str) {
    let size = f.size();
    let footer = Paragraph::new(text).style(Style::default().fg(Color::DarkGray));
    let area = ratatui::layout::Rect {
        x: size.x,
        y: size.height.saturating_sub(1),
        width: size.width,
        height: 1,
    };
    f.render_widget(footer, area);
}

fn dirs_home() -> std::path::PathBuf {
    std::env::var("USERPROFILE")
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|_| std::path::PathBuf::from("."))
}
