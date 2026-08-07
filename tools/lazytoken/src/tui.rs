//! ratatui 全屏 TUI：token 列表 → 操作菜单 → 结果页。
//! 键盘：↑↓ 移动 · Enter 确认 · c 复制 · t 测试 · d 删除 · Esc 退出。

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
    List,
    ActionMenu,
    AddToken,
    ShowResult,
    ConfirmDelete,
    Quit,
}

pub fn run_interactive(tokens: &mut Vec<engine::TokenDef>) -> io::Result<()> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let result = run_core(&mut terminal, tokens);

    disable_raw_mode()?;
    execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
    terminal.show_cursor()?;
    result
}

fn run_core(
    terminal: &mut Terminal<CrosstermBackend<io::Stdout>>,
    tokens: &mut Vec<engine::TokenDef>,
) -> io::Result<()> {
    let mut stage = Stage::List;
    let mut sel = 0usize;
    let mut result_msg = String::new();
    let mut result_type = String::new(); // "ok" or "err"
    let _delete_confirm = false;
    let mut add_label = String::new();
    let mut add_token = String::new();
    let mut add_notes = String::new();
    let mut add_step = 0; // 0=label, 1=token, 2=notes, 3=save

    while !matches!(stage, Stage::Quit) {
        terminal.draw(|f| {
            let size = f.size();
            match &stage {
                Stage::List => {
                    if tokens.is_empty() {
                        let para = Paragraph::new(vec![
                            Line::from("暂无 token"),
                            Line::from(""),
                            Line::from("按 a 添加新的 token"),
                            Line::from("按 Esc 退出"),
                        ])
                        .block(
                            Block::default()
                                .borders(Borders::ALL)
                                .title(" lazytoken · 平台 JWT 管理 ")
                                .border_style(Style::default().fg(Color::Gray)),
                        )
                        .style(Style::default().fg(Color::DarkGray));
                        f.render_widget(para, size);
                    } else {
                        let mut state = TableState::default();
                        state.select(Some(sel));
                        let rows: Vec<Row> = tokens
                            .iter()
                            .map(|t| {
                                let info = engine::parse_jwt(&t.token);
                                let status = match &info {
                                    Some(i) if i.expired => {
                                        Span::styled("❌", Style::default().fg(Color::Red))
                                    }
                                    Some(_) => {
                                        Span::styled("✅", Style::default().fg(Color::Green))
                                    }
                                    None => Span::styled("⚠️", Style::default().fg(Color::Yellow)),
                                };
                                let user_id =
                                    info.as_ref().map(|i| i.user_id.as_str()).unwrap_or("?");
                                let exp = info.as_ref().map(|i| i.exp_str.as_str()).unwrap_or("?");
                                let expires_in =
                                    info.as_ref().map(|i| i.expires_in.as_str()).unwrap_or("?");
                                Row::new(vec![
                                    Cell::from(status),
                                    Cell::from(t.label.clone()),
                                    Cell::from(format!("#{}", user_id)),
                                    Cell::from(engine::mask_token(&t.token)),
                                    Cell::from(exp.to_string()),
                                    Cell::from(expires_in.to_string()),
                                ])
                            })
                            .collect();
                        let table = Table::new(
                            rows,
                            [
                                Constraint::Length(4),
                                Constraint::Length(14),
                                Constraint::Length(6),
                                Constraint::Length(24),
                                Constraint::Length(18),
                                Constraint::Length(12),
                            ],
                        )
                        .block(
                            Block::default()
                                .borders(Borders::ALL)
                                .title(" lazytoken · 平台 JWT 管理 ")
                                .border_style(Style::default().fg(Color::Gray)),
                        )
                        .header(
                            Row::new(vec![
                                Cell::from(" "),
                                Cell::from("名称"),
                                Cell::from("用户"),
                                Cell::from("Token"),
                                Cell::from("过期时间"),
                                Cell::from("剩余"),
                            ])
                            .style(
                                Style::default()
                                    .fg(Color::Cyan)
                                    .add_modifier(Modifier::BOLD),
                            ),
                        )
                        .highlight_style(
                            Style::default()
                                .fg(Color::Black)
                                .bg(Color::Cyan)
                                .add_modifier(Modifier::BOLD),
                        )
                        .highlight_symbol("▶ ");
                        f.render_stateful_widget(table, size, &mut state);
                        render_footer(f, "↑↓ 移动 · Enter 操作 · a 添加 · Esc 退出");
                    }
                }
                Stage::ActionMenu => {
                    let actions = vec![
                        ("c", "复制 token 到剪贴板"),
                        ("t", "测试 token 有效性"),
                        ("d", "删除此 token"),
                        ("Esc", "返回"),
                    ];
                    let rows: Vec<Row> = actions
                        .iter()
                        .map(|(key, desc)| Row::new(vec![Cell::from(*key), Cell::from(*desc)]))
                        .collect();
                    let table =
                        Table::new(rows, [Constraint::Length(8), Constraint::Percentage(90)])
                            .block(
                                Block::default()
                                    .borders(Borders::ALL)
                                    .title(format!(" 操作 - {}", tokens[sel].label))
                                    .border_style(Style::default().fg(Color::Gray)),
                            )
                            .highlight_style(
                                Style::default()
                                    .fg(Color::Black)
                                    .bg(Color::Cyan)
                                    .add_modifier(Modifier::BOLD),
                            )
                            .highlight_symbol("▶ ");
                    let mut state = TableState::default();
                    state.select(Some(0));
                    f.render_stateful_widget(table, size, &mut state);
                    render_footer(f, "按快捷键执行操作 · Esc 返回");
                }
                Stage::ShowResult => {
                    let style = if result_type == "ok" {
                        Style::default().fg(Color::Green)
                    } else {
                        Style::default().fg(Color::Red)
                    };
                    let para = Paragraph::new(vec![
                        Line::from(Span::styled(result_msg.clone(), style)),
                        Line::from(""),
                        Line::from("按 Enter 返回 · Esc 退出"),
                    ])
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title(" 结果 ")
                            .border_style(Style::default().fg(Color::Gray)),
                    );
                    f.render_widget(para, size);
                }
                Stage::ConfirmDelete => {
                    let para = Paragraph::new(vec![
                        Line::from(format!("确定删除 token: {} ?", tokens[sel].label)),
                        Line::from(""),
                        Line::from("y 确认删除 · Esc 取消"),
                    ])
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title(" 确认删除 ")
                            .border_style(Style::default().fg(Color::Red)),
                    );
                    f.render_widget(para, size);
                }
                Stage::AddToken => {
                    let prompt = match add_step {
                        0 => "输入 token 名称（标签）:",
                        1 => "输入 JWT token:",
                        2 => "输入备注（可选，直接 Enter 跳过）:",
                        _ => "",
                    };
                    let input = match add_step {
                        0 => add_label.clone(),
                        1 => engine::mask_token(&add_token),
                        2 => add_notes.clone(),
                        _ => String::new(),
                    };
                    let area = Layout::default()
                        .direction(Direction::Vertical)
                        .constraints([Constraint::Length(3), Constraint::Min(1)])
                        .split(size);
                    let input_block = Paragraph::new(input.as_str())
                        .block(
                            Block::default()
                                .borders(Borders::ALL)
                                .title(format!(" 添加 token ({}/3) · {}", add_step + 1, prompt))
                                .border_style(Style::default().fg(Color::Cyan)),
                        )
                        .style(Style::default().fg(Color::White));
                    f.render_widget(input_block, area[0]);
                    f.set_cursor(area[0].x + input.len() as u16 + 1, area[0].y + 1);
                    render_footer(f, "Enter 确认 · Esc 取消");
                }
                Stage::Quit => {}
            }
        })?;

        match event::read()? {
            Event::Paste(text) => {
                if matches!(stage, Stage::AddToken) {
                    let cleaned = text.replace(['\r', '\n'], "");
                    match add_step {
                        0 => add_label.push_str(&cleaned),
                        1 => add_token.push_str(&cleaned),
                        2 => add_notes.push_str(&cleaned),
                        _ => {}
                    }
                }
                continue;
            }
            Event::Key(key) => {
                if key.kind != KeyEventKind::Press {
                    continue;
                }
                // Ctrl+C 全局退出
                if key.modifiers.contains(event::KeyModifiers::CONTROL)
                    && matches!(key.code, KeyCode::Char('c'))
                {
                    stage = Stage::Quit;
                    continue;
                }

                match &stage {
                    Stage::List => match key.code {
                        KeyCode::Up => {
                            if !tokens.is_empty() {
                                sel = (sel + tokens.len() - 1) % tokens.len()
                            }
                        }
                        KeyCode::Down => {
                            if !tokens.is_empty() {
                                sel = (sel + 1) % tokens.len()
                            }
                        }
                        KeyCode::Esc => stage = Stage::Quit,
                        KeyCode::Enter => {
                            if !tokens.is_empty() {
                                stage = Stage::ActionMenu;
                            }
                        }
                        KeyCode::Char('a') => {
                            add_label.clear();
                            add_token.clear();
                            add_notes.clear();
                            add_step = 0;
                            stage = Stage::AddToken;
                        }
                        _ => {}
                    },
                    Stage::ActionMenu => match key.code {
                        KeyCode::Esc => stage = Stage::List,
                        KeyCode::Char('c') => {
                            match engine::copy_to_clipboard(&tokens[sel].token) {
                                Ok(_) => {
                                    result_msg =
                                        format!("✅ 已复制 token 到剪贴板: {}", tokens[sel].label);
                                    result_type = "ok".to_string();
                                }
                                Err(e) => {
                                    result_msg = format!("❌ 复制失败: {}", e);
                                    result_type = "err".to_string();
                                }
                            }
                            stage = Stage::ShowResult;
                        }
                        KeyCode::Char('t') => {
                            // Can't easily do async in TUI, so just show instruction
                            result_msg =
                                format!("直接在终端运行: lazytoken test {}", tokens[sel].label);
                            result_type = "ok".to_string();
                            stage = Stage::ShowResult;
                        }
                        KeyCode::Char('d') => {
                            stage = Stage::ConfirmDelete;
                        }
                        _ => {}
                    },
                    Stage::ShowResult => match key.code {
                        KeyCode::Esc => stage = Stage::Quit,
                        KeyCode::Enter => stage = Stage::List,
                        _ => {}
                    },
                    Stage::ConfirmDelete => match key.code {
                        KeyCode::Esc => stage = Stage::ActionMenu,
                        KeyCode::Char('y') | KeyCode::Char('Y') => {
                            let removed = tokens.remove(sel);
                            if !tokens.is_empty() && sel >= tokens.len() {
                                sel = tokens.len() - 1;
                            }
                            let _ = engine::save_tokens(tokens);
                            result_msg = format!("✅ 已删除: {}", removed.label);
                            result_type = "ok".to_string();
                            stage = Stage::ShowResult;
                        }
                        _ => {}
                    },
                    Stage::AddToken => match key.code {
                        KeyCode::Esc => stage = Stage::List,
                        KeyCode::Enter => match add_step {
                            0 => {
                                if !add_label.trim().is_empty() {
                                    add_step = 1;
                                }
                            }
                            1 => {
                                if !add_token.trim().is_empty() {
                                    add_step = 2;
                                }
                            }
                            2 => {
                                tokens.push(engine::TokenDef {
                                    label: add_label.trim().to_string(),
                                    token: add_token.trim().to_string(),
                                    notes: add_notes.trim().to_string(),
                                });
                                let _ = engine::save_tokens(tokens);
                                result_msg = format!("✅ 已添加 token: {}", add_label);
                                result_type = "ok".to_string();
                                stage = Stage::ShowResult;
                            }
                            _ => {}
                        },
                        KeyCode::Backspace => match add_step {
                            0 => {
                                add_label.pop();
                            }
                            1 => {
                                add_token.pop();
                            }
                            2 => {
                                add_notes.pop();
                            }
                            _ => {}
                        },
                        KeyCode::Char(c) => {
                            // Ctrl+V handled via bracketed paste, but also check here
                            if key.modifiers.contains(event::KeyModifiers::CONTROL) && c == 'v' {
                                // Try clipboard
                                if let Ok(clip) = read_clipboard() {
                                    let cleaned = clip.replace(['\r', '\n'], "");
                                    match add_step {
                                        0 => add_label.push_str(&cleaned),
                                        1 => add_token.push_str(&cleaned),
                                        2 => add_notes.push_str(&cleaned),
                                        _ => {}
                                    }
                                }
                            } else if !key.modifiers.contains(event::KeyModifiers::CONTROL) {
                                match add_step {
                                    0 => add_label.push(c),
                                    1 => add_token.push(c),
                                    2 => add_notes.push(c),
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    },
                    Stage::Quit => {}
                }
            }
            _ => {}
        }
    }
    Ok(())
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

fn read_clipboard() -> std::io::Result<String> {
    let out = std::process::Command::new("pwsh")
        .args(["-NoProfile", "-Command", "Get-Clipboard"])
        .output()?;
    Ok(String::from_utf8_lossy(&out.stdout).trim_end().to_string())
}
