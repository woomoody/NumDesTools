//! ratatui 全屏 TUI：整表集中显示所有未选格，↑↓←→ 光标，o/t 选格，O/T 整行，Enter 确认，q 放弃。

use crate::model::{CellConflictDto, ConflictChoice, FileDiffDto, RowDiffType, SelectionDto, SelectionResultDto};
use crossterm::{
    event::{self, EnableMouseCapture, DisableMouseCapture, Event, KeyCode, KeyEventKind, MouseButton, MouseEventKind},
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
use similar::{ChangeTag, TextDiff};
use std::io;

/// 字符级 diff 段：(文本, kind) kind 0=相同 1=仅我方 2=仅对方。
fn char_diff(a: &str, b: &str) -> Vec<(String, u8)> {
    let diff = TextDiff::from_chars(a, b);
    let mut segs: Vec<(String, u8)> = Vec::new();
    for change in diff.iter_all_changes() {
        let kind = match change.tag() {
            ChangeTag::Equal => 0u8,
            ChangeTag::Delete => 1u8,
            ChangeTag::Insert => 2u8,
        };
        match segs.last_mut() {
            Some((text, k)) if *k == kind => text.push_str(change.value()),
            _ => segs.push((change.value().to_string(), kind)),
        }
    }
    segs
}

/// 围绕第一处差异截断并高亮：本方视角只显示自己独有的段，diff 段着色（我方红底/对方绿底）。
fn diff_snippet_spans(segs: &[(String, u8)], ours_view: bool, max: usize) -> Vec<Span<'static>> {
    let base_color = if ours_view { Color::Blue } else { Color::Yellow };
    let diff_style = if ours_view {
        Style::default().fg(Color::White).bg(Color::Red)
    } else {
        Style::default().fg(Color::Black).bg(Color::Green)
    };

    let first_diff = segs.iter().position(|(_, k)| *k != 0);
    let Some(first_diff) = first_diff else {
        let full: String = segs.iter().map(|(t, _)| t.as_str()).collect();
        return vec![Span::styled(truncate(&full, max), Style::default().fg(base_color))];
    };

    let pre: String = segs[..first_diff].iter().map(|(t, _)| t.as_str()).collect();
    let mut end_diff = first_diff;
    while end_diff < segs.len() && segs[end_diff].1 != 0 {
        end_diff += 1;
    }
    let want_kind = if ours_view { 1u8 } else { 2u8 };
    let diff_str: String = segs[first_diff..end_diff]
        .iter()
        .filter(|(_, k)| *k == want_kind)
        .map(|(t, _)| t.as_str())
        .collect();
    let post: String = segs[end_diff..]
        .iter()
        .take_while(|(_, k)| *k == 0)
        .map(|(t, _)| t.as_str())
        .collect();

    let ctx = (max / 3).max(4);
    let pre_show = if pre.chars().count() > ctx {
        format!("…{}", tail_chars(&pre, ctx))
    } else {
        pre
    };
    let post_show = if post.chars().count() > ctx {
        format!("{}…", &post.chars().take(ctx).collect::<String>())
    } else {
        post
    };

    let mut spans = vec![Span::styled(pre_show, Style::default().fg(base_color))];
    if !diff_str.is_empty() {
        spans.push(Span::styled(truncate(&diff_str, max / 2), diff_style));
    }
    spans.push(Span::styled(post_show, Style::default().fg(base_color)));
    spans
}

fn tail_chars(s: &str, n: usize) -> String {
    let chars: Vec<char> = s.chars().collect();
    let start = chars.len().saturating_sub(n);
    chars[start..].iter().collect()
}

/// 拍平所有未选格成一维列表（行=冲突格，列=我方/对方值）。
struct ConflictEntry {
    sheet_name: String,
    row_key: String,
    col_name: Option<String>, // None = 整行（OnlyOurs/OnlyTheirs）
    ours_display: String,
    theirs_display: String,
    remark: String,
    diff_segs: Vec<(String, u8)>,
    cell: CellConflictDto, // 借用原数据，选择时改这里
}

pub fn run_interactive(diff: &FileDiffDto) -> io::Result<Option<SelectionResultDto>> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    let result = run_core(&mut terminal, diff);

    disable_raw_mode()?;
    execute!(terminal.backend_mut(), DisableMouseCapture, LeaveAlternateScreen)?;
    terminal.show_cursor()?;
    result
}

fn run_core(
    terminal: &mut Terminal<CrosstermBackend<io::Stdout>>,
    diff: &FileDiffDto,
) -> io::Result<Option<SelectionResultDto>> {
    // 拍平所有未自动判定的冲突格
    let mut entries = flatten_unresolved(diff);
    if entries.is_empty() {
        return Ok(Some(SelectionResultDto {
            confirmed: true,
            selections: vec![],
        }));
    }

    let mut cursor_row = 0usize;
    let mut cursor_col = 0usize; // 0=我方值列, 1=对方值列
    let mut quit = false;
    let mut confirmed = false;
    let mut table_area = ratatui::layout::Rect::default(); // 记录整表渲染区域，鼠标点击换算行/列列用

    while !quit && !confirmed {
        terminal.draw(|f| {
            let size = f.size();
            table_area = size;
            let mut state = TableState::default();
            state.select(Some(cursor_row));

            let rows: Vec<Row> = entries
                .iter()
                .enumerate()
                .map(|(i, e)| {
                    let is_cursor = i == cursor_row;
                    let prefix = if is_cursor { "▶" } else { " " };
                    let choice_str = if !e.cell.is_explicit {
                        "未选(默认我方)"
                    } else if e.cell.choice == ConflictChoice::Ours {
                        "我方"
                    } else {
                        "对方"
                    };
                    let ours_cell = if e.cell.is_explicit && e.cell.choice == ConflictChoice::Ours {
                        Cell::from(Span::styled(
                            format!("{} ✓", truncate(&e.ours_display, 30)),
                            Style::default().fg(Color::Black).bg(Color::Blue).add_modifier(Modifier::BOLD),
                        ))
                    } else {
                        Cell::from(Line::from(diff_snippet_spans(&e.diff_segs, true, 30)))
                    };
                    let theirs_cell = if e.cell.is_explicit && e.cell.choice == ConflictChoice::Theirs {
                        Cell::from(Span::styled(
                            format!("{} ✓", truncate(&e.theirs_display, 30)),
                            Style::default().fg(Color::Black).bg(Color::Yellow).add_modifier(Modifier::BOLD),
                        ))
                    } else {
                        Cell::from(Line::from(diff_snippet_spans(&e.diff_segs, false, 30)))
                    };
                    let remark_short = truncate(&e.remark, 22);
                    Row::new(vec![
                        Cell::from(format!("{}{}", prefix, i + 1)),
                        Cell::from(e.row_key.clone()),
                        Cell::from(remark_short),
                        Cell::from(e.col_name.clone().unwrap_or_else(|| "(整行)".to_string())),
                        ours_cell,
                        theirs_cell,
                        Cell::from(choice_str),
                    ])
                })
                .collect();

            let table = Table::new(
                rows,
                [
                    Constraint::Length(4),
                    Constraint::Length(10),
                    Constraint::Length(22),
                    Constraint::Length(12),
                    Constraint::Percentage(30),
                    Constraint::Percentage(30),
                    Constraint::Length(14),
                ],
            )
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .title(format!(" {}/{} 个冲突格待选 ", entries.iter().filter(|e| !e.cell.is_explicit).count(), entries.len()))
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

            let footer = Paragraph::new("↑↓←→移动光标  s选当前列版本(默认我方)  O全选我方  T全选对方  鼠标点左/右半屏=选我方/对方  Enter确认  q放弃")
                .style(Style::default().fg(Color::DarkGray));
            let area = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Min(3), Constraint::Length(1)])
                .split(size);
            f.render_widget(footer, area[1]);
        })?;

        match event::read()? {
            Event::Mouse(m) if m.kind == MouseEventKind::Down(MouseButton::Left) => {
                // 表格无表头，顶部边框占 1 行，数据行从 table_area.y+1 开始；点左半屏=我方，右半屏=对方
                let row_idx = m.row.saturating_sub(table_area.y + 1) as usize;
                if row_idx < entries.len() {
                    cursor_row = row_idx;
                    cursor_col = if m.column < table_area.x + table_area.width / 2 { 0 } else { 1 };
                    entries[cursor_row].cell.choice = if cursor_col == 0 {
                        ConflictChoice::Ours
                    } else {
                        ConflictChoice::Theirs
                    };
                    entries[cursor_row].cell.is_explicit = true;
                }
            }
            Event::Key(key) => {
            if key.kind != KeyEventKind::Press {
                continue;
            }
            match key.code {
                KeyCode::Up => cursor_row = (cursor_row + entries.len() - 1) % entries.len(),
                KeyCode::Down => cursor_row = (cursor_row + 1) % entries.len(),
                KeyCode::Left => cursor_col = (cursor_col + 1) % 2,
                KeyCode::Right => cursor_col = (cursor_col + 1) % 2,
                KeyCode::Enter => {
                    // 未选格全用默认我方
                    for e in &mut entries {
                        if !e.cell.is_explicit {
                            e.cell.choice = ConflictChoice::Ours;
                            e.cell.is_explicit = true;
                        }
                    }
                    confirmed = true;
                }
                KeyCode::Char('s') => {
                    entries[cursor_row].cell.choice = if cursor_col == 0 {
                        ConflictChoice::Ours
                    } else {
                        ConflictChoice::Theirs
                    };
                    entries[cursor_row].cell.is_explicit = true;
                    if cursor_row < entries.len() - 1 {
                        cursor_row += 1;
                    }
                }
                KeyCode::Char('O') => {
                    for e in &mut entries {
                        e.cell.choice = ConflictChoice::Ours;
                        e.cell.is_explicit = true;
                    }
                }
                KeyCode::Char('T') => {
                    for e in &mut entries {
                        e.cell.choice = ConflictChoice::Theirs;
                        e.cell.is_explicit = true;
                    }
                }
                KeyCode::Char('q') => quit = true,
                _ => {}
            }
            }
            _ => {}
        }
    }

    if quit {
        return Ok(None);
    }

    // 生成 selections（只传变化量）
    let selections: Vec<SelectionDto> = entries
        .into_iter()
        .map(|e| SelectionDto {
            sheet_name: e.sheet_name,
            row_key: e.row_key,
            col_name: e.col_name,
            choice: e.cell.choice,
        })
        .collect();
    Ok(Some(SelectionResultDto {
        confirmed: true,
        selections,
    }))
}

fn flatten_unresolved(diff: &FileDiffDto) -> Vec<ConflictEntry> {
    let mut entries = Vec::new();
    for sheet in &diff.sheets {
        for row in &sheet.rows {
            if row.diff_type != RowDiffType::Modified {
                continue; // OnlyOurs/OnlyTheirs 默认已解决
            }
            for cell in &row.cells {
                if cell.is_explicit {
                    continue; // 已三方预选
                }
                let remark = row
                    .ours_full_row
                    .as_ref()
                    .or(row.theirs_full_row.as_ref())
                    .and_then(|r| {
                        r.iter()
                            .find(|(k, _)| k.starts_with('#'))
                            .and_then(|(_, v)| v.clone())
                    })
                    .unwrap_or_default();
                let ours_display = cell.ours_value.clone().unwrap_or_else(|| "(空)".to_string());
                let theirs_display = cell.theirs_value.clone().unwrap_or_else(|| "(空)".to_string());
                let diff_segs = char_diff(&ours_display, &theirs_display);
                entries.push(ConflictEntry {
                    sheet_name: sheet.sheet_name.clone(),
                    row_key: row.row_key.clone(),
                    col_name: Some(cell.col_name.clone()),
                    ours_display,
                    theirs_display,
                    remark,
                    diff_segs,
                    cell: cell.clone(),
                });
            }
        }
    }
    entries
}

fn truncate(s: &str, max: usize) -> String {
    if s.len() <= max {
        s.to_string()
    } else {
        format!("{}…", &s[..max])
    }
}
