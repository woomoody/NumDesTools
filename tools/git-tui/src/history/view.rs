use std::io;

use ratatui::{
    layout::{Constraint, Direction, Layout},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table},
};

use super::{terminal::Tui, App, Stage};

pub(super) fn draw(terminal: &mut Tui, app: &mut App) -> io::Result<()> {
    terminal.draw(|frame| {
        let areas = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(1),
                Constraint::Min(3),
                Constraint::Length(1),
                Constraint::Length(1),
            ])
            .split(frame.size());
        app.table_area = areas[1];
        frame.render_widget(render_title(app), areas[0]);

        let message_width = usize::from(frame.size().width.saturating_sub(58)).max(12);
        let rows = app.visible.iter().enumerate().map(|(position, index)| {
            let commit = &app.commits[*index];
            Row::new(vec![
                Cell::from(position.to_string()),
                Cell::from(commit.sha.clone()).style(Style::default().fg(Color::Yellow)),
                Cell::from(commit.date.clone()).style(Style::default().fg(Color::Magenta)),
                Cell::from(truncate(&commit.author, 18)).style(Style::default().fg(Color::Blue)),
                Cell::from(truncate(&commit.message, message_width)),
            ])
        });
        let table = Table::new(
            rows,
            [
                Constraint::Length(5),
                Constraint::Length(9),
                Constraint::Length(18),
                Constraint::Length(20),
                Constraint::Min(12),
            ],
        )
        .header(
            Row::new(["#", "sha", "日期", "作者", "message"])
                .style(Style::default().add_modifier(Modifier::BOLD))
                .bottom_margin(1),
        )
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(" file-history "),
        )
        .highlight_symbol("▶ ")
        .highlight_style(
            Style::default()
                .fg(Color::Black)
                .bg(Color::Cyan)
                .add_modifier(Modifier::BOLD),
        );
        frame.render_stateful_widget(table, areas[1], &mut app.table_state);

        let input = match app.stage {
            Stage::AuthorFilter => format!("作者筛选（实时 git log）: {}█", app.author_filter),
            Stage::Search => format!("搜索 message: {}█", app.search),
            Stage::CommitList | Stage::Quit => app.error.clone().unwrap_or_else(|| {
                app.first_compare_sha
                    .as_ref()
                    .map(|sha| format!("已选中 {sha}，请移动到另一个 commit 按 c 比对"))
                    .unwrap_or_default()
            }),
        };
        frame.render_widget(
            Paragraph::new(input).style(match app.stage {
                Stage::AuthorFilter => Style::default().fg(Color::Blue),
                Stage::Search => Style::default().fg(Color::Yellow),
                Stage::CommitList | Stage::Quit if app.error.is_some() => {
                    Style::default().fg(Color::Red)
                }
                Stage::CommitList | Stage::Quit => Style::default().fg(Color::Cyan),
            }),
            areas[2],
        );
        frame.render_widget(
            Paragraph::new(
                "↑↓ 移动 · Enter 与工作区比对 · c 比对两个提交 · a 作者筛选 · / 搜索 · q 退出",
            )
            .style(Style::default().fg(Color::DarkGray)),
            areas[3],
        );
    })?;
    Ok(())
}

fn render_title(app: &App) -> Paragraph<'static> {
    Paragraph::new(Line::from(vec![
        Span::styled(
            format!(" {} ", app.display_file),
            Style::default()
                .fg(Color::Cyan)
                .add_modifier(Modifier::BOLD),
        ),
        Span::raw(format!("· {} 个提交", app.visible.len())),
        Span::styled(
            format!("  作者: {}", value_or_all(&app.author_filter)),
            Style::default().fg(Color::Blue),
        ),
        Span::styled(
            format!("  搜索: {}", value_or_all(&app.search)),
            Style::default().fg(Color::Yellow),
        ),
    ]))
}

fn truncate(value: &str, max: usize) -> String {
    if value.chars().count() <= max {
        value.to_string()
    } else {
        format!(
            "{}…",
            value
                .chars()
                .take(max.saturating_sub(1))
                .collect::<String>()
        )
    }
}

fn value_or_all(value: &str) -> &str {
    if value.is_empty() {
        "全部"
    } else {
        value
    }
}
