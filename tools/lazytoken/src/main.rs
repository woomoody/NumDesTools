use chrono::{Datelike, Local};
use crossterm::{
    event::{self, Event, KeyCode, KeyEventKind},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout},
    style::{Color, Modifier, Style},
    widgets::{Block, Borders, Cell, Paragraph, Row, Table, TableState},
    Terminal,
};
use serde_json::Value;
use std::{fs, io, path::PathBuf, process::Command};

fn canonical_json() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(r"..\..\NumDesTools.Tests\lazytoken_stats.json")
}

fn run_canonical() -> io::Result<()> {
    let script = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join(r"..\..\NumDesTools.Tests\claude_token_stats.py");
    let python = std::env::var("PYTHON").unwrap_or_else(|_| "python".into());
    let mut command = Command::new(python);
    command.env("LAZYTOKEN_NO_BROWSER", "1");
    if command
        .arg(script)
        .arg("--date")
        .arg("today")
        .status()?
        .success()
    {
        Ok(())
    } else {
        Err(io::Error::other("claude_token_stats.py failed"))
    }
}

fn n(v: &Value, key: &str) -> u64 {
    v.get(key).and_then(Value::as_u64).unwrap_or(0)
}
fn money(v: &Value) -> f64 {
    v.get("cost").and_then(Value::as_f64).unwrap_or(0.0)
}
fn fmt(n: u64) -> String {
    let mut s = n.to_string();
    let mut i = s.len() as i32 - 3;
    while i > 0 {
        s.insert(i as usize, ',');
        i -= 3;
    }
    s
}

#[derive(Default, Clone)]
struct Metric {
    input: u64,
    output: u64,
    cache: u64,
    cost: f64,
}
impl Metric {
    fn add(&mut self, v: &Value) {
        self.input += n(v, "input");
        self.output += n(v, "output");
        self.cache += n(v, "cache_read") + n(v, "cache_write");
        self.cost += money(v)
    }
    fn total(&self) -> u64 {
        self.input + self.output + self.cache
    }
}

fn range(daily: &serde_json::Map<String, Value>, start: Option<String>, end: String) -> Metric {
    let mut m = Metric::default();
    for (date, v) in daily {
        if start.as_ref().map_or(true, |x| date >= x) && date <= &end {
            m.add(v)
        }
    }
    m
}
fn recent(daily: &serde_json::Map<String, Value>, days: i64) -> Metric {
    let today = Local::now().date_naive();
    range(
        daily,
        Some((today - chrono::Duration::days(days - 1)).to_string()),
        today.to_string(),
    )
}
fn month_now(daily: &serde_json::Map<String, Value>) -> Metric {
    let t = Local::now().date_naive();
    range(
        daily,
        Some(format!("{}-{:02}-01", t.year(), t.month())),
        t.to_string(),
    )
}
fn metric_row(label: &str, m: &Metric, selected: bool) -> Row<'static> {
    Row::new(vec![
        Cell::from(if selected { "▶" } else { " " }),
        Cell::from(label.to_string()),
        Cell::from(fmt(m.input)),
        Cell::from(fmt(m.output)),
        Cell::from(fmt(m.cache)),
        Cell::from(format!("${:.4}", m.cost)),
    ])
}
fn render_table(
    f: &mut ratatui::Frame,
    area: ratatui::layout::Rect,
    title: &str,
    rows: Vec<Row<'static>>,
    selected: usize,
) {
    let mut state = TableState::default();
    state.select(Some(selected.min(rows.len().saturating_sub(1))));
    let table = Table::new(
        rows,
        [
            Constraint::Length(2),
            Constraint::Min(24),
            Constraint::Length(16),
            Constraint::Length(16),
            Constraint::Length(16),
            Constraint::Length(16),
        ],
    )
    .header(
        Row::new(["", "项目", "输入", "输出", "缓存", "美元"]).style(
            Style::default()
                .fg(Color::Cyan)
                .add_modifier(Modifier::BOLD),
        ),
    )
    .block(Block::default().borders(Borders::ALL).title(title));
    f.render_stateful_widget(table, area, &mut state)
}

fn main() -> io::Result<()> {
    if std::env::args().any(|a| a == "--help" || a == "-h") {
        println!("lazytoken - bat canonical stats TUI\n仅保留：总览、自然月汇总\nTab/←→ 切换 · ↑↓ 浏览 · Q/Esc 退出");
        return Ok(());
    }
    run_canonical()?;
    let data: Value =
        serde_json::from_str(&fs::read_to_string(canonical_json())?).map_err(io::Error::other)?;
    enable_raw_mode()?;
    let mut out = io::stdout();
    execute!(out, EnterAlternateScreen)?;
    let mut terminal = Terminal::new(CrosstermBackend::new(out))?;
    let mut tab = 0usize;
    let mut selected = 0usize;
    let result = loop {
        terminal.draw(|f| draw(f, &data, tab, selected))?;
        if let Event::Key(key) = event::read()? {
            if key.kind != KeyEventKind::Press {
                continue;
            }
            match key.code {
                KeyCode::Char('q') | KeyCode::Esc => break Ok(()),
                KeyCode::Tab | KeyCode::Right => {
                    tab = (tab + 1) % 2;
                    selected = 0
                }
                KeyCode::BackTab | KeyCode::Left => {
                    tab = (tab + 1) % 2;
                    selected = 0
                }
                KeyCode::Down => selected = selected.saturating_add(1),
                KeyCode::Up => selected = selected.saturating_sub(1),
                _ => {}
            }
        }
    };
    disable_raw_mode()?;
    execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
    terminal.show_cursor()?;
    result
}

fn draw(f: &mut ratatui::Frame, data: &Value, tab: usize, selected: usize) {
    let areas = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(4),
            Constraint::Length(3),
            Constraint::Min(5),
            Constraint::Length(2),
        ])
        .split(f.size());
    let daily = data.get("daily").and_then(Value::as_object);
    let all = daily
        .map(|x| range(x, None, Local::now().date_naive().to_string()))
        .unwrap_or_default();
    let top = Table::new(
        vec![Row::new(vec![
            Cell::from("总 Token"),
            Cell::from(fmt(all.total())),
            Cell::from("美元估算"),
            Cell::from(format!("${:.4}", all.cost)),
        ])],
        [
            Constraint::Length(14),
            Constraint::Length(20),
            Constraint::Length(14),
            Constraint::Min(20),
        ],
    )
    .block(
        Block::default()
            .borders(Borders::ALL)
            .title("bat canonical · lazytoken"),
    );
    f.render_widget(top, areas[0]);
    f.render_widget(
        ratatui::widgets::Tabs::new(vec!["总览", "自然月汇总"])
            .select(tab)
            .highlight_style(Style::default().fg(Color::Black).bg(Color::Cyan))
            .block(Block::default().borders(Borders::ALL)),
        areas[1],
    );
    if tab == 0 {
        let empty = serde_json::Map::new();
        let ds = daily.unwrap_or(&empty);
        let rows = [
            ("本日", recent(ds, 1)),
            ("近 3 日", recent(ds, 3)),
            ("近 7 日", recent(ds, 7)),
            ("近 30 日", recent(ds, 30)),
            ("本月", month_now(ds)),
        ]
        .iter()
        .enumerate()
        .map(|(i, (k, m))| metric_row(k, m, i == selected))
        .collect();
        render_table(f, areas[2], "总览 · 时间窗口", rows, selected)
    } else {
        let mut rows = Vec::new();
        if let Some(months) = data.get("monthly").and_then(Value::as_object) {
            let mut keys: Vec<_> = months.keys().collect();
            keys.sort_by(|a, b| b.cmp(a));
            for (i, key) in keys.iter().enumerate() {
                let m = &months[*key];
                let metric = Metric {
                    input: n(m, "input"),
                    output: n(m, "output"),
                    cache: n(m, "cache_read") + n(m, "cache_write"),
                    cost: money(m),
                };
                rows.push(metric_row(key, &metric, i == selected))
            }
        }
        render_table(f, areas[2], "自然月汇总 · 每月一行", rows, selected)
    }
    f.render_widget(
        Paragraph::new("仅保留 HTML 对应的总览与自然月汇总 · Tab/←→ 切换 · ↑↓ 浏览 · Q/Esc 退出"),
        areas[3],
    );
}
