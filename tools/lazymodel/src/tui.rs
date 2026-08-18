use crate::engine::{self, Catalog, Route, RoutesFile};
use crossterm::{
    event::{self, Event, KeyCode, KeyEvent, KeyEventKind},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
    backend::CrosstermBackend,
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table, TableState},
    Frame, Terminal,
};
use std::{collections::HashSet, io};

pub fn run(data: &mut RoutesFile, catalog: &Catalog) -> io::Result<()> {
    enable_raw_mode()?;
    let mut out = io::stdout();
    execute!(out, EnterAlternateScreen)?;
    let mut term = Terminal::new(CrosstermBackend::new(out))?;
    let result = loop_ui(&mut term, data, catalog);
    disable_raw_mode()?;
    execute!(term.backend_mut(), LeaveAlternateScreen)?;
    term.show_cursor()?;
    result
}

fn target(scope: &str, harness: &str) -> String {
    match (harness, scope) {
        ("opencode", "global") => "%USERPROFILE%\\.config\\opencode\\opencode.jsonc".into(),
        ("opencode", "CCDS") => "%USERPROFILE%\\CCDS\\opencode.jsonc".into(),
        ("opencode", "CCglm") => "%USERPROFILE%\\CCglm\\opencode.jsonc".into(),
        ("opencode", "CCKimi") => "%USERPROFILE%\\CCKimi\\opencode.jsonc".into(),
        ("opencode", "CCGame") => "%USERPROFILE%\\CCGame\\.opencode\\opencode.jsonc".into(),
        ("oh-my-openagent", "global") => {
            "%USERPROFILE%\\.config\\opencode\\oh-my-openagent.json".into()
        }
        ("oh-my-openagent", "CCGame") => {
            "%USERPROFILE%\\CCGame\\.opencode\\oh-my-openagent.json".into()
        }
        ("hermes", _) => "%LOCALAPPDATA%\\hermes\\config.yaml".into(),
        ("claude-code", _) => "%USERPROFILE%\\.claude\\agents\\*.md".into(),
        ("kilo", _) => {
            "%USERPROFILE%\\.config\\kilo\\kilo.jsonc（用量：.local\\share\\kilo\\kilo.db）".into()
        }
        ("dsh", _) => "%USERPROFILE%\\.dsh\\settings.yaml（CCGame DSH profile）".into(),
        _ => "(由 scope 决定)".into(),
    }
}
fn zh_role(r: &str) -> &str {
    match r {
        "primary" => "主模型",
        "small_model" => "廉价快速模型",
        "visual-engineering" => "视觉工程",
        "quick/explore/librarian" => "快速探索/资料",
        "general agents/categories" => "一般 Agent/类别",
        "vision" => "视觉理解",
        "web/compression/title/curator/session_search" => "辅助任务",
        "architect" => "架构设计",
        "code-reviewer" => "代码审查",
        "code-implementer" => "代码实现",
        "debugger" => "调试修复",
        "code-explorer" => "代码探索",
        "translator" => "翻译格式化",
        x => x,
    }
}
fn cost_label(tier: &str) -> String {
    match tier {
        "便宜" => "便宜".into(),
        "高" | "中高" => "贵".into(),
        "中" | "中低" => "中".into(),
        "图像" => "图像".into(),
        _ => "-".into(),
    }
}

fn strength(model: &str) -> String {
    if model.starts_with("kilo-auto/free") {
        return "Kilo 自动免费路由".into();
    }
    if model.starts_with("kilo-auto/small") {
        return "Kilo 快速小模型".into();
    }
    if model.starts_with("deepseek-v4-flash") {
        return "复杂分析/调试/长链路推理".into();
    }
    if model == "gpt-5.6-luna" {
        return "一般编码/日常修改/常规问答".into();
    }
    if model == "gpt-5.6-terra" {
        return "难点编码/架构分析/复杂调试".into();
    }
    if model == "gpt-5.6-sol" {
        return "核心推理/方案设计/关键决策".into();
    }
    if model.starts_with("gpt-5.3-codex") {
        return "代码实现/工具调用/工程任务".into();
    }
    if model.starts_with("gpt-") {
        return "通用编码/分析/工具调用".into();
    }
    if model.starts_with("gemini") && (model.contains("pro") || model.contains("image")) {
        return "视觉理解/界面分析/图片任务".into();
    }
    if model.starts_with("gemini") {
        return "快速视觉/轻量分析/资料处理".into();
    }
    if model.starts_with("qwen") {
        return "通用编码/中文任务/方案整理".into();
    }
    if model.starts_with("kimi") {
        return "代码阅读/中文分析/长上下文".into();
    }
    if model.starts_with("glm") {
        return "中文问答/常规编码/工具任务".into();
    }
    if model.starts_with("claude") {
        return "复杂编码/审查/架构分析".into();
    }
    "通用任务".into()
}

fn info(c: &Catalog, m: &str) -> (String, String, String) {
    let i = engine::model_info(c, m);
    if i.tier != "-" {
        return (i.tier.clone(), i.csharp, cost_label(&i.tier));
    }
    let tier = if m.starts_with("deepseek") || m.contains("flash-lite") || m.contains("mini") {
        "便宜"
    } else if m.starts_with("kilo-auto") {
        "便宜"
    } else if m.starts_with("gemini") {
        "中"
    } else {
        "-"
    };
    (tier.to_string(), "-".into(), cost_label(tier))
}

struct ModelPicker {
    index: usize,
    models: Vec<String>,
}
struct EffortPicker {
    index: usize,
}

fn loop_ui(
    term: &mut Terminal<CrosstermBackend<io::Stdout>>,
    data: &mut RoutesFile,
    catalog: &Catalog,
) -> io::Result<()> {
    let mut selected = 0usize;
    let mut marked: HashSet<usize> = (0..data.routes.len()).collect();
    let mut model_picker: Option<ModelPicker> = None;
    let mut effort_picker: Option<EffortPicker> = None;
    loop {
        term.draw(|f| {
            let size = f.size();
            let chunks = Layout::default()
                .direction(Direction::Vertical)
                .constraints([Constraint::Min(4), Constraint::Length(3)])
                .split(size);
            let rows: Vec<Row> = data
                .routes
                .iter()
                .enumerate()
                .map(|(i, r)| {
                    let (tier, csharp, price) = info(catalog, &r.model);
                    Row::new(vec![
                        Cell::from(if marked.contains(&i) { "[x]" } else { "[ ]" }),
                        Cell::from(r.harness.clone()),
                        Cell::from(r.scope.clone()),
                        Cell::from(zh_role(&r.role).to_string()),
                        Cell::from(r.model.clone()),
                        Cell::from(r.effort.clone()),
                        Cell::from(format!("{} / {}", tier, csharp)),
                        Cell::from(price),
                        Cell::from(strength(&r.model)),
                        Cell::from(target(&r.scope, &r.harness)),
                    ])
                })
                .collect();
            let mut state = TableState::default();
            state.select(Some(selected.min(rows.len().saturating_sub(1))));
            let table = Table::new(
                rows,
                [
                    Constraint::Length(5),
                    Constraint::Length(16),
                    Constraint::Length(9),
                    Constraint::Length(20),
                    Constraint::Length(23),
                    Constraint::Length(8),
                    Constraint::Length(12),
                    Constraint::Length(14),
                    Constraint::Min(28),
                    Constraint::Min(35),
                ],
            )
            .header(
                Row::new([
                    "写入",
                    "Harness",
                    "Scope",
                    "角色",
                    "模型",
                    "Effort",
                    "性能",
                    "价格",
                    "定位/擅长",
                    "目标文件",
                ])
                .style(
                    Style::default()
                        .fg(Color::Cyan)
                        .add_modifier(Modifier::BOLD),
                ),
            )
            .block(Block::default().borders(Borders::ALL).title(format!(
                " lazymodel · LiteLLM {} models ({}) · 已选 {}/{} ",
                catalog.models.len(),
                catalog.date,
                marked.len(),
                data.routes.len()
            )))
            .highlight_style(
                Style::default()
                    .bg(Color::Cyan)
                    .fg(Color::Black)
                    .add_modifier(Modifier::BOLD),
            )
            .highlight_symbol("▶ ");
            f.render_stateful_widget(table, chunks[0], &mut state);
            f.render_widget(
                Paragraph::new(
                    "↑↓选择 · Space勾选 · Enter选择模型 · A应用已勾选 · U全选 · N全不选 · Q退出",
                )
                .block(Block::default().borders(Borders::ALL)),
                chunks[1],
            );
            if let Some(p) = &model_picker {
                render_model_picker(f, size, p, catalog)
            } else if let Some(p) = &effort_picker {
                render_effort_picker(f, size, p)
            }
        })?;
        if let Event::Key(KeyEvent { code, kind, .. }) = event::read()? {
            if kind != KeyEventKind::Press {
                continue;
            }
            if let Some(p) = &mut model_picker {
                match code {
                    KeyCode::Esc => model_picker = None,
                    KeyCode::Up => p.index = p.index.saturating_sub(1),
                    KeyCode::Down => p.index = (p.index + 1).min(p.models.len().saturating_sub(1)),
                    KeyCode::Enter => {
                        data.routes[selected].model = p.models[p.index].clone();
                        model_picker = None;
                        effort_picker = Some(EffortPicker { index: 0 });
                    }
                    _ => {}
                }
                continue;
            }
            if let Some(p) = &mut effort_picker {
                let opts = ["default", "low", "medium", "high", "xhigh"];
                match code {
                    KeyCode::Esc => effort_picker = None,
                    KeyCode::Up => p.index = p.index.saturating_sub(1),
                    KeyCode::Down => p.index = (p.index + 1).min(opts.len() - 1),
                    KeyCode::Enter => {
                        data.routes[selected].effort = opts[p.index].into();
                        engine::save_routes(data)?;
                        effort_picker = None;
                    }
                    _ => {}
                }
                continue;
            }
            match code {
                KeyCode::Char('q') | KeyCode::Esc => break Ok(()),
                KeyCode::Up => selected = selected.saturating_sub(1),
                KeyCode::Down => selected = (selected + 1).min(data.routes.len().saturating_sub(1)),
                KeyCode::Char(' ') => {
                    if !marked.remove(&selected) {
                        marked.insert(selected);
                    }
                }
                KeyCode::Char('u') | KeyCode::Char('U') => {
                    marked = (0..data.routes.len()).collect()
                }
                KeyCode::Char('n') | KeyCode::Char('N') => marked.clear(),
                KeyCode::Char('a') | KeyCode::Char('A') => {
                    let routes: Vec<Route> = data
                        .routes
                        .iter()
                        .enumerate()
                        .filter(|(i, _)| marked.contains(i))
                        .map(|(_, r)| r.clone())
                        .collect();
                    engine::apply_routes(&RoutesFile {
                        version: data.version,
                        routes,
                    })?;
                }
                KeyCode::Enter => {
                    model_picker = Some(ModelPicker {
                        index: catalog
                            .models
                            .iter()
                            .position(|m| m == &data.routes[selected].model)
                            .unwrap_or(0),
                        models: {
                            let mut models = catalog.models.clone();
                            for route in &data.routes {
                                if (route.harness == "kilo" || route.harness == "dsh")
                                    && !models.contains(&route.model)
                                {
                                    models.push(route.model.clone());
                                }
                            }
                            models
                        },
                    })
                }
                _ => {}
            }
        }
    }
}
fn render_model_picker(f: &mut Frame, size: Rect, p: &ModelPicker, c: &Catalog) {
    let h = 20.min(size.height.saturating_sub(4));
    let area = Rect {
        x: size.width.saturating_sub(100) / 2,
        y: size.height.saturating_sub(h) / 2,
        width: 100.min(size.width.saturating_sub(4)),
        height: h,
    };
    f.render_widget(Clear, area);
    let rows: Vec<Row> = p
        .models
        .iter()
        .map(|m| {
            let (tier, score, price) = info(c, m);
            Row::new(vec![
                Cell::from(m.clone()),
                Cell::from(tier),
                Cell::from(score),
                Cell::from(price),
                Cell::from(strength(m)),
            ])
        })
        .collect();
    let mut s = TableState::default();
    s.select(Some(p.index));
    let t = Table::new(
        rows,
        [
            Constraint::Length(30),
            Constraint::Length(10),
            Constraint::Length(10),
            Constraint::Length(20),
        ],
    )
    .header(
        Row::new(["模型", "档位", "C#性能", "价格", "定位/擅长"]).style(
            Style::default()
                .fg(Color::Yellow)
                .add_modifier(Modifier::BOLD),
        ),
    )
    .block(
        Block::default()
            .borders(Borders::ALL)
            .title(" 选择模型 · ↑↓移动 Enter确认 Esc取消 "),
    )
    .highlight_style(Style::default().bg(Color::Yellow).fg(Color::Black));
    f.render_stateful_widget(t, area, &mut s)
}
fn render_effort_picker(f: &mut Frame, size: Rect, p: &EffortPicker) {
    let opts = ["default", "low", "medium", "high", "xhigh"];
    let area = Rect {
        x: size.width / 2 - 25,
        y: size.height / 2 - 5,
        width: 50,
        height: 10,
    };
    f.render_widget(Clear, area);
    let rows: Vec<Row> = opts.iter().map(|x| Row::new([*x])).collect();
    let mut s = TableState::default();
    s.select(Some(p.index));
    let t = Table::new(
        rows,
        [
            Constraint::Length(26),
            Constraint::Length(10),
            Constraint::Length(10),
            Constraint::Length(10),
            Constraint::Min(30),
        ],
    )
    .block(
        Block::default()
            .borders(Borders::ALL)
            .title(" 选择 Effort "),
    )
    .highlight_style(Style::default().bg(Color::Yellow).fg(Color::Black));
    f.render_stateful_widget(t, area, &mut s)
}
