//! 文件提交历史的全屏 TUI。

mod git;
mod terminal;
mod view;

use std::{
    io::{self, stdout},
    path::{Path, PathBuf},
};

use crossterm::event::{
    self, Event, KeyCode, KeyEvent, KeyEventKind, KeyModifiers, MouseButton, MouseEventKind,
};
use ratatui::{backend::CrosstermBackend, layout::Rect, widgets::TableState, Terminal};

use git::{
    filter_by_message, load_commits, pipe_diff_to_delta, pipe_diff_two_to_delta, resolve_file,
    Commit,
};
use terminal::{TerminalGuard, Tui};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Stage {
    CommitList,
    AuthorFilter,
    Search,
    Quit,
}

#[derive(Debug, PartialEq, Eq)]
enum CompareAction {
    Selected,
    Cancelled,
    Compare { first: String, second: String },
}

enum DiffAction {
    None,
    Workspace(String),
    Commits { first: String, second: String },
}

struct App {
    repo_root: PathBuf,
    file: PathBuf,
    display_file: String,
    commits: Vec<Commit>,
    visible: Vec<usize>,
    author_filter: String,
    search: String,
    stage: Stage,
    table_state: TableState,
    table_area: Rect,
    error: Option<String>,
    first_compare_sha: Option<String>,
    loading_more: bool,
}

pub fn run(file: &Path) -> io::Result<()> {
    let (repo_root, relative_file) = resolve_file(file)?;
    let commits = load_commits(&repo_root, &relative_file, "")?;
    if commits.is_empty() {
        return Err(io::Error::new(
            io::ErrorKind::NotFound,
            "没有找到该文件的提交记录",
        ));
    }
    let mut app = App {
        display_file: relative_file.display().to_string(),
        repo_root,
        file: relative_file,
        visible: (0..commits.len()).collect(),
        commits,
        author_filter: String::new(),
        search: String::new(),
        stage: Stage::CommitList,
        table_state: TableState::default().with_selected(Some(0)),
        table_area: Rect::default(),
        error: None,
        first_compare_sha: None,
        loading_more: false,
    };

    let previous_hook = terminal::install_panic_hook();
    let mut guard = TerminalGuard::enter()?;
    let backend = CrosstermBackend::new(stdout());
    let mut terminal = Terminal::new(backend)?;
    let result = run_loop(&mut terminal, &mut guard, &mut app);
    let restore_result = guard.restore(&mut terminal);
    guard.active = false;
    std::panic::set_hook(previous_hook);
    result.and(restore_result)
}

fn run_loop(terminal: &mut Tui, guard: &mut TerminalGuard, app: &mut App) -> io::Result<()> {
    while app.stage != Stage::Quit {
        view::draw(terminal, app)?;
        match event::read()? {
            Event::Paste(text) => handle_paste(app, &text)?,
            Event::Key(key) if key.kind == KeyEventKind::Press => {
                if is_ctrl_c(key) {
                    app.stage = Stage::Quit;
                } else {
                    let action = handle_key(app, key)?;
                    show_diff(terminal, guard, app, action)?;
                }
            }
            Event::Mouse(mouse) if mouse.kind == MouseEventKind::Down(MouseButton::Left) => {
                select_clicked_row(app, mouse.row);
            }
            Event::FocusGained
            | Event::FocusLost
            | Event::Key(_)
            | Event::Mouse(_)
            | Event::Resize(_, _) => {}
        }
    }
    Ok(())
}

fn handle_key(app: &mut App, key: KeyEvent) -> io::Result<DiffAction> {
    app.error = None;
    match app.stage {
        Stage::CommitList => match key.code {
            KeyCode::Up => move_selection(app, false),
            KeyCode::Down => {
                move_selection(app, true);
                if app
                    .table_state
                    .selected()
                    .is_some_and(|i| i + 1 >= app.visible.len())
                {
                    load_more_if_needed(app)?;
                }
            }
            KeyCode::Enter => {
                if let Some(commit) = selected_commit(app) {
                    return Ok(DiffAction::Workspace(commit.sha.clone()));
                }
            }
            KeyCode::Char('c') => {
                if let Some(commit) = selected_commit(app) {
                    let sha = commit.sha.clone();
                    match update_compare_selection(&mut app.first_compare_sha, &sha) {
                        CompareAction::Selected => {}
                        CompareAction::Cancelled => {
                            app.error = Some("已取消提交比对".to_string());
                        }
                        CompareAction::Compare { first, second } => {
                            return Ok(DiffAction::Commits { first, second });
                        }
                    }
                }
            }
            KeyCode::Char('a') => app.stage = Stage::AuthorFilter,
            KeyCode::Char('/') => app.stage = Stage::Search,
            KeyCode::Char('q') | KeyCode::Esc => app.stage = Stage::Quit,
            _ => {}
        },
        Stage::AuthorFilter => match key.code {
            KeyCode::Esc | KeyCode::Enter => app.stage = Stage::CommitList,
            KeyCode::Backspace => {
                app.author_filter.pop();
                reload_author(app)?;
            }
            KeyCode::Char(character) if !key.modifiers.contains(KeyModifiers::CONTROL) => {
                app.author_filter.push(character);
                reload_author(app)?;
            }
            _ => {}
        },
        Stage::Search => match key.code {
            KeyCode::Esc | KeyCode::Enter => app.stage = Stage::CommitList,
            KeyCode::Backspace => {
                app.search.pop();
                apply_search(app);
            }
            KeyCode::Char(character) if !key.modifiers.contains(KeyModifiers::CONTROL) => {
                app.search.push(character);
                apply_search(app);
            }
            _ => {}
        },
        Stage::Quit => {}
    }
    Ok(DiffAction::None)
}

fn handle_paste(app: &mut App, text: &str) -> io::Result<()> {
    let clean = text.replace(['\r', '\n'], "");
    match app.stage {
        Stage::AuthorFilter => {
            app.author_filter.push_str(&clean);
            reload_author(app)
        }
        Stage::Search => {
            app.search.push_str(&clean);
            apply_search(app);
            Ok(())
        }
        Stage::CommitList | Stage::Quit => Ok(()),
    }
}

fn load_more_if_needed(app: &mut App) -> io::Result<()> {
    if app.loading_more || app.commits.len() < 30 {
        return Ok(());
    }
    app.loading_more = true;
    let result = git::load_more_commits(
        &app.repo_root,
        &app.file,
        &app.author_filter,
        app.commits.len(),
    );
    app.loading_more = false;
    match result {
        Ok(more) => {
            let start = app.commits.len();
            app.commits.extend(more);
            app.visible.extend(start..app.commits.len());
        }
        Err(error) => app.error = Some(error.to_string()),
    }
    Ok(())
}

fn reload_author(app: &mut App) -> io::Result<()> {
    match load_commits(&app.repo_root, &app.file, &app.author_filter) {
        Ok(commits) => {
            app.commits = commits;
            apply_search(app);
            app.error = None;
        }
        Err(error) => app.error = Some(error.to_string()),
    }
    Ok(())
}

fn apply_search(app: &mut App) {
    app.visible = filter_by_message(&app.commits, &app.search);
    app.table_state
        .select((!app.visible.is_empty()).then_some(0));
}

fn move_selection(app: &mut App, forward: bool) {
    let len = app.visible.len();
    if len == 0 {
        app.table_state.select(None);
        return;
    }
    let current = app.table_state.selected().unwrap_or(0);
    let next = if forward {
        (current + 1).min(len - 1)
    } else {
        current.checked_sub(1).unwrap_or(len - 1)
    };
    app.table_state.select(Some(next));
}

fn select_clicked_row(app: &mut App, row: u16) {
    let first_data_row = app.table_area.y.saturating_add(3);
    if row < first_data_row || row >= app.table_area.bottom().saturating_sub(1) {
        return;
    }
    let clicked = usize::from(row - first_data_row) + app.table_state.offset();
    if clicked < app.visible.len() {
        app.table_state.select(Some(clicked));
    }
}

fn selected_commit(app: &App) -> Option<&Commit> {
    app.table_state
        .selected()
        .and_then(|position| app.visible.get(position))
        .and_then(|index| app.commits.get(*index))
}

fn update_compare_selection(first_compare_sha: &mut Option<String>, sha: &str) -> CompareAction {
    let Some(first) = first_compare_sha.take() else {
        *first_compare_sha = Some(sha.to_string());
        return CompareAction::Selected;
    };
    if first == sha {
        return CompareAction::Cancelled;
    }
    CompareAction::Compare {
        first,
        second: sha.to_string(),
    }
}

fn show_diff(
    terminal: &mut Tui,
    guard: &mut TerminalGuard,
    app: &mut App,
    action: DiffAction,
) -> io::Result<()> {
    let result = match action {
        DiffAction::None => return Ok(()),
        DiffAction::Workspace(sha) => {
            guard.suspend(terminal)?;
            print_loading("正在比对与工作区的差异");
            let r = pipe_diff_to_delta(&app.repo_root, &app.file, &sha);
            clear_loading();
            r
        }
        DiffAction::Commits { first, second } => {
            guard.suspend(terminal)?;
            print_loading("正在比对两个提交的差异");
            let r = pipe_diff_two_to_delta(&app.repo_root, &app.file, &first, &second);
            clear_loading();
            r
        }
    };
    guard.resume(terminal)?;
    if let Err(error) = result {
        app.error = Some(error.to_string());
    }
    Ok(())
}

/// 打印加载提示（suspend 后、diff 生成前，给用户即时反馈）
fn print_loading(label: &str) {
    use std::io::Write;
    print!("\r⏳ {label}...");
    let _ = std::io::stdout().flush();
}

/// 清除加载提示
fn clear_loading() {
    use std::io::Write;
    print!("\r{}\r", " ".repeat(60));
    let _ = std::io::stdout().flush();
}

fn is_ctrl_c(key: KeyEvent) -> bool {
    key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('c')
}

#[cfg(test)]
mod tests {
    use super::{update_compare_selection, CompareAction};

    #[test]
    fn selects_first_commit_when_compare_is_empty() {
        let mut first = None;

        let action = update_compare_selection(&mut first, "abc1234");

        assert_eq!(action, CompareAction::Selected);
        assert_eq!(first.as_deref(), Some("abc1234"));
    }

    #[test]
    fn cancels_compare_when_second_commit_matches_first() {
        let mut first = Some("abc1234".to_string());

        let action = update_compare_selection(&mut first, "abc1234");

        assert_eq!(action, CompareAction::Cancelled);
        assert_eq!(first, None);
    }

    #[test]
    fn compares_and_clears_selection_when_second_commit_differs() {
        let mut first = Some("abc1234".to_string());

        let action = update_compare_selection(&mut first, "def5678");

        assert_eq!(
            action,
            CompareAction::Compare {
                first: "abc1234".to_string(),
                second: "def5678".to_string(),
            }
        );
        assert_eq!(first, None);
    }
}
