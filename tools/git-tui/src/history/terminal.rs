use std::io::{self, stdout, Stdout, Write};

use crossterm::{
    event::{DisableBracketedPaste, DisableMouseCapture, EnableBracketedPaste, EnableMouseCapture},
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{backend::CrosstermBackend, Terminal};

pub(super) type Tui = Terminal<CrosstermBackend<Stdout>>;

pub(super) struct TerminalGuard {
    pub(super) active: bool,
}

impl TerminalGuard {
    pub(super) fn enter() -> io::Result<Self> {
        enable_raw_mode()?;
        execute!(
            stdout(),
            EnterAlternateScreen,
            EnableMouseCapture,
            EnableBracketedPaste
        )?;
        Ok(Self { active: true })
    }

    pub(super) fn suspend(&mut self, terminal: &mut Tui) -> io::Result<()> {
        terminal.show_cursor()?;
        terminal.backend_mut().flush()?;
        self.restore(terminal)?;
        self.active = false;
        Ok(())
    }

    pub(super) fn resume(&mut self, terminal: &mut Tui) -> io::Result<()> {
        enable_raw_mode()?;
        execute!(
            terminal.backend_mut(),
            EnterAlternateScreen,
            EnableMouseCapture,
            EnableBracketedPaste
        )?;
        terminal.clear()?;
        self.active = true;
        Ok(())
    }

    pub(super) fn restore(&self, terminal: &mut Tui) -> io::Result<()> {
        disable_raw_mode()?;
        execute!(
            terminal.backend_mut(),
            DisableBracketedPaste,
            DisableMouseCapture,
            LeaveAlternateScreen
        )?;
        terminal.show_cursor()
    }
}

impl Drop for TerminalGuard {
    fn drop(&mut self) {
        if self.active {
            restore_without_terminal();
        }
    }
}

pub(super) fn install_panic_hook() -> Box<dyn Fn(&std::panic::PanicHookInfo<'_>) + Sync + Send> {
    let previous = std::panic::take_hook();
    std::panic::set_hook(Box::new(|info| {
        restore_without_terminal();
        eprintln!("{info}");
    }));
    previous
}

fn restore_without_terminal() {
    let _ = disable_raw_mode();
    let _ = execute!(
        stdout(),
        DisableBracketedPaste,
        DisableMouseCapture,
        LeaveAlternateScreen,
        crossterm::cursor::Show
    );
}
