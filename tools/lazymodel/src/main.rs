mod engine;
mod tui;

fn main() -> std::io::Result<()> {
    if std::env::args().any(|arg| arg == "--check") {
        let catalog = engine::refresh_catalog();
        let routes = engine::load_routes()?;
        println!(
            "LiteLLM catalog: {} models ({})",
            catalog.models.len(),
            catalog.date
        );
        println!("Routes: {}", routes.routes.len());
        for route in routes.routes {
            println!(
                "{} / {} / {} -> {} [{}]",
                route.harness, route.scope, route.role, route.model, route.effort
            );
        }
        return Ok(());
    }
    if std::env::args().any(|arg| arg == "--help" || arg == "-h") {
        println!("lazymodel - Rust TUI for cross-harness model routing");
        println!("  lazymodel.exe        Open the full-screen route editor");
        println!("  ↑↓ select · Enter edit · a apply · q exit");
        println!("  --check              Refresh catalog and print routes without opening TUI");
        println!("  LiteLLM model catalog refreshes on the first launch each day");
        return Ok(());
    }
    let catalog = engine::refresh_catalog();
    let mut routes = engine::load_routes()?;
    tui::run(&mut routes, &catalog)
}
