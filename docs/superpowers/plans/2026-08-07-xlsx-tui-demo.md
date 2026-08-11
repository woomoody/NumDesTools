# xlsx-tui 最小演示实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在不改动现有 WPF XlsxEditor 的前提下，先验证 xlsx 保存保真度，再交付一个可打开、多 Sheet 浏览、键鼠定位、编辑单元格并覆盖保存的独立 Rust TUI 演示。

**Architecture:** 新建 `tools/xlsx-tui` Rust 二进制。读取层与 TUI 数据模型解耦；保存层先以 `umya-spreadsheet` 做真实工作簿 POC，只有确认未编辑内容和公式不会被静默破坏时才接入，否则使用 ZIP 原始复制加 worksheet XML 定点补丁。TUI 使用轻量 Normal/Insert 状态机，首版不引入公式计算、撤销或完整样式模型。

**Tech Stack:** Rust 2021；`ratatui 0.26`；`crossterm 0.27`；候选读取 `calamine`；候选保存 `umya-spreadsheet`；候选输入 `tui-textarea`；候选剪贴板 `arboard`；`unicode-width`；现有 `lazykey` 与 `conflict-tui` 生命周期模式。

## Global Constraints

- 不修改 `NumDesTools.XlsxEditor`、现有 WPF 代码或现有 Rust TUI 工具。
- 首版 Ctrl+S 覆盖原文件，但必须先写临时文件并成功后原子替换。
- 支持多个 worksheet；不支持多工作簿并行打开。
- 公式可以保留，也可以变为值；不强制全工作簿公式转值，但不得静默丢失公式或公式缓存结果。
- 不实现排序、筛选、冻结窗格、整行整列选择、撤销/重做、公式计算、样式/合并/图表编辑和 Ribbon 集成。
- 只使用 ASCII 源码标识符和注释；用户可见中文可以保留。
- 使用 `cargo fmt`、`cargo clippy`、`cargo test` 验证 Rust 改动。
- 保存失败时保留原文件和 dirty 状态；TUI 退出前恢复终端 raw mode/alternate screen。

## 文件结构

### Task 1 保存层 POC

**Files:**
- Create: `tools/xlsx-tui-poc/Cargo.toml`
- Create: `tools/xlsx-tui-poc/src/main.rs`
- Create: `tools/xlsx-tui-poc/tests/roundtrip.rs`
- Modify: `.gitignore` only if the POC creates a local target/output path not already ignored

**Interfaces:**
- Produces a command that copies a supplied workbook to a temporary output, changes one ordinary cell, saves, and emits a machine-readable summary of workbook/sheet/ZIP checks.
- The POC must expose a small internal function with a stable signature: `fn run_roundtrip(input: &Path, output: &Path) -> Result<RoundtripReport, Box<dyn Error>>`.

- [ ] **Step 1: Add the POC manifest with pinned candidate dependencies**

Use `umya-spreadsheet` for the first candidate and `zip`/`quick-xml` only for inspection. Do not add TUI dependencies to the POC.

```toml
[package]
name = "xlsx-tui-poc"
version = "0.1.0"
edition = "2021"

[dependencies]
umya-spreadsheet = "3.0.1"
zip = "2"
quick-xml = "0.31"
```

- [ ] **Step 2: Write the failing round-trip checks**

The test must locate a fixture through `XLSX_TUI_FIXTURE`; if it is absent, fail with an explicit message instead of silently skipping. Assert that the output remains a valid ZIP, has the same sheet count and names, and that the target changed.

```rust
#[test]
fn roundtrip_preserves_workbook_shape_and_changes_one_cell() {
    let input = std::env::var_os("XLSX_TUI_FIXTURE")
        .map(std::path::PathBuf::from)
        .expect("set XLSX_TUI_FIXTURE to a real workbook fixture");
    let output = tempfile::NamedTempFile::new().unwrap();
    let report = run_roundtrip(&input, output.path()).unwrap();
    assert!(report.valid_zip);
    assert_eq!(report.input_sheets, report.output_sheets);
    assert!(report.target_changed);
}
```

Add `tempfile = "3"` as a dev-dependency, or use a unique file under `std::env::temp_dir()` if the POC is kept dependency-minimal.

- [ ] **Step 3: Implement only the umya read/change/write path**

Open the input with `umya_spreadsheet::reader::xlsx::read`, select the first worksheet, read `A1`, set `A1` to a deterministic sentinel such as `xlsx-tui-poc`, and write the output with `umya_spreadsheet::writer::xlsx::write`. Never touch the user’s original file in the POC.

- [ ] **Step 4: Inspect preservation properties**

Add ZIP/XML inspection helpers that compare sheet names, workbook relationships, worksheet count, merged-cell XML, column definitions, cell style `s` attributes, formula `<f>` elements, and entry names before/after. Record differences in `RoundtripReport`; do not claim byte-for-byte preservation for umya.

- [ ] **Step 5: Run the POC against a real fixture**

Run:

```powershell
$env:XLSX_TUI_FIXTURE = 'C:\M1Work\public\Excels\Tables\Item.xlsx'
cargo test --manifest-path tools/xlsx-tui-poc/Cargo.toml -- --nocapture
```

Expected: the output parses as ZIP/XML, sheet names/count remain equal, target cell changes, and no formula/style/merge loss is observed. If any preservation check fails, mark the umya path rejected and do not reuse it in Task 2.

- [ ] **Step 6: Commit the POC only after evidence is recorded**

Use a focused commit only if the repository workflow requests commits; otherwise leave the working tree for review. Suggested message: `poc: validate xlsx round-trip preservation`.

### Task 2 Select the final save adapter

**Files:**
- Create: `tools/xlsx-tui/src/io.rs`
- Create: `tools/xlsx-tui/src/model.rs`
- Create: `tools/xlsx-tui/tests/io_tests.rs`
- Modify: `tools/xlsx-tui/Cargo.toml`

**Interfaces:**
- `pub struct WorkbookModel { pub sheets: Vec<SheetModel> }`.
- `pub struct SheetModel { pub name: String, pub rows: usize, pub cols: usize, pub cells: Vec<Vec<String>> }`.
- `pub struct CellAddress { pub row: usize, pub col: usize }`.
- `pub fn load_workbook(path: &Path) -> Result<WorkbookModel, XlsxError>`.
- `pub fn save_workbook(path: &Path, model: &WorkbookModel, changes: &[CellChange]) -> Result<(), XlsxError>`.
- `pub struct CellChange { pub sheet_index: usize, pub address: CellAddress, pub value: String }`.

- [ ] **Step 1: Lock the save decision from Task 1 evidence**

If umya passes the real-fixture preservation checks, implement the adapter around umya. If it fails, implement a ZIP raw-copy adapter: read display data with calamine; copy every ZIP entry unchanged except targeted worksheet XML; write edited strings as `inlineStr`, preserve cell `s`, and leave formulas untouched unless the user edits that cell.

- [ ] **Step 2: Add model-level tests**

Test load dimensions, sheet names, blank cells, ordinary strings, numbers rendered as text, and cell changes. Test that no change list produces no output replacement.

- [ ] **Step 3: Add atomic replacement**

Write to `path.with_extension("xlsx.tui.tmp")` or a unique sibling temporary path. Flush and close the output before replacing the original. On Windows use a replace operation that does not delete the original before the new file is complete; retain the temp path in the returned error if replacement fails.

- [ ] **Step 4: Verify formulas are not silently lost**

Use a fixture containing at least one formula. After an ordinary-cell edit, inspect the output and assert either the original formula remains or the adapter explicitly reports a documented formula-to-value conversion. Do not accept an unexplained disappearance.

### Task 3 TUI state and input

**Files:**
- Create: `tools/xlsx-tui/src/app.rs`
- Create: `tools/xlsx-tui/src/editor.rs`
- Create: `tools/xlsx-tui/src/input.rs`
- Create: `tools/xlsx-tui/tests/app_tests.rs`

**Interfaces:**
- `enum Mode { Normal, Insert, Help, QuitConfirm }`.
- `struct App { workbook: WorkbookModel, sheet_index: usize, cursor: CellAddress, row_offset: usize, col_offset: usize, mode: Mode, edit_buffer: String, dirty: bool, status: String }`.
- `enum AppAction { Move { row_delta: isize, col_delta: isize }, Page { row_delta: isize }, SheetNext, SheetPrevious, BeginEdit, EditChar(char), Backspace, CommitEdit, CancelEdit, Save, Quit, ConfirmQuit(bool), ToggleHelp }`.
- `fn apply_action(app: &mut App, action: AppAction) -> Result<(), AppError>`.

- [ ] **Step 1: Write tests for movement and editing state**

Cover upper/lower bounds, sheet transitions, begin/commit/cancel editing, empty value clearing, and dirty state transitions.

- [ ] **Step 2: Implement the small state machine**

Normal mode handles movement, sheet switching, save, help, and edit entry. Insert mode appends Unicode characters, handles Backspace, commits on Enter, and restores the old value on Esc. Do not add Vim operator/motion machinery.

- [ ] **Step 3: Implement keyboard and mouse translation**

Translate crossterm key events and mouse events into `AppAction`. Use the rendered grid rectangle and fixed row/column metrics to map mouse coordinates. Ignore clicks outside the grid and clamp resulting coordinates.

### Task 4 TUI rendering and executable

**Files:**
- Create: `tools/xlsx-tui/src/main.rs`
- Create: `tools/xlsx-tui/src/tui.rs`
- Modify: `tools/xlsx-tui/Cargo.toml`

**Interfaces:**
- `fn draw(frame: &mut ratatui::Frame, app: &App)`.
- `fn run(path: &Path) -> Result<(), Box<dyn Error>>`.

- [ ] **Step 1: Add TUI dependencies only after the model compiles**

Use the existing versions first: `ratatui = "0.26"`, `crossterm = "0.27"`, `calamine` at the version required by the chosen IO adapter, and `unicode-width`. Add `tui-textarea` or `arboard` only if the current dependency graph resolves without a ratatui major-version split.

- [ ] **Step 2: Implement terminal guard**

Follow `tools/lazykey/src/tui.rs` and `tools/conflict-tui/src/tui.rs`: enable raw mode, enter alternate screen, enable mouse capture, run the loop, and restore all terminal state on normal exit and panic/error paths.

- [ ] **Step 3: Render the minimal layout**

Render title, sheet tabs, row-number column, visible data cells, current-cell highlight, edit buffer, dirty marker, and status/footer. Use Unicode display width for truncation. Render only visible rows/columns.

- [ ] **Step 4: Wire event loop and save**

Poll crossterm events, dispatch `AppAction`, redraw after actions, call `save_workbook` on Ctrl+S, and keep dirty state on save failure. Add a quit confirmation screen for dirty workbooks.

### Task 5 Build, regression verification, and manual trial

**Files:**
- Modify: `tools/xlsx-tui/README.md`
- Modify: `docs/superpowers/specs/2026-08-07-xlsx-tui-demo-design.md` only if behavior differs from tested implementation

- [ ] **Step 1: Run Rust formatting, tests, and clippy**

```powershell
cargo fmt --all -- --check
cargo test --manifest-path tools/xlsx-tui/Cargo.toml
cargo clippy --manifest-path tools/xlsx-tui/Cargo.toml --all-targets -- -D warnings
```

- [ ] **Step 2: Build the release demo**

```powershell
cargo build --manifest-path tools/xlsx-tui/Cargo.toml --release
```

- [ ] **Step 3: Run the real manual trial**

Use `C:\M1Work\public\Excels\Tables\Item.xlsx`, switch at least two sheets, click and move through cells, edit a normal value, save, reopen with Excel, and check that no repair prompt appears. Do not use the original as the first destructive trial; copy it to a disposable fixture before the first run even though the final demo supports in-place save.

- [ ] **Step 4: Document limitations and launch command**

README must show `cargo run --manifest-path tools/xlsx-tui/Cargo.toml -- path\to\file.xlsx`, the supported key/mouse actions, the in-place save warning, and the exact unsupported features.
