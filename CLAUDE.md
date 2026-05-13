# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build

```bash
# Debug build
dotnet build NumDesTools.sln -c Debug

# Release build (also packs .xll via Excel-DNA post-build → packFromBin/)
dotnet build NumDesTools.sln -c Release

# Run tests
dotnet test NumDesTools.sln
```

Post-build calls `packFromBin\ReNamePack.bat` which renames packed XLL files to `NumDesToolsPack64.xll` / `NumDesToolsPack.xll`.

Code must pass **0 errors** before reporting done. Follow **ReSharper** naming rules and format with **CSharpier**.

## Architecture

Three projects in one solution:

| Project | Type | Purpose |
|---------|------|---------|
| **NumDesTools** | Excel XLL add-in (net9.0-windows) | Main plugin: ribbon, UDFs, UI windows |
| **NumDesTools.Scanner** | Console app (net9.0-windows) | CLI validator for activity config tables + Feishu integration |
| **NumDesTools.Tests** | xUnit test (net9.0-windows) | Unit tests for LTE calc, scanner, map operations |

### NumDesTools — Key Modules

**Entry point:** `NumDesAddIn.cs` — implements `IExcelAddIn` + `ExcelRibbon`; holds the global `App` (Excel.Application) reference; routes all ribbon button clicks through a debounced (500 ms) dictionary dispatch.

**Ribbon definition:** `RibbonUI.xml` — XML layout. Label/state driven by `GlobalVariable.cs` toggle strings.

**UDFs:** `ScUDFs.cs` (2100+ lines) — all `[ExcelFunction]`-decorated functions exposed to Excel: FindKey*, Trans2Array*, JSON export helpers, game-specific calculations (LTE chains, Alice, Dota sim).

**Config:** `Config/GlobalVariable.cs` — loads/saves `Documents\NumDesGlobalKey.json`; holds paths, API keys, UI toggle labels, spotlight mode, git root.

**ConflictResolver namespace** — Git-based Excel conflict resolution UI:
- `ExcelConflictDiffer.cs` — diffs two Excel files into `FileDiff` / `SheetDiff` / `RowConflict` / `CellConflict` models
- `ConflictModels.cs` — data models with `IsResolved`, drag-selection `IsSelected`, `INotifyPropertyChanged`
- `ConflictApplier.cs` — writes chosen values back, calls `git add`, appends to `.git/MERGE_MSG`
- `UI/ExcelConflictWindow.xaml(.cs)` — WPF window: drag multi-select, filter-select, selection action bar, unresolved-check before apply

**Cell highlighting:** `CellHighlighter.cs` — `ViewportHelper.GetViewportRange` (unions all `win.Panes[i].VisibleRange`, intersects with UsedRange) shared by `CellHighlighter` (Find/FindNext same-value) and `CellSpotlightHighlighter` (fill-mode row/col color). `CrosslightOverlay.cs` handles the overlay (transparent cross-lines) mode.

**Advance namespace:** `ExcelDataToDB.cs`, `IdPrefixIndex.cs` — game data extraction and cross-table indexing.

**ExcelToLua namespace:** `ExcelReader.cs` — converts Excel table schemas (row 2 = col names, row 3 = types, row 4 = labels) to Lua with type definitions.

### Data Conventions (Excel tables)

- Sheet prefix `c_*` = client-side, `s_*` = server-side, `#*` = hidden/meta
- Row 1 = header, Row 2 = column names, Row 3 = types (`int`/`string`/…), Row 4 = Chinese labels
- `#` prefix columns = comment/display columns; shown in conflict resolver row headers

### Threading

Long operations use `ExcelAsyncUtil.QueueAsMacro()`. Never block the Excel UI thread directly.

### Key dependencies

`ExcelDna.AddIn 1.9.0`, `EPPlus 8.2.0`, `MiniExcel 1.42.0`, `LibGit2Sharp 0.31.0`, `NLua`, `MathNet.Numerics`.

### Excel 读写规范

- **读取 xlsx** → EPPlus（`OfficeOpenXml`）。参考 `NumDesTools.Scanner/ExcelReader.cs`。
- **写入 xlsx** → EPPlus（`OfficeOpenXml`）。参考 `NumDesTools.Scanner/ActivityWriter.cs`（追加行）和 `NumDesTools.Scanner/LteMapWriter.cs`（新建带样式/列宽/行高的 sheet）。
- 每个入口必须先调用 `ExcelPackage.License.SetNonCommercialPersonal("NumDesTools")`。
- MiniExcel 已在依赖中但**不用于本项目**，不要引入。
