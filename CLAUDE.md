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

### 输出目录规范

所有值得保留的产出文件（xlsx 报告、md 分析文档、html、json 分析产物）统一写到 **`OutputRootPath`** 下，默认值为 `Documents\NumDesOutput\`（本地独立 git 仓库，不推送）。

**子目录规范：**

| 子目录 | 用途 |
|--------|------|
| `reports\` | xlsx/html 报告（竞品分析、地编信息、LTE 配置模版等） |
| `analysis\` | md 分析文档（竞品深度分析、设计规范、配置草稿等） |
| `misc\` | 插件偶发产出（溯源结果.xlsx、表格关系.json 等） |

**写文件规则（CC 和代码都适用）：**
- Scanner 代码：用 `OutputPaths.Reports` / `OutputPaths.Analysis` / `OutputPaths.Misc`（`NumDesTools.Scanner/OutputPaths.cs`），不要 hardcode 路径
- 主插件代码：用 `OutputPaths.Reports` / `OutputPaths.Analysis` / `OutputPaths.Misc`（`NumDesTools/OutputPaths.cs`）
- 新功能需要新子目录时：在对应 `OutputPaths.cs` 加一个属性，不要直接 `Path.Combine`
- CC（我）手动写文件：直接写到 `OutputRootPath` 对应子目录，**写完立即执行**：
  ```bash
  git -C "C:\Users\cent\Documents\NumDesOutput" add -A
  git -C "C:\Users\cent\Documents\NumDesOutput" diff --cached --quiet || git -C "C:\Users\cent\Documents\NumDesOutput" commit -m "[描述] 说明内容"
  ```

**不纳入 OutputRootPath 的：**
- `Documents\workspace\plugin.log` — 运行日志
- `Documents\NumDesTools\Config\` — 飞书工作流配置，路径被其他系统依赖
- `AppData\NumDesTools\` — 个人操作历史
- `C:\tmp\` — 原始 ADB 数据和中间产物
- `M1Work\` 写回 — 游戏配置表，不是插件产出

### 源文件编码

**C# 源文件必须用 UTF-8 保存**，禁止 GBK/ANSI。GBK 编码的中文提交后在 Git 里变成乱码（`"链"` → `"��"`），后续修复极易还原错误引发业务 bug（曾因此将类型判断 `"链"` 错误还原为 `"合"`，导致 ID 计算逻辑反转）。

- Visual Studio：文件 → 高级保存选项 → UTF-8
- 提交前 `git diff` 检查中文是否正常，出现方块乱码立即排查编码再提交

### Excel 读写规范

- **读取 xlsx** → EPPlus（`OfficeOpenXml`）。参考 `NumDesTools.Scanner/ExcelReader.cs`。
- **写入 xlsx** → EPPlus（`OfficeOpenXml`）。参考 `NumDesTools.Scanner/ActivityWriter.cs`（追加行）和 `NumDesTools.Scanner/LteMapWriter.cs`（新建带样式/列宽/行高的 sheet）。
- 每个入口必须先调用 `ExcelPackage.License.SetNonCommercialPersonal("NumDesTools")`。
- MiniExcel 已在依赖中但**不用于本项目**，不要引入。

## 多模型自动路由

**用户无需指定模型。CC（sonnet/opus）作为调度器，根据任务类型自动选择最优模型执行，不向用户说明切换过程，直接给出结果。**

LiteLLM 网关地址：`https://litellm.solotopia.net/v1/chat/completions`，Key：见 `ANTHROPIC_AUTH_TOKEN` 环境变量。

### 任务→模型映射

| 任务类型 | 执行方式 | 原因 |
|---------|---------|------|
| 多语言翻译（游戏文案、UI 字符串） | Workflow → `deepseek-v4-flash` | 满分且最省，本地化最地道 |
| 批量数据处理、格式化、简单分类 | Workflow → `deepseek-v4-flash` | 便宜快，不需要推理 |
| 中文游戏配置分析（数值合理性、配置审查） | 当前会话模型（sonnet/opus） | 有项目 Memory 和设计规范加成，实测优于 qwen |
| 复杂推理、数值系统设计、留存分析 | Workflow → `claude-opus-4-8` | 能识别前提矛盾，系统思维最强 |
| 竞品分析（多源采集+综合） | Workflow → fan-out `deepseek-v4-flash` 采集，`claude-opus-4-8` 综合 | 脏活便宜做，深度分析用强模型 |
| 需要读写文件/代码/git 的任务 | 当前会话模型（sonnet/opus）直接处理 | 其他模型没有 CC 工具访问权限 |
| 跨类型混合任务、多轮对话、无法归类 | 当前会话模型（sonnet/opus）兜底 | 保持上下文连贯 |

### 多 Agent / Workflow 规则

- **至少 1 个 agent 必须是 `claude-sonnet-4-6` 或 `claude-opus-4-8`**，承担协调、验证或综合角色。
- 其他 agent 可用任意 LiteLLM 模型，按任务类型从映射表选取。
- 不确定质量的结果必须经过 sonnet/opus 二次核查后再返回用户。

### 触发原则

- 任务**不需要 CC 工具（文件、shell、git）** 且匹配上表中有 Workflow 的行 → 主动调用 Workflow。
- 任务**需要 CC 工具** → 用当前会话模型，LLM 密集的子步骤（如翻译某段文本、分析某个字段含义）可在 Workflow 子 agent 中委托。
- **不要等用户说"用 deepseek"或"开 Workflow"**，这是自动行为。
