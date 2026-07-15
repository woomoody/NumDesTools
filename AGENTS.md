# AGENTS.md

给任何在这个仓库里写代码的 agent（Omp、Claude Code 等）或人看的通用工程规则，跟用的是哪个工具无关。
Claude Code 专属的行为约定（多模型路由等）在 `CLAUDE.md`，不重复写在这里。

## Build

```bash
# Debug build
dotnet build NumDesTools.sln -c Debug

# Release build (also packs .xll via Excel-DNA post-build → packFromBin/)
dotnet build NumDesTools.sln -c Release

# Run tests
dotnet test NumDesTools.sln
```

Post-build 调用 `packFromBin\ReNamePack.bat`，把打包出的 XLL 重命名为 `NumDesToolsPack64.xll` / `NumDesToolsPack.xll`。

改完代码必须**0 error**才算完成。命名遵循 **ReSharper** 规则，格式化用 **CSharpier**。

## Architecture

三个项目在一个 solution 里：

| Project | Type | Purpose |
|---------|------|---------|
| **NumDesTools** | Excel XLL add-in (net9.0-windows) | 主插件：ribbon、UDF、UI 窗口 |
| **NumDesTools.Scanner** | Console app (net9.0-windows) | 活动配置表校验 CLI + 飞书集成 |
| **NumDesTools.Tests** | xUnit test (net9.0-windows) | LTE 计算、scanner、map 操作的单测 |

### NumDesTools — 关键模块

**入口：** `NumDesAddIn.cs` — 实现 `IExcelAddIn` + `ExcelRibbon`；持有全局 `App`（Excel.Application）引用；所有 ribbon 按钮点击都走带防抖（500ms）的字典分发。

**Ribbon 定义：** `RibbonUI.xml` — XML 布局。Label/状态由 `GlobalVariable.cs` 的 toggle 字符串驱动。

**UDF：** `ScUDFs.cs`（2100+ 行）— 所有暴露给 Excel 的 `[ExcelFunction]` 函数：FindKey*、Trans2Array*、JSON 导出辅助、游戏专用计算（LTE 链、Alice、Dota 模拟）。

**配置：** `Config/GlobalVariable.cs` — 加载/保存 `Documents\NumDesGlobalKey.json`；存路径、API key、UI toggle 标签、spotlight 模式、git root。

**ConflictResolver 命名空间** — 基于 Git 的 Excel 冲突解决 UI：
- `ExcelConflictDiffer.cs` — 把两个 Excel 文件的 diff 拆成 `FileDiff` / `SheetDiff` / `RowConflict` / `CellConflict` 模型
- `ConflictModels.cs` — 数据模型，带 `IsResolved`、拖选 `IsSelected`、`INotifyPropertyChanged`
- `ConflictApplier.cs` — 写回选中的值，调 `git add`，追加到 `.git/MERGE_MSG`
- `UI/ExcelConflictWindow.xaml(.cs)` — WPF 窗口：拖选多选、筛选选择、选区操作栏、apply 前的未解决检查

**单元格高亮：** `CellHighlighter.cs` — `ViewportHelper.GetViewportRange`（合并所有 `win.Panes[i].VisibleRange`，跟 UsedRange 求交集）被 `CellHighlighter`（Find/FindNext 同值）和 `CellSpotlightHighlighter`（填色模式行列高亮）共用。`CrosslightOverlay.cs` 处理十字线覆盖层模式。

**Advance 命名空间：** `ExcelDataToDB.cs`、`IdPrefixIndex.cs` — 游戏数据提取和跨表索引。

**ExcelToLua 命名空间：** `ExcelReader.cs` — 把 Excel 表结构（第2行=列名，第3行=类型，第4行=中文标签）转成带类型定义的 Lua。

### 数据约定（Excel 表）

- Sheet 前缀 `c_*` = 客户端，`s_*` = 服务端，`#*` = 隐藏/元数据
- 第1行=表头，第2行=列名，第3行=类型（`int`/`string`/…），第4行=中文标签
- `#` 前缀列 = 注释/展示列；冲突解决器行头会显示这些

### Threading

长耗时操作用 `ExcelAsyncUtil.QueueAsMacro()`。不能直接阻塞 Excel UI 线程。

### Key dependencies

`ExcelDna.AddIn 1.9.0`、`EPPlus 8.2.0`、`MiniExcel 1.42.0`、`LibGit2Sharp 0.31.0`、`NLua`、`MathNet.Numerics`。

### 输出目录规范

所有值得保留的产出文件（xlsx 报告、md 分析文档、html、json 分析产物）统一写到 **`OutputRootPath`** 下，默认值为 `Documents\NumDesOutput\`（本地独立 git 仓库，不推送）。

**子目录规范：**

| 子目录 | 用途 |
|--------|------|
| `reports\` | xlsx/html 报告（竞品分析、地编信息、LTE 配置模版等） |
| `analysis\` | md 分析文档（竞品深度分析、设计规范、配置草稿等） |
| `misc\` | 插件偶发产出（溯源结果.xlsx、表格关系.json 等） |

**写文件规则：**
- Scanner 代码：用 `OutputPaths.Reports` / `OutputPaths.Analysis` / `OutputPaths.Misc`（`NumDesTools.Scanner/OutputPaths.cs`），不要 hardcode 路径
- 主插件代码：用 `OutputPaths.Reports` / `OutputPaths.Analysis` / `OutputPaths.Misc`（`NumDesTools/OutputPaths.cs`）
- 新功能需要新子目录时：在对应 `OutputPaths.cs` 加一个属性，不要直接 `Path.Combine`

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

- **默认方案：读写 xlsx 一律优先 EPPlus（`OfficeOpenXml`）**。参考 `NumDesTools.Scanner/ExcelReader.cs`、`NumDesTools.Scanner/ActivityWriter.cs`、`NumDesTools.Scanner/LteMapWriter.cs`。
- **只有明确属于“大量查询 / 高性能只读路径”时才用 MiniExcel**。典型场景：全表扫描、跨文件索引构建、全局搜索、diff / 冲突分析、找到目标后可提前 break 的流式读取。
- 判断标准：如果任务主要是“读很多、筛很多、尽快找到值”，优先考虑 MiniExcel；如果任务涉及写回、样式、图片、合并单元格、工作表结构、精确单元格控制，使用 EPPlus。
- 主插件里如果走 MiniExcel，统一传 `NumDesAddIn.OnOffMiniExcelCatches`。
- **MiniExcel 只用于读取，不用于写入**。
- 每个 EPPlus 入口必须先调用 `ExcelPackage.License.SetNonCommercialPersonal("NumDesTools")`。
- **不要为本项目新增 Python 生成 xlsx 的路径**。生成的文件结构不规范会触发 Excel 修复提示；xlsx 产物统一走 C# + EPPlus。
- 需要更细的步骤或示例时：Claude 侧参考 `.claude/skills/xlsx-rw/SKILL.md`，Codex/plugin 侧参考 `.agents/plugins/plugins/numdestools-codex/skills/numdestools-xlsx-rw/SKILL.md`。

### Secrets 与本地配置

- 带 token、key、cookie、header 凭据的配置文件不得提交到仓库。
- `.mcp.json` 是本地私有文件，仅供本机 Claude Code 使用；仓库只保留 `.mcp.example.json` 模板。
- Repo 内的 `.codex/config.toml` 只能放无密项目配置；带认证信息的 MCP 配置放到用户级 `~/.codex/config.toml`。
- 如果 secret 曾经进入 git 历史，除了从当前树移除之外，还应视情况做 token 轮换。

## 多 agent 协作

- Claude Code 与执行方（Omp）协作时，实时交接协议见 `docs/agent-handoff.md`
- 运行态交接文件统一放 `.remember/agent-handoff/<task-name>/`，不提交
- 值得长期保留的结论必须回写到版本化文件，不要只留在交接文件里
- 没有具体任务前不要预建 worktree；并行改动时再按 `docs/agent-handoff.md` 建

### 任务文档格式：一个任务一份 md，不滚雪球

给执行方（Omp）派活的任务文档（`Documents\NumDesOutput\codex-tasks\*.md`）**一个任务只维护一份文件，
永不因为纠错/追加需求而新开文件**（不要出现 task2 → task2b → task2c → task2d → task2e 这种
滚雪球式命名——这曾经真实发生过，导致同一个任务拆成5个文件，每次纠错都要重新贴一遍背景，读者
要按时间顺序翻好几个文件才能拼出"现在到底要做什么"，容易看错版本、漏看最新要求）。

固定的文档结构：

```markdown
# 任务N：<任务名>

> 格式约定：这份文档是任务N的唯一真相源。以后这个任务有新问题/新要求，不要再开新文件，直接改
> 这份文档——更新"当前待办"区块为最新内容，并在"历史日志"里追加一行记录。

## 当前待办（<日期> 更新）
<这次真正要做的事，写完整，不依赖读者先翻历史>

## 历史日志
- **<日期> <一句话摘要>**：<这轮改了什么/否定了什么>

## 参考资料
<调研结论、格式约定等稳定内容>
```

执行方接单/交付时：
- 有新问题或用户反馈要处理，**直接编辑对应任务的文档**，把"当前待办"替换成最新要做的事，旧
  内容提炼成一行塞进"历史日志"，不要创建新的 `taskNx.md` 文件。
- 一个任务彻底完成、不会再迭代时，可以删除任务文档本身，只保留它产出的调研报告/交付物说明
  （如果有独立价值）。

<!-- OPENWIKI:START -->

## OpenWiki

This repository uses OpenWiki for recurring code documentation. Start with `openwiki/quickstart.md`, then follow its links to architecture, workflows, domain concepts, operations, integrations, testing guidance, and source maps.

The scheduled OpenWiki GitHub Actions workflow refreshes the repository wiki. Do not hand-edit generated OpenWiki pages unless explicitly asked; prefer updating source code/docs and letting OpenWiki regenerate.

<!-- OPENWIKI:END -->
