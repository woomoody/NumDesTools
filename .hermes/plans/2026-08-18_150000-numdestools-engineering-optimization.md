# NumDesTools 工程优化实施计划

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** 在不改动 ExcelToLua（导表）和“不支持/扩展 xlsm、xlsx 搜索范围”的前提下，分阶段提升 NumDesTools 的索引一致性、线程安全、Excel 生命周期安全、数据写回安全和测试可靠性。

**Architecture:** 先修 P0 数据正确性与宿主生命周期问题，再治理 ExcelIndex 后台构建与缓存一致性，随后收敛搜索/Scanner/冲突处理的公共模型，最后拆分大文件和补充诊断面板。所有后台任务遵循“Excel 主线程采集纯数据 → 后台纯托管计算 → Excel 主线程提交结果”的边界。

**Tech Stack:** C#/.NET 9、ExcelDna、EPPlus、MiniExcel、WPF/WinForms CTP、xUnit、FileSystemWatcher、CancellationToken、ConcurrentQueue/Channel。

---

## 明确不做的范围

本计划明确排除：

1. **ExcelToLua / 导表功能不优化、不重构、不改写入策略。** 现有导表相关风险只记录，不纳入本轮实施。
2. **不新增或扩展全局搜索的 xlsm/xlsx 文件类型范围。** 继续保持当前项目既定搜索范围；不改 watcher、文件收集器和 fallback 的扩展名策略。
3. 不处理本次工作区已有的 DSH/Kilo/lazykey/lazymodel/lazytoken/git-tui 动态修改；这些变化必须与插件优化分开提交。
4. 不在未确认业务预期前修改两个失败测试对应的生产算法。

---

## 当前基线

### 工程规模

- `NumDesTools/`：约 86,729 行 C#
- `NumDesTools.Core/`：约 16,820 行 C#
- `NumDesTools.Tests/`：约 11,118 行 C#
- `NumDesTools.Scanner/`：约 2,016 行 C#

### 已确认测试基线

```text
dotnet test NumDesTools.Core.Tests/NumDesTools.Core.Tests.csproj -c Debug --no-restore
```

结果：47 passed，1 failed。

完整测试审查还发现：

- `LteCoreUnitTests.Arr_ProducesPairedList`
  - expected `[a,1],[b,2]`
  - actual `[a,1,a],[b,2,b]`
- `ColumnStoreExcelLoaderTests.Load_Item_KnownLiteralValues`
  - expected `7616068834`
  - actual `7616068417`

这些先作为基线问题隔离，第一阶段需要确认是生产回归还是过期测试夹具。

### 主要高风险区域

- `NumDesTools.Core/ExcelIndex/ExcelIndexManager.cs`
- `NumDesTools.Core/ExcelIndex/ExcelIndexBuilder.cs`
- `NumDesTools/PubMetToExcelFunc.cs`
- `NumDesTools/NumDesAddIn.Buttons.cs`
- `NumDesTools/NumDesAddIn.cs`
- `NumDesTools/NumDesCTP.cs`
- `NumDesTools/Battle/DotaLegendBattleParallel.cs`
- `NumDesTools/Battle/DotaLegendBattleSerial.cs`
- `NumDesTools/AutoInsert/ExcelDataAutoInsertLanguage.cs`
- `NumDesTools.Core/ConflictResolver/ExcelConflictDiffer.cs`
- `NumDesTools.Core/ConflictResolver/ConflictApplier.cs`
- `NumDesTools.Core/Scanner/ExcelReader.cs`

---

# Phase 0：基线、边界和回归保护

## Task 0.1：冻结当前测试基线

**Files:**

- Modify: `.hermes/plans` only for this plan; implementation does not begin here.
- Test: `NumDesTools.Core.Tests/`
- Test: `NumDesTools.Tests/`

**Steps:**

1. 记录 Core、主 Tests、Scanner Tests 的当前通过/失败/跳过数量。
2. 单独运行两个已知失败测试，保存完整 assertion 差异。
3. 将动态生成文件、`.xll`、`.debug-journal.md` 与源码改动分离。
4. 后续每个任务只允许增加新测试失败，不得把基线失败误报为回归。

**Validation:**

```bash
dotnet test NumDesTools.Core.Tests/NumDesTools.Core.Tests.csproj -c Debug --no-restore
dotnet test NumDesTools.Tests/NumDesTools.Tests.csproj -c Debug --no-restore
```

## Task 0.2：确认优化边界

**Steps:**

1. 为索引、后台任务、Excel 状态、冲突处理、AutoInsert 建立修改清单。
2. 明确不触碰 `ExcelToLua/`。
3. 明确不修改全局搜索的 xlsm/xlsx 范围。
4. 每个后续 PR/commit 只覆盖一个主题，避免把宿主修复和业务行为混在一起。

---

# Phase 1：P0 数据正确性和 Excel 生命周期

## Task 1.1：修复跨项目索引复用风险

**Files:**

- Modify: `NumDesTools.Core/ExcelIndex/ExcelIndexManager.cs`
- Modify: `NumDesTools/PubMetToExcelFunc.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Test: `NumDesTools.Tests/ExcelIndexBuilderIncrementalTests.cs`
- Add/Modify: `NumDesTools.Core.Tests/ExcelIndexManagerTests.cs`

**Design:**

1. 增加规范化 root 方法：`Path.GetFullPath`、目录分隔符统一、大小写不敏感比较。
2. 每次 `StartForPath` 递增 `generation`。
3. 搜索快速路径必须验证：
   - 当前 `rootPath` 与 `ExcelsRoot` 相同；
   - index generation 与当前 generation 相同；
   - index 未被取消。
4. 项目切换期间不复用旧项目索引；返回明确的“索引切换中”状态或走当前项目的既定 fallback。
5. 构建完成发布前再次检查 root/generation/token，旧项目结果直接丢弃。

**Tests:**

- 项目 A 索引完成后切换 B，B 搜索不得返回 A 路径。
- A→B→A 快速连续切换时最终只发布当前 generation。
- root 仅大小写/斜杠差异时应视为同一项目。

## Task 1.2：建立统一 Excel 状态恢复 Scope

**Files:**

- Add: `NumDesTools/ExcelStateScope.cs`
- Modify: `NumDesTools/NumDesAddIn.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Test: `NumDesTools.Tests/ExcelStateScopeTests.cs`

**Design:**

1. 进入操作前保存 `Calculation`、`ScreenUpdating`、`EnableEvents`、必要时 `DisplayAlerts`。
2. 设置临时状态。
3. `Dispose` 时按原值恢复，不恢复固定默认值。
4. 恢复过程独立 `try/catch` 并写 `PluginLog`。
5. Scope 只在 Excel 主线程创建/释放，禁止跨线程持有 COM。

**Validation:**

- 用户原本为手动计算时，操作后仍是手动计算。
- 异常路径也恢复原状态。
- 恢复失败有日志但不覆盖原始异常。

## Task 1.3：清理战斗模块并行共享状态

**Files:**

- Modify: `NumDesTools/Battle/DotaLegendBattleParallel.cs`
- Modify: `NumDesTools/Battle/DotaLegendBattleSerial.cs`
- Add/Modify: `NumDesTools.Tests/DotaLegendBattle*Tests.cs`

**Design:**

1. Excel 主线程一次性复制输入为纯托管数组/record。
2. `Parallel.For` 每个任务返回局部统计对象。
3. 主线程或单独合并阶段统一累加。
4. 删除静态 `Worksheet`/`Range` 长期引用。
5. 删除对 `ExcelDnaUtil.Application` 的析构 `Dispose`。
6. 写回只在 Excel 主线程执行。

**Validation:**

- 相同输入重复运行结果完全一致。
- 并行结果与 Serial 结果逐项相等。
- Excel 关闭/切换时后台任务取消，不访问旧 COM 对象。

## Task 1.4：修复后台任务与 CTP 生命周期

**Files:**

- Modify: `NumDesTools/NumDesCTP.cs`
- Modify: `NumDesTools/NumDesAddIn.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Add/Modify: `NumDesTools.Tests/CtpLifecycleTests.cs`

**Design:**

1. 所有 CTP 创建/登记/显示放在同一个 `QueueAsMacro` 事务中。
2. 每个 CTP 引入 `Creating/Visible/Closing/Closed` 状态。
3. `AutoClose` 最终调用 `NumDesCTP.DisposeAll()`。
4. Dispose 幂等，释放失败记录具体对象和异常。
5. 所有 `Task.Run` 保存 Task，并绑定 workbook/CTP cancellation token。
6. 用 `async Task` 替代裸 `ContinueWith`；回主线程前检查 fault/cancel。

**Validation:**

- 快速重复打开同名 CTP 不重复创建。
- 工作簿关闭后延迟回调不会访问已释放 CTP。
- 插件卸载后无遗留 CTP/ElementHost。

---

# Phase 2：索引后台构建一致性

## Task 2.1：单 worker 串行化索引重建

**Files:**

- Modify: `NumDesTools.Core/ExcelIndex/ExcelIndexManager.cs`
- Add/Modify: `NumDesTools.Core.Tests/ExcelIndexManagerTests.cs`

**Design:**

1. watcher 只把变化加入 `ConcurrentDictionary`/Channel。
2. 只允许一个 build worker。
3. worker 执行：快照 pending → 构建 → 发布 → 再检查 pending。
4. 构建期间新事件不会丢失，也不会启动第二个并发 builder。
5. 磁盘缓存临时文件名包含 generation 或 GUID。
6. 写入采用临时文件后原子替换。

**Validation:**

- 构建期间连续修改同一个文件，最终索引包含最后版本。
- 连续修改多个文件只触发合并后的重建。
- 不出现旧构建覆盖新构建。

## Task 2.2：完善 watcher 过滤和异常恢复

**Files:**

- Modify: `NumDesTools.Core/ExcelIndex/ExcelIndexManager.cs`
- Modify: `NumDesTools.Core/SelfExcelFileCollector.cs`（只统一当前既有扩展范围，不扩展到用户明确排除的范围）
- Test: `NumDesTools.Tests/ExcelIndexBuilderIncrementalTests.cs`

**Design:**

1. 入口过滤临时文件、排除目录、删除后不存在路径。
2. 监听 `Error`，特别是 `InternalBufferOverflowException`。
3. watcher 溢出时排队一次完整重建。
4. watcher/Timer 在项目切换和插件关闭时释放。
5. 事件处理只做轻量入队，不在 FileSystemWatcher 线程读取 Excel。

## Task 2.3：索引文件集与 cell hit 分离

**Files:**

- Modify: `NumDesTools.Core/ExcelIndex/ExcelIndexBuilder.cs`
- Modify: `NumDesTools.Core/ExcelIndex/ExcelSearchIndex.cs`
- Test: `NumDesTools.Tests/ExcelIndexBuilderIncrementalTests.cs`
- Test: `NumDesTools.Tests/ExcelSearchIndexContainsTests.cs`

**Design:**

1. 扫描文件开始就登记文件相对路径、fingerprint、状态。
2. 每个有效 Sheet 即使没有 cell hit 也登记。
3. cell 倒排命中与文件/Sheet 清单分离。
4. 空文件、空 Sheet、只有标题的 Sheet 不再被当作“未缓存”。
5. 读取失败文件保留旧索引并标记失败，不静默删除旧结果。

## Task 2.4：轻量 fingerprint 优化

**Files:**

- Modify: `NumDesTools.Core/ExcelIndex/ExcelIndexBuilder.cs`
- Modify: `NumDesTools.Core/ExcelIndex/ExcelSearchIndex.cs`
- Test: `NumDesTools.Tests/ExcelIndexBuilderIncrementalTests.cs`

**Design:**

1. 默认先比较 `Length + LastWriteTimeUtc`。
2. fingerprint 未变则不读取完整 MD5。
3. fingerprint 变化才计算 MD5/扫描。
4. 记录构建耗时、hash 耗时、扫描耗时和变更文件数量。
5. 保留定期/手动全量校验入口作为自愈机制。

---

# Phase 3：搜索服务和结果诊断

## Task 3.1：统一搜索结果与错误状态

**Files:**

- Add: `NumDesTools.Core/ExcelSearch/ExcelSearchResult.cs`
- Add: `NumDesTools.Core/ExcelSearch/ExcelSearchStatus.cs`
- Modify: `NumDesTools/PubMetToExcelFunc.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Test: `NumDesTools.Tests/ExcelSearchResultTests.cs`

**Design:**

统一区分：

```text
Matched
NoMatch
IndexUpdating
ReadFailed
InvalidFormat
PermissionDenied
Cancelled
```

结果中增加：

```text
IndexGeneration
IsStale
ScannedFiles
MatchedFiles
FailedFiles
SkippedFiles
```

UI 不再把“读取失败”显示成“没有匹配”。

## Task 3.2：搜索索引状态模型

**Files:**

- Modify: `ExcelIndexManager.cs`
- Add: `NumDesTools.Core/ExcelIndex/ExcelIndexStatus.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Add/Modify: UI 状态窗口或现有 CTP

**功能：**

显示：

```text
当前项目
缓存路径
索引版本
最后构建时间
文件数量
待处理文件数量
失败文件数量
后台状态
```

支持：

```text
立即增量重建
立即全量重建
清理缓存
重试失败文件
```

## Task 3.3：统一慢速 fallback 的取消和进度

**Files:**

- Modify: `NumDesTools/PubMetToExcelFunc.cs`
- Modify: `NumDesTools/NumDesAddIn.Buttons.cs`
- Add: `NumDesTools.Core/ExcelSearch/ExcelSearchProgress.cs`
- Test: `NumDesTools.Tests/ExcelSearchCancellationTests.cs`

**Design:**

1. 索引未就绪时仍保持既有行为，不改变用户确定的搜索范围。
2. 慢速扫描增加 CancellationToken。
3. UI 显示已扫描/总文件数。
4. 搜索按钮支持取消。
5. 统一记录失败文件，不静默吞掉。

---

# Phase 4：数据写回安全（不包含 ExcelToLua）

## Task 4.1：多语言 AutoInsert 显式配对和事务计划

**Files:**

- Modify: `NumDesTools/AutoInsert/ExcelDataAutoInsertLanguage.cs`
- Add: `NumDesTools/AutoInsert/LanguageInsertPlan.cs`
- Add/Modify: `NumDesTools.Tests/AutoInsertLanguagePlanTests.cs`

**Design:**

1. 读取表格后先验证偶数/配对关系。
2. 验证表名、模板 ID、列数和行类型。
3. 先生成纯托管 `LanguageInsertPlan`。
4. 所有模板验证通过后才删除/插入。
5. 写入临时 xlsx，成功后原子替换。
6. 合并目前重复的三套行处理逻辑。
7. 未知文件名/列映射直接报错，不默认使用第 0 列。

## Task 4.2：冲突工具重复 key 显式阻断

**Files:**

- Modify: `NumDesTools.Core/ConflictResolver/ExcelConflictDiffer.cs`
- Modify: `NumDesTools.Core/ConflictResolver/ConflictApplier.cs`
- Add/Modify: `NumDesTools.Core.Tests/ExcelConflictDifferTests.cs`
- Add/Modify: `NumDesTools.Core.Tests/ConflictModelsTests.cs`

**Design:**

1. diff 前检测双方重复 key。
2. 重复 key 作为结构冲突返回。
3. 默认禁止自动 apply。
4. UI 显示两侧重复行号。
5. 将 key 列、数据起始行、表头结构从硬编码提升为显式表规则。
6. 读取异常必须变成不可 apply 的错误状态。

## Task 4.3：Scanner 缓存失效和诊断

**Files:**

- Modify: `NumDesTools.Core/Scanner/ExcelReader.cs`
- Modify: `NumDesTools.Scanner/*.cs`
- Add/Modify: `NumDesTools.Core.Tests/ScannerTests.cs`

**Design:**

1. 缓存 key 增加 Length/LastWriteTimeUtc 或 fingerprint。
2. 外部文件变化后自动失效。
3. `cell.Text` 与原始值语义明确分开。
4. 前导零、数字、日期、百分比增加测试。
5. 重复字段名不再静默覆盖，输出诊断。

---

# Phase 5：UI 和模块结构整理

## Task 5.1：拆分 `NumDesAddIn.Buttons.cs`

**目标文件：**

- `NumDesTools/ExcelSearchCommands.cs`
- `NumDesTools/ActivityCommands.cs`
- `NumDesTools/ConflictCommands.cs`
- `NumDesTools/ExportCommands.cs`
- `NumDesTools/UiResultPresenter.cs`

**原则：**

Ribbon handler 只做：

```text
获取上下文
调用服务
提交 UI 结果
```

不在按钮类里直接实现文件扫描、并行、Excel 写入。

## Task 5.2：拆分 `PubMetToExcelFunc.cs`

**目标文件：**

- `ExcelGlobalSearchService.cs`
- `ExcelIdSearchService.cs`
- `ExcelSheetSearchService.cs`
- `ExcelWriteBackService.cs`
- `ExcelFileEnumerator.cs`

本阶段不扩展 xlsm/xlsx 搜索范围，只做结构拆分和公共逻辑复用。

## Task 5.3：统一后台任务生命周期

**Files:**

- Add: `NumDesTools/PluginOperationCoordinator.cs`
- Modify: `NumDesAddIn.cs`
- Modify: `NumDesAddIn.Buttons.cs`
- Modify: AI panels and CTP manager

**Capabilities:**

```text
Start(operationId, workbookId)
Cancel(operationId)
CancelForWorkbook(workbookId)
CancelAll()
ObserveFaults()
```

每个操作记录：

```text
开始时间
结束时间
当前阶段
已处理数量
异常
取消状态
```

---

# Phase 6：测试和交付门禁

## Task 6.1：先处理两个基线失败

**Files:**

- `NumDesTools.Core.Tests/LteCoreUnitTests.cs`
- 对应生产实现（确认后再改）
- `NumDesTools.Tests/ColumnStoreExcelLoaderTests.cs`
- 对应 fixture/生产实现（确认后再改）

要求：

1. 先确认业务期望。
2. 不通过“改 expected”掩盖生产回归。
3. 每个失败补充最小复现说明。

## Task 6.2：增加测试门禁

每次涉及 Core/插件逻辑：

```bash
dotnet build NumDesTools.Core/NumDesTools.Core.csproj -c Debug --no-restore
dotnet test NumDesTools.Core.Tests/NumDesTools.Core.Tests.csproj -c Debug --no-restore
dotnet test NumDesTools.Tests/NumDesTools.Tests.csproj -c Debug --no-restore
```

涉及 Release/发布：

```bash
dotnet build NumDesTools.sln -c Release --no-restore
```

涉及索引：

```text
索引单项目测试
索引跨项目切换测试
watcher连续变更测试
缓存加载/过期测试
取消测试
```

涉及写回：

```text
异常不产生半文件
重复 key 阻断 apply
模板缺失不删除旧数据
```

## Task 6.3：日志和诊断门禁

统一检查：

```text
不允许裸 catch 静默吞错
后台任务必须可观察
所有失败文件带路径
所有缓存重建带 generation
所有 Excel 状态恢复按原值
```

---

# 风险和取舍

## 最大风险

- ExcelDna/COM 线程边界改动可能影响大量旧功能；必须小步拆分。
- AutoInsert 和 ConflictApplier 属于写回逻辑，必须先加 fixture/计划测试再改。
- 索引 generation/worker 改造可能改变搜索时序，需要保留旧缓存优先返回行为。
- 不扩展 xlsm/xlsx 搜索范围意味着部分文件仍不会进入索引，这是本计划的明确约束，不当作遗漏。

## 先不做的内容

- ExcelToLua 重构和原子写入：用户明确排除。
- 扩大搜索扩展名：用户明确排除。
- 重新设计整个插件 UI：先做状态和错误可见性，不做视觉大改。
- 大规模替换 EPPlus/MiniExcel：先统一边界和测试，再评估。

---

# 推荐执行顺序

```text
Phase 0  基线与边界
Phase 1  P0：跨项目索引隔离、Excel 状态、COM/并行、CTP 生命周期
Phase 2  索引单 worker、generation、watcher 一致性、fingerprint
Phase 3  统一搜索结果、stale 状态、取消和进度
Phase 4  AutoInsert 配对、冲突重复 key、Scanner 缓存
Phase 5  拆大文件、统一后台操作协调器
Phase 6  全量测试门禁和发布验证
```

**完成标准：**

- P0 问题全部有回归测试；
- 索引不跨项目复用；
- 后台重建单 worker 且不会丢事件/旧覆盖新；
- Excel 原始状态始终恢复；
- COM 不跨线程持有；
- AutoInsert/Conflict apply 在结构不安全时明确阻断；
- Core 与主测试的基线失败已归因并处理；
- ExcelToLua 和搜索扩展名保持不变。
