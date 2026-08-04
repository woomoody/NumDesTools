# XlsxEditor 优化计划

## 分支
`xlsxeditor-optimization`

## 目标
对 NumDesTools.XlsxEditor 进行全面优化，覆盖：死代码清理、数字排序、撤销/重做安全、筛选 UX、增量写回路径、测试。

## 任务清单

### P0 核心清理
- [ ] **清理 OoxmlLazyReader 死代码**：`OoxmlLazyReader.cs` 中 `ReadRows` 已废弃（实际加载走 `ColumnStoreExcelLoader`），只保留 `ReadDimension` 或整文件删除
- [ ] **验证全量加载内存压力**：用 6.5 万行 × 85 列文件实测 ColumnStore 内存占用，必要时加窗口化加载

### P1 功能完善
- [ ] **数字列类型推断 + 排序**：`ColumnTypeDetector` 接入 `VirtualizingSortableView` 排序比较器，数字列走数字比较，fallback 字典序
- [ ] **筛选 UX 升级**：`VirtualizingSortableView` 筛选直接对 `_rowOrder` 做 `Where`，不走 DataView RowFilter
- [ ] **撤销/重做索引安全**：增删行列时清空 undo/redo 栈，或改成相对列名/行号索引

### P2 架构打磨
- [ ] **增量写回覆盖所有保存路径**：检查 `SaveCurrentFileAsync` 调用链，确保优先走 `IncrementalOoxmlWriteBack.TryWrite`，fallback 到 `ExcelWriteBack.Write`
- [ ] **MainWindow.xaml.cs 死代码/技术债清理**：扫描未使用的字段、方法，清理 DataTable 时代遗留路径

### P3 测试与质量
- [ ] **增量写回边界测试**：空文件、单行文件、混合 inlineStr/sharedStrings、带公式/图表/透视表
- [ ] **UI 层基础测试**：打开文件 → 编辑 → 保存 → 验证文件内容