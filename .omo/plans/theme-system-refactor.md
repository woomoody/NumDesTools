# Theme System Refactor — 全量 UI 架构升级 + 终极美化

## 设计目标

1. **ThemeDictionaries XAML**：12 个语义色从 C# 代码移到 XAML ThemeDictionaries，深浅切换 WPF 自动解析，零 GC，改色不需编译
2. **Fluent 2 Design Token 规范**：命名、分层、状态 token 与 Windows 11 设计体系对齐
3. **窗口合并**：同类窗口合并为公共资产，减少 4 个文件
4. **技术债务清理**：遗留的 `isDark` 代码判断、非 FluentWindow、元素绑定 → 全改 DynamicResource

---

## 阶段执行计划

### Phase 0 — 分支 + 基础

- [ ] `git checkout -b refactor/theme-system`
- [ ] 创建 todo 列表

### Phase 1 — ThemeDictionaries XAML（基石，所有后续依赖此）

**文件**：新建 `UI/NumDesTools.ThemeDictionaries.xaml`

**内容**：
- `<ResourceDictionary.ThemeDictionaries>` 含 Dark/Light 两套 Color
- 12 个语义色（Ours/Theirs/Conflict/History/AiSuggestion 各 Text+Bg，OursActionBackground） + SemanticButtonTextColor
- 全局 13 个 `SolidColorBrush` 通过 `Color="{DynamicResource ...}"` 引用 ThemeDictionaries 色
- 参考 Fluent 2 命名：`OursTextColor` / `OursTextBrush` 等

**修改**：`MahAppsHelper.cs`
- 删除 `ApplySemanticBrushes()` 方法（12 个 new SolidColorBrush + hex 硬编码）
- 删除 `Changed` 事件中调用 `ApplySemanticBrushes()`
- 改为 `app.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = "UI/NumDesTools.ThemeDictionaries.xaml" })`
- 保留 `ThemesDictionary` + `ControlsDictionary` 初始化、`SystemThemeWatcher`、`Changed` 事件中的诊断日志

**验证**：构建 0 错，启动后语义色正常显示，深浅切换正常

### Phase 2 — 清理 GitExportSelectWindow 代码债务

**文件**：`UI/GitExportSelectWindow.xaml.cs`

**当前问题**：3 处 `isDark` + 硬编码 hex 色值（第 162-253 行，Ours/Theirs/History 文字色 + badge 背景色 + 非配置表文字色）

**改为**：使用 `TryFindResource("OursTextBrush")` 等 DynamicResource 引用自定义语义画刷，或直接 XAML 中绑定

**验证**：构建 0 错，所有 badge 颜色在深/浅模式下正确

### Phase 3 — 清理 BatchReplacePanel 元素绑定

**文件**：`UI/BatchReplacePanel.xaml` + `.cs`

**当前问题**：使用 `BgMain/FgMain/FgDim/AccentCol/BorderCol/BgInput/BgPanel` 7 个元素绑定属性，这些属性在 code-behind 中通过 `isDark` 返回硬编码颜色

**改为**：全部替换为 `DynamicResource` 直接引用 wpf-ui 主题画刷 + 自定义语义画刷
- `BgMain` → `ApplicationBackgroundBrush`
- `BgPanel` → `LayerFillColorDefaultBrush`
- `FgMain` → `TextFillColorPrimaryBrush`
- `FgDim` → `TextFillColorSecondaryBrush`
- `AccentCol` → `AccentFillColorDefaultBrush`
- `BorderCol` → `ControlElevationBorderBrush`
- `BgInput` → `ControlFillColorInputActiveBrush`

**验证**：构建 0 错，运行后颜色正常

### Phase 4 — 迁移 FileLockWindow 到 FluentWindow

**文件**：`FileLockWindow.xaml` + `.cs`

**当前问题**：使用原始 `Window`，硬编码 `#F5F5F5` 背景，无主题支持

**改为**：`ui:FluentWindow` + `MahAppsHelper.EnsureInitialized()` + `SetExcelOwner()` + `DynamicResource` 主题画刷

**验证**：构建 0 错，深色模式下可读

### Phase 5 — 合并 InputBoxDialog + PasswordDialog → InputDialog

**新建**：`UI/InputDialog.xaml` + `.cs`

**删除**：`InputBoxDialog.xaml/.cs` + `PasswordDialog.xaml/.cs`

**设计**：
```csharp
public enum InputDialogMode { Text, Password, MultiLine }
public partial class InputDialog : FluentWindow
{
    public static string ShowText(string title, string prompt, string defaultValue = "") => ...
    public static string ShowPassword(string title, string prompt) => ...
}
```

**验证**：构建 0 错，调用方更新为 `InputDialog.ShowText`/`ShowPassword`

### Phase 6 — 合并 SettingsForm 窗口

**新建**：`UI/SettingsFormWindow.xaml` + `.cs`（数据驱动）

**删除**：`ActivityBackupSettingsWindow.xaml/.cs` + `XlsxSyncSettingsWindow.xaml/.cs`

**设计**：
```csharp
public record SettingsField(string Label, string? BrowseButtonText, string? Tooltip);
public partial class SettingsFormWindow : FluentWindow
{
    public SettingsFormWindow(string title, string description, params SettingsField[] fields) => ...
}
```

**验证**：构建 0 错，两个调用方正常

### Phase 7 — 通用 ConfirmDialog

**新建**：`UI/ConfirmDialog.xaml` + `.cs`（基于 ActivityBackupReportWindow 提取）

**保留**：`ActivityBackupReportWindow.xaml`（改为包装 ConfirmDialog 调用）

**替换**：全插件内 `MessageBox.Show` 调用 → `ConfirmDialog.Show`/`ConfirmDialog.ShowResult`

**验证**：构建 0 错，所有确认/提示框外观一致

### Phase 8 — FluentWindowBase 基类

**新建**：`UI/FluentWindowBase.cs`

**内容**：
- 自动调用 `MahAppsHelper.EnsureInitialized()` + `SetExcelOwner(this)`
- 自动挂 Esc 关闭
- 自动在 XAML Row 0 构建 TitleBar
- 所有窗口继承此类，删除重复代码

**注意**：wpf-ui 的 FluentWindow 是 sealed 吗？需要先验证是否能继承。如果不能，改为扩展方法 + 部分类模式。

**验证**：构建 0 错

### Phase 9 — Fluent 2 Design Token 命名规范化

**文件**：`UI/NumDesTools.ThemeDictionaries.xaml` + 所有引用 XAML

**Fluent 2 命名规范**：
```
<Category><SubCategory><Property><State>
OursTextColor          →  OursTextColor (保持)
OursBackgroundBrush    →  OursBackgroundFillColorDefaultBrush
SemanticButtonTextBrush → TextOnSemanticFillColorPrimaryBrush
```

**实际改动**：仅对 `NumDesTools.ThemeDictionaries.xaml` 内部命名做规范化，XAML 中引用保持兼容（可用旧 key 别名）

### Phase 10 — 全量 XAML 一致性审查

**审查范围**：全部 28 个 XAML 文件

**检查项**：
- 所有 `Background`/`Foreground`/`BorderBrush` 是否使用 `DynamicResource`
- 是否有遗漏的硬编码 hex/named 颜色
- Button 是否用了 `Appearance` 而非 `Background`
- 状态（hover/selected/disabled）是否使用 `SubtleFillColor*` / `ControlFillColor*` 等正确 token
- 是否有 `FontFamily` 硬编码（如 `FontFamily="Consolas"` 应保持）

---

## 验证清单

- [ ] 构建 0 error
- [ ] 73 个 CtpThemeAdaptationTests 全部通过
- [ ] 深色模式：所有窗口可读，语义色正确
- [ ] 浅色模式：所有窗口可读，语义色正确
- [ ] 切换主题时：颜色即时更新，无闪烁
- [ ] 文件变更：减少 4 个文件（合并后），总代码量缩减