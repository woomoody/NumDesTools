using System.IO;
using System.Text.RegularExpressions;

namespace NumDesTools.Tests;

/// <summary>
/// 主题适配 TDD 测试：验证 CTP UserControl XAML 不再使用硬编码颜色，
/// 改用 wpf-ui 主题画刷的 DynamicResource 引用。
///
/// 背景：CTP（如表格目录、表格查询结果）在 ElementHost 里承载 WPF UserControl，
/// 不挂 SystemThemeWatcher。硬编码颜色（Background="Black" 等）不响应
/// ApplicationThemeManager.Apply 触发的全局 MergedDictionaries swap，
/// 导致系统切浅色时 CTP 仍保持深色。改用 DynamicResource 后，
/// 全局 swap Light.xaml <-> Dark.xaml 时自动重解析画刷。
///
/// 此测试用纯字符串扫描验证 XAML 标记层正确性，不加载 WPF/Application，
/// 与项目既有测试风格一致（见 MainWindowBehaviorBaselineTests 的纯逻辑验证约定）。
/// </summary>
public class CtpThemeAdaptationTests
{
    private static readonly string UiDir = Path.Combine(
        AppContext.BaseDirectory,
        "..",
        "..",
        "..",
        "..",
        "NumDesTools",
        "UI"
    );

    private static readonly string[] ThemeAwareXamlFiles =
    [
        "SheetListControl.xaml",
        "CellSeachResult.xaml",
        "SheetCellSeachResult.xaml",
        "SheetSeachResult.xaml",
        "HelpWindow.xaml",
        "GitExportSelectWindow.xaml",
        "ActivityBackupReportWindow.xaml",
        "ActivityBackupSettingsWindow.xaml",
        "BatchReplacePanel.xaml",
        "BatchReplaceWindow.xaml",
        "AIAgentPanel.xaml",
        "AIChatTaskPanel.xaml",
        "ConflictRowItem.xaml",
        "DiffProgressWindow.xaml",
        "ExcelConflictWindow.xaml",
        "ExcelFilePickerWindow.xaml",
        "GitConflictPickerWindow.xaml",
        "GitHistoryPickerWindow.xaml",
        "ImagePreviewControl.xaml",
        "InputBoxDialog.xaml",
        "InputFormularWindow.xaml",
        "LoopRunCheckBoxWindow.xaml",
        "PasswordDialog.xaml",
        "PluginLogWindow.xaml",
        "SheetLinksWindow.xaml",
        "SuperFindAndReplaceWindow.xaml",
        "XlsxSlimmerWindow.xaml",
        "XlsxSyncSettingsWindow.xaml",
    ];

    /// <summary>
    /// 所有主题适配过的 CTP XAML 必须存在（防止路径约定变更后静默跳过测试）。
    /// </summary>
    [Fact]
    public void AllThemeAdaptedXamlFiles_ExistInUiDirectory()
    {
        foreach (var file in ThemeAwareXamlFiles)
        {
            var path = Path.Combine(UiDir, file);
            Assert.True(
                File.Exists(path),
                $"Expected XAML file not found: {path}. "
                    + "If the file moved, update UiDir or ThemeAwareXamlFiles."
            );
        }
    }

    /// <summary>
    /// 主题适配后的 XAML 不允许出现硬编码颜色字符串。
    /// 匹配 #RGB/#RRGGBB/#AARRGGBB 格式（含引号内和属性值）。
    /// 已知的主题画刷 DynamicResource 引用是允许的。
    /// </summary>
    [Theory]
    [MemberData(nameof(ThemeAwareXamlFileNames))]
    public void Xaml_DoesNotContain_HardcodedColors(string xamlFileName)
    {
        var content = ReadXaml(xamlFileName);

        // 匹配 #RGB / #RRGGBB / #AARRGGBB（含可能带 FF 前缀的 8 位形式）
        var hexColorPattern = new Regex(
            "#[0-9A-Fa-f]{3}(?![0-9A-Fa-f])|#[0-9A-Fa-f]{6}(?![0-9A-Fa-f])|#[0-9A-Fa-f]{8}(?![0-9A-Fa-f])"
        );
        var matches = hexColorPattern.Matches(content);
        Assert.Empty(matches);
    }

    /// <summary>
    /// 主题适配后的 XAML 不允许出现命名颜色（Black/White/Gray 等）作为 Background/Foreground。
    /// 这些是硬编码颜色，不响应主题切换。
    /// </summary>
    [Theory]
    [MemberData(nameof(ThemeAwareXamlFileNames))]
    public void Xaml_DoesNotContain_NamedColors_ForBackgroundForeground(string xamlFileName)
    {
        var content = ReadXaml(xamlFileName);

        // 匹配 Background="NamedColor" 或 Foreground="NamedColor" 中 NamedColor 不是 DynamicResource 的情况
        // Transparent 是合理的（让父容器主题色透出），排除
        var namedColorBrushPattern = new Regex(
            @"(Background|Foreground)=""
                (?!    # 后面不能是 DynamicResource/StaticResource/TemplateBinding/Binding
                (?:DynamicResource|StaticResource|TemplateBinding|\{Binding)
                )
                (?!Transparent\b)  # Transparent 允许（透明背景让父主题色透出）
                [A-Z][a-zA-Z]+  # 命名颜色：首字母大写 + 字母
                """,
            RegexOptions.IgnorePatternWhitespace
        );

        var matches = namedColorBrushPattern.Matches(content);
        Assert.Empty(matches);
    }

    /// <summary>
    /// 主题适配后的 XAML 的 ListBox/StatusBar 必须用 DynamicResource 引用主题画刷。
    /// 验证关键画刷 key 存在（ApplicationBackgroundBrush / TextFillColorPrimaryBrush 等）。
    /// </summary>
    [Theory]
    [InlineData("SheetListControl.xaml", "ApplicationBackgroundBrush")]
    [InlineData("SheetListControl.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("CellSeachResult.xaml", "ApplicationBackgroundBrush")]
    [InlineData("CellSeachResult.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("SheetCellSeachResult.xaml", "ApplicationBackgroundBrush")]
    [InlineData("SheetCellSeachResult.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("SheetSeachResult.xaml", "ApplicationBackgroundBrush")]
    [InlineData("SheetSeachResult.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("HelpWindow.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("HelpWindow.xaml", "TextFillColorSecondaryBrush")]
    [InlineData("HelpWindow.xaml", "LayerFillColorDefaultBrush")]
    [InlineData("GitExportSelectWindow.xaml", "TextFillColorPrimaryBrush")]
    [InlineData("GitExportSelectWindow.xaml", "TextFillColorSecondaryBrush")]
    [InlineData("GitExportSelectWindow.xaml", "LayerFillColorDefaultBrush")]
    public void Xaml_Contains_DynamicResourceBrush(string xamlFileName, string brushKey)
    {
        var content = ReadXaml(xamlFileName);
        var expected = $"{{DynamicResource {brushKey}}}";
        Assert.Contains(expected, content);
    }

    /// <summary>
    /// Separator 的 Background 必须用 DynamicResource 引用主题画刷（SeparatorBorderBrush），
    /// 不能是硬编码 Gray。
    /// </summary>
    [Theory]
    [InlineData("CellSeachResult.xaml")]
    [InlineData("SheetCellSeachResult.xaml")]
    public void Xaml_Separator_UsesDynamicResourceBrush(string xamlFileName)
    {
        var content = ReadXaml(xamlFileName);
        Assert.Contains("{DynamicResource SeparatorBorderBrush}", content);
    }

    private static string ReadXaml(string fileName)
    {
        var path = Path.Combine(UiDir, fileName);
        if (!File.Exists(path))
            throw new FileNotFoundException($"XAML file not found: {path}", path);
        return File.ReadAllText(path);
    }

    public static IEnumerable<object[]> ThemeAwareXamlFileNames =>
        ThemeAwareXamlFiles.Select(file => new object[] { file });
}
