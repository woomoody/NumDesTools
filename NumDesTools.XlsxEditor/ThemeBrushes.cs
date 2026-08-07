using System.Windows.Media;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// XlsxEditor 全局主题色刷——所有 brush 都是 mutable 的 SolidColorBrush，
/// 主题切换时调 <see cref="ApplyTheme"/> 更新颜色，所有引用它们的控件自动重绘。
/// </summary>
internal static class ThemeBrushes
{
    // ── 网格线 / 列头 ────────────────────────────────────────────────────
    public static readonly SolidColorBrush GridLine = new(Color.FromRgb(60, 60, 60));
    public static readonly SolidColorBrush HeaderBackground = new(Color.FromRgb(45, 45, 45));

    // ── 文字 ──────────────────────────────────────────────────────────────
    public static readonly SolidColorBrush TextForeground = new(Colors.White);
    public static readonly SolidColorBrush TextSecondary = new(Colors.Gray);
    public static readonly SolidColorBrush TextSelected = new(Colors.White);

    // ── 冻结列分界线 ──────────────────────────────────────────────────────
    public static readonly SolidColorBrush FreezeDivider = new(Color.FromRgb(120, 120, 120));

    // ── Sheet tab 颜色 ────────────────────────────────────────────────────
    public static readonly SolidColorBrush SheetPanelBg = new(Color.FromRgb(30, 30, 30));
    public static readonly SolidColorBrush SheetTabUnselectedBg = new(Color.FromRgb(45, 45, 45));
    public static readonly SolidColorBrush SheetTabSelectedBg = new(Color.FromRgb(80, 80, 80));
    public static readonly SolidColorBrush SheetTabBorder = new(Color.FromRgb(86, 156, 214));
    public static readonly SolidColorBrush SheetBorderSeparator = new(Color.FromRgb(120, 120, 120));

    // ── 工作簿 tab 颜色 ────────────────────────────────────────────────────
    public static readonly SolidColorBrush WbTabUnselectedBg = new(Color.FromRgb(50, 50, 50));
    public static readonly SolidColorBrush WbTabSelectedBg = new(Color.FromRgb(90, 90, 90));

    // ── 筛选框 ────────────────────────────────────────────────────────────
    public static readonly SolidColorBrush FilterBg = new(Color.FromRgb(45, 45, 45));
    public static readonly SolidColorBrush FilterBorder = new(Color.FromRgb(90, 90, 90));

    /// <summary>根据当前主题模式更新所有 brush 颜色。</summary>
    internal static void ApplyTheme(AppThemeMode mode)
    {
        var isDark = mode switch
        {
            AppThemeMode.Light => false,
            AppThemeMode.Dark => true,
            _ => ApplicationThemeIsDark(),
        };

        if (isDark)
            ApplyDark();
        else
            ApplyLight();
    }

    private static bool ApplicationThemeIsDark()
    {
        try
        {
            return Wpf.Ui.Appearance.ApplicationThemeManager.GetAppTheme()
                == Wpf.Ui.Appearance.ApplicationTheme.Dark;
        }
        catch
        {
            return true; // fallback 暗色
        }
    }

    private static void ApplyDark()
    {
        GridLine.Color = Color.FromRgb(60, 60, 60);
        HeaderBackground.Color = Color.FromRgb(45, 45, 45);
        TextForeground.Color = Colors.White;
        TextSecondary.Color = Colors.Gray;
        TextSelected.Color = Colors.White;
        FreezeDivider.Color = Color.FromRgb(120, 120, 120);
        SheetPanelBg.Color = Color.FromRgb(30, 30, 30);
        SheetTabUnselectedBg.Color = Color.FromRgb(45, 45, 45);
        SheetTabSelectedBg.Color = Color.FromRgb(80, 80, 80);
        SheetTabBorder.Color = Color.FromRgb(86, 156, 214);
        SheetBorderSeparator.Color = Color.FromRgb(120, 120, 120);
        WbTabUnselectedBg.Color = Color.FromRgb(50, 50, 50);
        WbTabSelectedBg.Color = Color.FromRgb(90, 90, 90);
        FilterBg.Color = Color.FromRgb(45, 45, 45);
        FilterBorder.Color = Color.FromRgb(90, 90, 90);
    }

    private static void ApplyLight()
    {
        GridLine.Color = Color.FromRgb(200, 200, 200);
        HeaderBackground.Color = Color.FromRgb(230, 230, 230);
        TextForeground.Color = Colors.Black;
        TextSecondary.Color = Color.FromRgb(100, 100, 100);
        TextSelected.Color = Colors.Black;
        FreezeDivider.Color = Color.FromRgb(150, 150, 150);
        SheetPanelBg.Color = Color.FromRgb(248, 248, 248);
        SheetTabUnselectedBg.Color = Color.FromRgb(240, 240, 240);
        SheetTabSelectedBg.Color = Color.FromRgb(210, 210, 210);
        SheetTabBorder.Color = Color.FromRgb(0, 90, 158);
        SheetBorderSeparator.Color = Color.FromRgb(180, 180, 180);
        WbTabUnselectedBg.Color = Color.FromRgb(235, 235, 235);
        WbTabSelectedBg.Color = Color.FromRgb(200, 200, 200);
        FilterBg.Color = Color.FromRgb(240, 240, 240);
        FilterBorder.Color = Color.FromRgb(180, 180, 180);
    }
}