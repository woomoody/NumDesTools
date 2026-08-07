using System.Windows;
using System.Windows.Media;
using Wpf.Ui.Appearance;

namespace NumDesTools.UI;

/// <summary>
/// 主题模式状态机：持久化到 GlobalVariable["ThemeMode"]（System/Light/Dark）。
/// 插件 OnLoad 时由 MahAppsHelper.EnsureInitialized 调用 LoadMode 应用持久化模式；
/// System 模式下系统深浅色联动由 wpfui 的 SystemThemeWatcher.Watch 在窗口级维持。
/// ModeChanged 事件供无法用 DynamicResource 响应的控件（WebView2/AvalonEdit）订阅。
/// </summary>
internal static class ThemeService
{
    internal enum ThemeMode
    {
        System,
        Light,
        Dark,
    }

    private const string ConfigKey = "ThemeMode";

    internal static ThemeMode CurrentMode { get; private set; } = ThemeMode.System;

    /// <summary>主题模式切换事件，供非 DynamicResource 控件订阅刷新。</summary>
    internal static event System.Action? ModeChanged;

    /// <summary>应用持久化的主题模式（插件 OnLoad 时调用一次）。</summary>
    internal static void LoadMode() => Apply(ParseMode(ReadConfig()));

    /// <summary>切换主题模式并持久化到配置文件。</summary>
    internal static void SetMode(ThemeMode mode)
    {
        if (mode == CurrentMode)
            return;
        Apply(mode);
        NumDesAddIn.GlobalValue.SaveValue(ConfigKey, mode.ToString());
    }

    internal static string ModeLabel(ThemeMode mode) =>
        mode switch
        {
            ThemeMode.System => "跟随系统",
            ThemeMode.Light => "亮色",
            ThemeMode.Dark => "暗色",
            _ => "跟随系统",
        };

    private static void Apply(ThemeMode mode)
    {
        CurrentMode = mode;
        switch (mode)
        {
            case ThemeMode.System:
                ApplicationThemeManager.ApplySystemTheme();
                break;
            case ThemeMode.Light:
                ApplicationThemeManager.Apply(ApplicationTheme.Light);
                break;
            case ThemeMode.Dark:
                ApplicationThemeManager.Apply(ApplicationTheme.Dark);
                break;
        }
        ForceRefreshDynamicResources();
        ModeChanged?.Invoke();
        LogResolvedThemeState();
    }

    /// <summary>输出主题切换后字典内关键资源实际值，用于定位 ElementHost 树不刷新问题。</summary>
    private static void LogResolvedThemeState()
    {
        var app = System.Windows.Application.Current;
        if (app is null)
            return;
        var bg = app.Resources["ApplicationBackgroundBrush"] as SolidColorBrush;
        var fg = app.Resources["TextFillColorPrimaryBrush"] as SolidColorBrush;
        PluginLog.Verbose(
            $"[Theme] Apply done: appTheme={ApplicationThemeManager.GetAppTheme()} "
                + $"AppBg={bg?.Color} TextPrimary={fg?.Color}"
        );
    }

    /// <summary>
    /// 兜底：把 wpf-ui 主题字典从合并字典中移除再插回，广播资源变更，
    /// 强制已加载元素（含 ElementHost 里的 CTP 控件）的 DynamicResource 重新求值。
    /// 正常情况下 Apply 已广播变更，此操作只是多一次无害广播。
    /// </summary>
    private static void ForceRefreshDynamicResources()
    {
        var app = System.Windows.Application.Current;
        if (app is null)
            return;
        var dicts = app.Resources.MergedDictionaries;
        for (var i = 0; i < dicts.Count; i++)
        {
            if (dicts[i] is not Wpf.Ui.Markup.ThemesDictionary)
                continue;
            var themeDict = dicts[i];
            dicts.RemoveAt(i);
            dicts.Insert(i, themeDict);
            break;
        }
    }

    private static string ReadConfig() =>
        NumDesAddIn.GlobalValue.Value.TryGetValue(ConfigKey, out var v) ? v : nameof(ThemeMode.System);

    private static ThemeMode ParseMode(string value) =>
        Enum.TryParse<ThemeMode>(value, ignoreCase: true, out var mode) ? mode : ThemeMode.System;
}
