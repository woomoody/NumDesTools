using System.IO;
using System.Text.Json;
using Wpf.Ui.Appearance;

namespace NumDesTools.XlsxEditor;

// 注意：不能叫 ThemeMode——WPF 内置的 Window.ThemeMode 属性会冲突
internal enum AppThemeMode
{
    System,
    Light,
    Dark,
}

/// <summary>
/// XlsxEditor 独立主题切换服务，不依赖 NumDesTools 的 GlobalVariable。
/// 持久化到 %LOCALAPPDATA%/NumDesTools/XlsxEditor/theme.json。
/// </summary>
internal static class ThemeService
{
    internal static AppThemeMode CurrentMode { get; private set; } = AppThemeMode.System;

    internal static event Action? ModeChanged;

    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "NumDesTools",
        "XlsxEditor"
    );

    private static readonly string ConfigPath = Path.Combine(ConfigDir, "theme.json");

    internal static void LoadMode()
    {
        var mode = ReadConfig();
        CurrentMode = mode;
        ApplyTheme(mode);
    }

    internal static void SetMode(AppThemeMode mode)
    {
        if (mode == CurrentMode)
            return;
        CurrentMode = mode;
        SaveConfig(mode);
        ApplyTheme(mode);
    }

    internal static string ModeLabel(AppThemeMode mode) =>
        mode switch
        {
            AppThemeMode.System => "跟随系统",
            AppThemeMode.Light => "亮色",
            AppThemeMode.Dark => "暗色",
            _ => "跟随系统",
        };

    private static void ApplyTheme(AppThemeMode mode)
    {
        switch (mode)
        {
            case AppThemeMode.System:
                ApplicationThemeManager.ApplySystemTheme();
                break;
            case AppThemeMode.Light:
                ApplicationThemeManager.Apply(ApplicationTheme.Light);
                break;
            case AppThemeMode.Dark:
                ApplicationThemeManager.Apply(ApplicationTheme.Dark);
                break;
        }
        ModeChanged?.Invoke();
    }

    private static AppThemeMode ReadConfig()
    {
        try
        {
            if (!File.Exists(ConfigPath))
                return AppThemeMode.System;
            var json = File.ReadAllText(ConfigPath);
            var data = JsonSerializer.Deserialize<ThemeConfig>(json);
            return data?.Mode switch
            {
                "Light" => AppThemeMode.Light,
                "Dark" => AppThemeMode.Dark,
                _ => AppThemeMode.System,
            };
        }
        catch
        {
            return AppThemeMode.System;
        }
    }

    private static void SaveConfig(AppThemeMode mode)
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
            var data = new ThemeConfig { Mode = mode.ToString() };
            var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
        }
        catch
        {
            // 持久化失败不影响使用
        }
    }

    private sealed class ThemeConfig
    {
        public string Mode { get; set; } = "System";
    }
}