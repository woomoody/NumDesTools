using System.IO;
using System.Text.Json;
using Wpf.Ui.Appearance;

namespace NumDesTools.XlsxEditor;

internal enum ThemeMode
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
    internal static ThemeMode CurrentMode { get; private set; } = ThemeMode.System;

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

    internal static void SetMode(ThemeMode mode)
    {
        if (mode == CurrentMode)
            return;
        CurrentMode = mode;
        SaveConfig(mode);
        ApplyTheme(mode);
    }

    internal static string ModeLabel(ThemeMode mode) =>
        mode switch
        {
            ThemeMode.System => "跟随系统",
            ThemeMode.Light => "亮色",
            ThemeMode.Dark => "暗色",
            _ => "跟随系统",
        };

    private static void ApplyTheme(ThemeMode mode)
    {
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
        ModeChanged?.Invoke();
    }

    private static ThemeMode ReadConfig()
    {
        try
        {
            if (!File.Exists(ConfigPath))
                return ThemeMode.System;
            var json = File.ReadAllText(ConfigPath);
            var data = JsonSerializer.Deserialize<ThemeConfig>(json);
            return data?.Mode switch
            {
                "Light" => ThemeMode.Light,
                "Dark" => ThemeMode.Dark,
                _ => ThemeMode.System,
            };
        }
        catch
        {
            return ThemeMode.System;
        }
    }

    private static void SaveConfig(ThemeMode mode)
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