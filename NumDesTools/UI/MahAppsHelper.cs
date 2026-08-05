using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace NumDesTools.UI;

internal static class MahAppsHelper
{
    private static bool _initialized;

    internal static void EnsureInitialized()
    {
        if (_initialized)
            return;
        _initialized = true;

        if (System.Windows.Application.Current is null)
            _ = new System.Windows.Application
            {
                ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown,
            };

        var app = System.Windows.Application.Current;

        // wpfui 全局合并
        if (
            !app.Resources.MergedDictionaries.Any(d =>
                d.Source?.OriginalString.Contains("Wpf.Ui") == true
            )
        )
        {
            app.Resources.MergedDictionaries.Add(
                new Wpf.Ui.Markup.ThemesDictionary
                {
                    Theme = Wpf.Ui.Appearance.ApplicationTheme.Dark,
                }
            );
            app.Resources.MergedDictionaries.Add(new Wpf.Ui.Markup.ControlsDictionary());
        }

        // 应用持久化的主题模式（System 走 ApplySystemTheme，亮/暗走 Apply）
        ThemeService.LoadMode();

        // 加载自定义语义色资源字典（NumDesTools.ThemeDictionaries.xaml）。
        // 该字典在 XAML 中定义了 13 个 SolidColorBrush，通过 DynamicResource
        // 引用 Color 值。主题切换时只需更新 Color 值，Brush 自动更新。
        var themeDictUri = new Uri(
            "pack://application:,,,/NumDesTools;component/UI/NumDesTools.ThemeDictionaries.xaml",
            UriKind.Absolute
        );
        if (!app.Resources.MergedDictionaries.Any(d => d.Source == themeDictUri))
        {
            app.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = themeDictUri });
        }

        // 初始注入语义色值（Color 级别，不是 Brush 级别）。
        // 订阅 Changed 事件：主题切换时更新 Color 值，XAML 中定义的 Brush 通过
        // DynamicResource 自动响应变化，无需重建 Brush 对象。
        ApplySemanticColorValues();
        Wpf.Ui.Appearance.ApplicationThemeManager.Changed += (_, _) =>
            ApplySemanticColorValues();

        PluginLog.Verbose(
            $"[Theme] Init: appTheme={Wpf.Ui.Appearance.ApplicationThemeManager.GetAppTheme()} "
                + $"systemTheme={Wpf.Ui.Appearance.ApplicationThemeManager.GetSystemTheme()}"
        );
    }

    /// <summary>
    /// 更新语义色值（Color 资源，不是 Brush）。
    /// 画刷在 XAML 中定义并通过 DynamicResource 引用这些色值，
    /// 主题切换时只需更新 Color 值，所有引用自动更新。
    /// </summary>
    private static void ApplySemanticColorValues()
    {
        var app = System.Windows.Application.Current;
        if (app is null)
            return;

        var isDark = Wpf.Ui.Appearance.ApplicationThemeManager.GetAppTheme()
            == Wpf.Ui.Appearance.ApplicationTheme.Dark;

        SetColor("OursTextColor", isDark, "#FFAAAA", "#B33A3A");
        SetColor("OursBgColor", isDark, "#3A2A2A", "#FFDDDD");
        SetColor("OursActionBgColor", isDark, "#5A2A2A", "#FFCCCC");
        SetColor("TheirsTextColor", isDark, "#A8FFCA", "#2A8A2A");
        SetColor("TheirsBgColor", isDark, "#1A5C3A", "#DDFFDD");
        SetColor("ConflictTextColor", isDark, "#AAAAFF", "#5555AA");
        SetColor("ConflictBgColor", isDark, "#3A3A5A", "#EEDDFF");
        SetColor("HistoryTextColor", isDark, "#88CCFF", "#1A6BB8");
        SetColor("HistoryBgColor", isDark, "#1A3A6E", "#D6E8FF");
        SetColor("AiSuggestionTextColor", isDark, "#FFD080", "#B87400");
        SetColor("AiSuggestionBgColor", isDark, "#2A2A1A", "#FFF8E0");
        SetColor("AiSuggestionBorderColor", isDark, "#554400", "#DDAA55");
        app.Resources["SemanticButtonTextColor"] = Colors.Black;
    }

    private static void SetColor(string key, bool isDark, string dark, string light)
    {
        var app = System.Windows.Application.Current;
        app.Resources[key] = (System.Windows.Media.Color)
            System.Windows.Media.ColorConverter.ConvertFromString(isDark ? dark : light);
    }

    internal static void SetExcelOwner(System.Windows.Window window)
    {
        var hwnd = (IntPtr)ExcelDnaUtil.WindowHandle;
        if (hwnd != IntPtr.Zero)
            new WindowInteropHelper(window).Owner = hwnd;
        window.Loaded += (_, _) =>
        {
            Wpf.Ui.Appearance.SystemThemeWatcher.Watch(window);
            AttachTitleBarDrag(window);
        };
    }

    private static void AttachTitleBarDrag(System.Windows.Window window)
    {
        if (
            window.Template?.FindName("PART_TitleBar", window)
            is not System.Windows.UIElement titleBar
        )
            return;

        var hwnd = new WindowInteropHelper(window).Handle;
        var hwndSource = HwndSource.FromHwnd(hwnd);
        if (hwndSource is null)
            return;

        int dragStartX = 0,
            dragStartY = 0,
            winX = 0,
            winY = 0;
        bool dragging = false;

        titleBar.PreviewMouseLeftButtonDown += (_, e) =>
        {
            if (window.WindowState != System.Windows.WindowState.Normal)
                return;
            GetCursorPos(out var pt);
            dragStartX = pt.X;
            dragStartY = pt.Y;
            GetWindowRect(hwnd, out var r);
            winX = r.Left;
            winY = r.Top;
            dragging = true;
            SetCapture(hwnd);
            e.Handled = true;
        };

        hwndSource.AddHook(
            (IntPtr h, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) =>
            {
                const int WmMousemove = 0x0200;
                const int WmLbuttonup = 0x0202;
                const int WmCapturechanged = 0x0215;

                if (msg == WmMousemove && dragging)
                {
                    GetCursorPos(out var pt);
                    SetWindowPos(
                        hwnd,
                        IntPtr.Zero,
                        winX + (pt.X - dragStartX),
                        winY + (pt.Y - dragStartY),
                        0,
                        0,
                        SwpNosize | SwpNozorder | SwpNoactivate
                    );
                    handled = true;
                }
                else if ((msg == WmLbuttonup || msg == WmCapturechanged) && dragging)
                {
                    dragging = false;
                    ReleaseCapture();
                    handled = msg == WmLbuttonup;
                }
                return IntPtr.Zero;
            }
        );
    }

    internal static void ApplyDarkTitleBar(System.Windows.Window window)
    {
        window.Loaded += (_, _) =>
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero)
                return;
            int dark = 1;
            DwmSetWindowAttribute(hwnd, 20, ref dark, sizeof(int));
        };
    }

    private const uint SwpNosize = 0x0001;
    private const uint SwpNozorder = 0x0004;
    private const uint SwpNoactivate = 0x0010;

    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        public int Left,
            Top,
            Right,
            Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Point
    {
        public int X,
            Y;
    }

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out Point pt);

    [DllImport("user32.dll")]
    private static extern IntPtr SetCapture(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr hwnd,
        IntPtr hwndAfter,
        int x,
        int y,
        int cx,
        int cy,
        uint flags
    );

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);

    [DllImport("user32.dll")]
    internal static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    internal static extern IntPtr SetFocus(IntPtr hWnd);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}