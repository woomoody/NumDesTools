using System.Runtime.InteropServices;
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

        // 手动 merge MahApps 核心资源（无 App.xaml 时必须）
        var app = System.Windows.Application.Current;

        // wpfui 全局合并（替代 MahApps）
        // 用 Theme setter 而非直接设 Source：ThemesDictionary 内部据此维护
        // IsSourcedFromThemeDictionary 状态，ApplicationThemeManager.UpdateDictionary
        // swap Light.xaml <-> Dark.xaml 时才能正确识别并替换这个字典槽位。
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

        // 自适应系统主题：读取 Windows 当前亮/暗状态并应用。
        // SystemThemeWatcher.Watch(window) 在每个窗口 Loaded 里挂上，
        // 之后系统主题切换时 wpf-ui 自动 swap Light.xaml <-> Dark.xaml。
        Wpf.Ui.Appearance.ApplicationThemeManager.ApplySystemTheme();

        // 全局注入语义画刷（Ours/Theirs/Conflict/History/AiSuggestion）。
        // 这些画刷在 ConflictRowItem.xaml 和 ExcelConflictWindow.xaml 里被 DynamicResource 引用，
        // 必须放在 app.Resources 才能被所有窗口/控件找到。
        // 订阅 Changed 事件：主题切换时重新注入（深浅色各一套值）。
        ApplySemanticBrushes();
        Wpf.Ui.Appearance.ApplicationThemeManager.Changed += (_, _) =>
            ApplySemanticBrushes();

        // 诊断日志：确认主题切换是否成功
        PluginLog.Verbose(
            $"[Theme] Init: appTheme={Wpf.Ui.Appearance.ApplicationThemeManager.GetAppTheme()} "
                + $"systemTheme={Wpf.Ui.Appearance.ApplicationThemeManager.GetSystemTheme()}"
        );
    }

    /// <summary>
    /// 全局注入语义画刷到 app.Resources。
    /// 深色/浅色各一套值，由当前 ApplicationTheme 决定。
    /// </summary>
    private static void ApplySemanticBrushes()
    {
        var app = System.Windows.Application.Current;
        if (app is null)
            return;

        var isDark = Wpf.Ui.Appearance.ApplicationThemeManager.GetAppTheme()
            == Wpf.Ui.Appearance.ApplicationTheme.Dark;

        app.Resources["OursTextBrush"] = SemanticBrush(
            isDark,
            "#FFAAAA",
            "#B33A3A"
        );
        app.Resources["OursBackgroundBrush"] = SemanticBrush(
            isDark,
            "#3A2A2A",
            "#FFDDDD"
        );
        app.Resources["OursActionBackgroundBrush"] = SemanticBrush(
            isDark,
            "#5A2A2A",
            "#FFCCCC"
        );
        app.Resources["TheirsTextBrush"] = SemanticBrush(
            isDark,
            "#A8FFCA",
            "#2A8A2A"
        );
        app.Resources["TheirsBackgroundBrush"] = SemanticBrush(
            isDark,
            "#1A5C3A",
            "#DDFFDD"
        );
        app.Resources["ConflictTextBrush"] = SemanticBrush(
            isDark,
            "#AAAAFF",
            "#5555AA"
        );
        app.Resources["ConflictBackgroundBrush"] = SemanticBrush(
            isDark,
            "#3A3A5A",
            "#EEDDFF"
        );
        app.Resources["HistoryTextBrush"] = SemanticBrush(
            isDark,
            "#88CCFF",
            "#1A6BB8"
        );
        app.Resources["HistoryBackgroundBrush"] = SemanticBrush(
            isDark,
            "#1A3A6E",
            "#D6E8FF"
        );
        app.Resources["AiSuggestionTextBrush"] = SemanticBrush(
            isDark,
            "#FFD080",
            "#B87400"
        );
        app.Resources["AiSuggestionBackgroundBrush"] = SemanticBrush(
            isDark,
            "#2A2A1A",
            "#FFF8E0"
        );
        app.Resources["AiSuggestionBorderBrush"] = SemanticBrush(
            isDark,
            "#554400",
            "#DDAA55"
        );
        // 语义色按钮文字：固定黑色。语义背景深浅已自动切换（深色主题=深色背景，浅色主题=浅色背景），
        // 黑字在两种背景上都能看清（深色背景如 #5A2A2A 足够深，浅色背景如 #FFCCCC 足够浅）。
        app.Resources["SemanticButtonTextBrush"] = new SolidColorBrush(Colors.Black);
    }

    private static SolidColorBrush SemanticBrush(bool isDark, string dark, string light) =>
        new(
            (System.Windows.Media.Color)
                System.Windows.Media.ColorConverter.ConvertFromString(isDark ? dark : light)
        );

    internal static void SetExcelOwner(System.Windows.Window window)
    {
        var hwnd = (IntPtr)ExcelDnaUtil.WindowHandle;
        if (hwnd != IntPtr.Zero)
            new WindowInteropHelper(window).Owner = hwnd;
        window.Loaded += (_, _) =>
        {
            // SystemThemeWatcher 挂 Win32 消息钩子（WM_THEMECHANGED 等），
            // 系统亮/暗切换时自动 Apply 新主题到这个窗口 + 全局资源字典。
            Wpf.Ui.Appearance.SystemThemeWatcher.Watch(window);
            AttachTitleBarDrag(window);
        };
    }

    private static void AttachTitleBarDrag(System.Windows.Window window)
    {
        // PART_TitleBar 是 MetroThumbContentControl，会吞 MouseLeftButtonDown，用 Preview 截获。
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

        // PreviewMouseLeftButtonDown 用于记录起点并用 Win32 SetCapture 接管后续消息。
        // 之后的 WM_MOUSEMOVE / WM_LBUTTONUP 直接从 HwndSourceHook 处理，
        // 完全绕开 WPF 输入管道，与原生拖动同等流畅。
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

    /// <summary>
    /// 非 modals WPF 窗口在 Excel 进程内，Excel 的消息循环拦截 WM_KEYDOWN。
    /// 用 SetForegroundWindow + SetFocus 强制把键盘焦点拉回 WPF 窗口。
    /// </summary>
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
