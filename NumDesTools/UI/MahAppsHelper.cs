using System.Runtime.InteropServices;
using System.Windows.Interop;
using ControlzEx.Theming;
using MahApps.Metro.Controls;

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
        foreach (
            var uri in new[]
            {
                "pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml",
                "pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml",
                "pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml",
            }
        )
        {
            var rd = new System.Windows.ResourceDictionary { Source = new Uri(uri) };
            app.Resources.MergedDictionaries.Add(rd);
        }

        ThemeManager.Current.ChangeTheme(app, "Dark.Steel");
    }

    internal static void SetExcelOwner(System.Windows.Window window)
    {
        new WindowInteropHelper(window).Owner = (IntPtr)ExcelDnaUtil.WindowHandle;
        window.Loaded += (_, _) => AttachTitleBarDrag(window);
    }

    private static void AttachTitleBarDrag(System.Windows.Window window)
    {
        // PART_TitleBar 实际类型是 MetroThumbContentControl，它会吃掉 MouseLeftButtonDown，
        // 必须用 PreviewMouseLeftButtonDown 才能在 Thumb 处理前拿到事件。
        if (
            window.Template?.FindName("PART_TitleBar", window)
            is not System.Windows.UIElement titleBar
        )
            return;

        var hwnd = new WindowInteropHelper(window).Handle;
        System.Windows.Point dragStartScreen = default;
        int winX = 0,
            winY = 0;
        bool dragging = false;

        titleBar.PreviewMouseLeftButtonDown += (_, e) =>
        {
            if (window.WindowState != System.Windows.WindowState.Normal)
                return;
            dragStartScreen = titleBar.PointToScreen(e.GetPosition(titleBar));
            GetWindowRect(hwnd, out var r);
            winX = r.Left;
            winY = r.Top;
            dragging = true;
            titleBar.CaptureMouse();
        };

        titleBar.PreviewMouseMove += (_, e) =>
        {
            if (!dragging || e.LeftButton != System.Windows.Input.MouseButtonState.Pressed)
            {
                if (dragging)
                {
                    dragging = false;
                    titleBar.ReleaseMouseCapture();
                }
                return;
            }
            var cur = titleBar.PointToScreen(e.GetPosition(titleBar));
            var dx = (int)(cur.X - dragStartScreen.X);
            var dy = (int)(cur.Y - dragStartScreen.Y);
            // SetWindowPos 直接移动 Win32 窗口，不触发 WPF 布局重排，流畅度与原生拖动一致
            SetWindowPos(
                hwnd,
                IntPtr.Zero,
                winX + dx,
                winY + dy,
                0,
                0,
                SwpNosize | SwpNozorder | SwpNoactivate
            );
        };

        titleBar.PreviewMouseLeftButtonUp += (_, _) =>
        {
            dragging = false;
            titleBar.ReleaseMouseCapture();
        };

        titleBar.LostMouseCapture += (_, _) => dragging = false;
    }

    internal static void ApplyDarkTitleBar(MetroWindow window)
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

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}
