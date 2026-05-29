using System.Runtime.InteropServices;
using System.Windows.Input;
using System.Windows.Interop;
using ControlzEx.Theming;
using MahApps.Metro.Controls;
using SWWindow = System.Windows.Window;

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

        // 无 App.xaml 时必须手动合并 MahApps 核心资源
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

    internal static void SetExcelOwner(SWWindow window)
    {
        var excelHwnd = (IntPtr)ExcelDnaUtil.WindowHandle;

        window.SourceInitialized += (_, _) =>
        {
            // ShowDialog 调用后 WPF 的 set_OwnerHandle 会抛异常（_showingAsDialog 守卫），
            // 用 SetWindowLong(GWL_HWNDPARENT) 直接写 Win32 父窗口关系，兼容 Show/ShowDialog。
            SetWindowLong(new WindowInteropHelper(window).Handle, GwlHwndparent, excelHwnd.ToInt32());
        };

        // EnsureHandle 预建 HWND，使 SourceInitialized 在 ShowDialog 进入前同步触发，
        // 此时 _showingAsDialog 尚为 false，SetWindowLong 调用安全。
        new WindowInteropHelper(window).EnsureHandle();

        // 纯 WPF 拖动：所有 Win32 SC_MOVE/DragMove 方案在 Excel 宿主消息泵下均失效，
        // 在 Loaded 后找 PART_TitleBar，用 MouseMove 直接更新 Left/Top 绕开 Win32 modal loop。
        window.Loaded += (_, _) => AttachTitleBarDrag(window);
    }

    private static void AttachTitleBarDrag(SWWindow window)
    {
        if (window.Template?.FindName("PART_TitleBar", window) is not System.Windows.UIElement titleBar)
            return;

        System.Windows.Point dragStart = default;
        double winLeft = 0, winTop = 0;
        bool dragging = false;

        titleBar.MouseLeftButtonDown += (_, e) =>
        {
            if (window.WindowState != System.Windows.WindowState.Normal)
                return;
            dragging = true;
            dragStart = e.GetPosition(null);
            winLeft = window.Left;
            winTop = window.Top;
            titleBar.CaptureMouse();
            e.Handled = true;
        };

        titleBar.MouseMove += (_, e) =>
        {
            if (!dragging || e.LeftButton != MouseButtonState.Pressed)
            {
                if (dragging)
                {
                    dragging = false;
                    titleBar.ReleaseMouseCapture();
                }
                return;
            }
            var cur = e.GetPosition(null);
            window.Left = winLeft + (cur.X - dragStart.X);
            window.Top = winTop + (cur.Y - dragStart.Y);
        };

        titleBar.MouseLeftButtonUp += (_, _) =>
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
            var hwnd = new WindowInteropHelper(window).Handle;
            if (hwnd == IntPtr.Zero)
                return;
            int dark = 1;
            DwmSetWindowAttribute(hwnd, 20, ref dark, sizeof(int));
        };
    }

    private const int GwlHwndparent = -8;

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hwnd, int nIndex, int dwNewLong);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}
