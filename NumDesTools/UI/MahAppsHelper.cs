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
        var excelHwnd = (IntPtr)ExcelDnaUtil.WindowHandle;
        window.SourceInitialized += (_, _) =>
        {
            var helper = new WindowInteropHelper(window);
            helper.Owner = excelHwnd;

            // MahApps 标题栏拖动走 Win32 SC_MOVE 同步 modal loop，Excel 消息泵阻塞时会卡死。
            // 改为拦截 WM_NCLBUTTONDOWN(HTCAPTION) 后调 WPF DragMove()，走 WPF InputManager，
            // 不依赖 Win32 modal loop，能在 Excel 宿主的消息泵下正常工作。
            var hwndSource = HwndSource.FromHwnd(helper.Handle);
            hwndSource?.AddHook(
                (IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) =>
                {
                    const int WmNclbuttondown = 0x00A1;
                    const int HtCaption = 2;
                    if (
                        msg == WmNclbuttondown
                        && wParam.ToInt32() == HtCaption
                        && window.WindowState == System.Windows.WindowState.Normal
                    )
                    {
                        handled = true;
                        window.DragMove();
                    }
                    return IntPtr.Zero;
                }
            );
        };
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

    [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}
