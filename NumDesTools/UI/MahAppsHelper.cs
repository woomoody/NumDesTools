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

    internal static void SetExcelOwner(System.Windows.Window window)
    {
        var excelHwnd = (IntPtr)ExcelDnaUtil.WindowHandle;

        window.SourceInitialized += (_, _) =>
        {
            var helper = new WindowInteropHelper(window);
            var hwnd = helper.Handle;

            // WPF 的 set_OwnerHandle 在 ShowDialog 启动后会抛 InvalidOperationException（_showingAsDialog 守卫）。
            // 用 Win32 SetWindowLong(GWL_HWNDPARENT) 直接写父窗口，绕过 WPF 托管层检查，
            // 同时兼容 Show() 和 ShowDialog() 两种打开方式。
            SetWindowLong(hwnd, GwlHwndparent, excelHwnd.ToInt32());

            // ReleaseCapture + PostMessage(SC_MOVE) 是非阻塞拖动：
            // PostMessage 把消息投入队列后立即返回，不进入 Win32 modal loop，
            // Excel 宿主消息泵不会被阻塞。
            // 不走 WPF DragMove()：DragMove 在 HwndSourceHook 时序下 Mouse.LeftButton 尚未变 Pressed，
            // 内部检查会静默失败。
            var hwndSource = HwndSource.FromHwnd(hwnd);
            hwndSource?.AddHook(
                (IntPtr h, int msg, IntPtr wParam, IntPtr lParam, ref bool handled) =>
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
                        ReleaseCapture();
                        PostMessage(h, WmSyscommand, new IntPtr(ScMove | HtCaption), IntPtr.Zero);
                    }
                    return IntPtr.Zero;
                }
            );
        };

        // EnsureHandle() 预建 HWND 并同步触发上方的 SourceInitialized 回调，
        // 必须在 Show/ShowDialog 之前调用，否则 SourceInitialized 在 ShowDialog 内触发时
        // _showingAsDialog 已为 true，此时再设父窗口关系会受 WPF 守卫阻断。
        new WindowInteropHelper(window).EnsureHandle();
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
    private const int WmSyscommand = 0x0112;
    private const int ScMove = 0xF010;

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hwnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}
