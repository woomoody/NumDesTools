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

        // 手动 merge 资源（无 App.xaml 时必须）
        // 注意：迁移到 wpfui FluentWindow 后，不再需要 MahApps 主题——
        // MahApps Dark.Steel 主题会和 wpfui Dark 冲突，导致 wpfui 控件文字黑底黑字。
        // 只合并 wpfui 资源。MahApps 字典仅在个别遗留控件（如 AIAgentPanel 的 mah:NumericUpDown）需要，
        // 但那些控件不在独立窗口里，不影响。
        var app = System.Windows.Application.Current;

        // wpfui 深色主题资源字典（wpfui 控件如 ui:Button/ui:TextBlock 渲染依赖此字典）
        // 无 App.xaml 的插件宿主必须手动合并，否则深色 Mica 背景上文字不可见（黑框问题）
        // 用 ThemesDictionary + ControlsDictionary 标记类（和 FileLockPreview 验证通过的方式一致）
        if (
            !app.Resources.MergedDictionaries.Any(d =>
                d.Source?.OriginalString.Contains("Wpf.Ui") == true
            )
        )
        {
            try
            {
                app.Resources.MergedDictionaries.Add(
                    new Wpf.Ui.Markup.ThemesDictionary
                    {
                        Source = new Uri(
                            "pack://application:,,,/Wpf.Ui;component/Resources/Theme/Dark.xaml"
                        ),
                    }
                );
                app.Resources.MergedDictionaries.Add(new Wpf.Ui.Markup.ControlsDictionary());
                WriteDebugLog(
                    "wpfui Dark + Controls resources merged OK (ThemesDictionary, no MahApps)"
                );
            }
            catch (Exception ex)
            {
                WriteDebugLog($"wpfui resource merge FAILED: {ex}");
            }
        }
        else
        {
            WriteDebugLog("wpfui resources already merged, skip");
        }
    }

    private static void WriteDebugLog(string msg)
    {
        try
        {
            var path = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "workspace",
                "xlsx-editor-wpfui-init.log"
            );
            System.IO.File.AppendAllText(path, $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\n");
        }
        catch { }
    }

    internal static void SetExcelOwner(System.Windows.Window window)
    {
        var hwnd = (IntPtr)ExcelDnaUtil.WindowHandle;
        if (hwnd != IntPtr.Zero)
            new WindowInteropHelper(window).Owner = hwnd;
        window.Loaded += (_, _) => AttachTitleBarDrag(window);
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
