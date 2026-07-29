using System.Runtime.InteropServices;
using System.Windows.Interop;
using ControlzEx.Theming;
using MahApps.Metro.Controls;

namespace NumDesTools.UI;

internal static class MahAppsHelper
{
    internal static void EnsureInitialized()
    {
        if (System.Windows.Application.Current is null)
            _ = new System.Windows.Application
            {
                ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown,
            };

        // 手动 merge 资源（无 App.xaml 时必须）
        var app = System.Windows.Application.Current;

        // MahApps 核心资源——CTP（CustomTaskPane）里的 mah:NumericUpDown/mah:TextBoxHelper.Watermark 依赖这些。
        // 去掉后 CTP 控件没样式 fallback 默认（黑块 + 大字体大间距）。
        // FluentWindow 不受影响（wpfui 字典后合并覆盖 + Apply Dark 每次调用）。
        var mahappsUris = new[]
        {
            "pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml",
            "pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml",
            "pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml",
        };
        foreach (
            var uri in mahappsUris.Where(u =>
                !app.Resources.MergedDictionaries.Any(d => d.Source?.OriginalString == u)
            )
        )
        {
            try
            {
                app.Resources.MergedDictionaries.Add(
                    new System.Windows.ResourceDictionary { Source = new Uri(uri) }
                );
            }
            catch (Exception ex)
            {
                WriteDebugLog($"MahApps merge FAILED ({uri}): {ex}");
            }
        }
        try
        {
            ControlzEx.Theming.ThemeManager.Current.ChangeTheme(app, "Dark.Steel");
        }
        catch (Exception ex)
        {
            WriteDebugLog($"MahApps ThemeManager FAILED: {ex}");
        }

        // wpfui 深色主题资源字典（wpfui 控件如 ui:Button/ui:TextBlock 渲染依赖此字典）
        // 在 MahApps 之后合并——后合并的覆盖前面的，FluentWindow 的 ui: 控件用 wpfui 深色样式。
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

        // 每次调用都强制 Apply 深色主题——
        // 第一次调用时 Application 刚建，wpfui 主题可能未真正生效（第一次窗口黑，第二次好的根因）。
        // 每次窗口构造时重新 Apply，确保主题真正应用到当前 Dispatcher 状态。
        try
        {
            Wpf.Ui.Appearance.ApplicationThemeManager.Apply(
                Wpf.Ui.Appearance.ApplicationTheme.Dark
            );
            // 强制同步处理 Background 优先级的 Dispatcher 操作——
            // wpfui Apply 内部可能用 Dispatcher 异步回调应用主题，ShowDialog 前不刷新则第一次窗口黑。
            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(
                System.Windows.Threading.DispatcherPriority.Background,
                new System.Action(() => { })
            );
            WriteDebugLog("ApplicationThemeManager.Apply(Dark) + Dispatcher flush called");
        }
        catch (Exception ex)
        {
            WriteDebugLog($"ApplicationThemeManager.Apply FAILED: {ex}");
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
