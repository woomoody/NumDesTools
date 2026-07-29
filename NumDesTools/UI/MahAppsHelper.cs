using System.Runtime.InteropServices;
using System.Windows.Interop;
using ControlzEx.Theming;
using MahApps.Metro.Controls;

namespace NumDesTools.UI;

internal static class MahAppsHelper
{
    /// <summary>
    /// 判断某个 URI 片段的资源字典是否已存在于 MergedDictionaries 中——
    /// 抽出为纯函数，不依赖真实 Application/STA 线程，可单测覆盖
    /// （见 NumDesOutput/analysis/wpfui-migration-optimization-report-2026-07-29.md P1）。
    /// 接受 <see cref="Uri"/> 序列而非 <see cref="System.Windows.ResourceDictionary"/>——
    /// ResourceDictionary.Source 的 setter 会触发真实资源加载（网络/pack 解析），
    /// 单测里不应该构造会触发 IO 的对象，调用方在 EnsureInitialized 里传
    /// mergedDictionaries.Select(d => d.Source) 即可复用同一份逻辑。
    /// </summary>
    /// <param name="exactMatch">
    /// true：URI 的完整字符串必须与 <paramref name="uriOrFragment"/> 相等（MahApps 资源用此模式）。
    /// false（默认）：URI 字符串只需包含 <paramref name="uriOrFragment"/> 子串（wpfui 资源用此模式）。
    /// </param>
    internal static bool IsResourceMerged(
        IEnumerable<Uri?> mergedSources,
        string uriOrFragment,
        bool exactMatch = false
    )
    {
        return mergedSources.Any(source =>
            exactMatch
                ? source?.OriginalString == uriOrFragment
                : source?.OriginalString.Contains(uriOrFragment) == true
        );
    }

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
                !IsResourceMerged(
                    app.Resources.MergedDictionaries.Select(d => d.Source),
                    u,
                    exactMatch: true
                )
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
        if (!IsResourceMerged(app.Resources.MergedDictionaries.Select(d => d.Source), "Wpf.Ui"))
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
                    "wpfui Dark + Controls merged OK (global, CTP overrides applied per-Control)"
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

        // 主题跟随系统（不强制固定 Dark）——
        // ApplySystemTheme 根据系统当前浅色/深色应用主题，系统切换时自动响应。
        // 之前用 Apply(Dark) 强制 Dark 会压住系统主题跟随，导致 CTP 背景不随系统走。
        try
        {
            Wpf.Ui.Appearance.ApplicationThemeManager.ApplySystemTheme();
            WriteDebugLog("ApplicationThemeManager.ApplySystemTheme called (follow system)");
        }
        catch (Exception ex)
        {
            WriteDebugLog($"ApplicationThemeManager.ApplySystemTheme FAILED: {ex}");
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
