using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

internal static class WpfUiHelper
{
    private static bool _themeLoaded;

    /// <summary>
    /// 在 InitializeComponent() 之前调用，把 WPF-UI 深色主题资源 merge 到窗口/控件。
    /// 首次调用后缓存已加载标志，避免重复 IO。
    /// </summary>
    internal static void ApplyDarkTheme(FrameworkElement target)
    {
        EnsureApplication();

        foreach (
            var uri in new[]
            {
                "pack://application:,,,/Wpf.Ui;component/Resources/Wpf.Ui.xaml",
                "pack://application:,,,/Wpf.Ui;component/Resources/Theme/Dark.xaml",
            }
        )
        {
            try
            {
                var rd = new ResourceDictionary { Source = new Uri(uri) };
                target.Resources.MergedDictionaries.Add(rd);
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[WpfUiHelper] 加载 {uri} 失败: {ex.Message}");
            }
        }

        if (target is Window w)
        {
            if (target.Resources["ApplicationBackgroundBrush"] is System.Windows.Media.Brush bg)
                w.Background = bg;
            w.Foreground =
                target.Resources["TextFillColorPrimaryBrush"] as System.Windows.Media.Brush;
            w.Loaded += (_, _) => ApplyDarkTitleBar(w);
        }
    }

    /// <summary>
    /// 强制深色标题栏（激活/非激活状态都保持深色）。
    /// </summary>
    internal static void ApplyDarkTitleBar(Window window)
    {
        var hwnd = new WindowInteropHelper(window).Handle;
        if (hwnd == IntPtr.Zero)
            return;
        int dark = 1;
        DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref dark, sizeof(int));
    }

    /// <summary>
    /// 确保 System.Windows.Application 实例存在（pack:// URI 解析需要）。
    /// </summary>
    internal static void EnsureApplication()
    {
        if (System.Windows.Application.Current is null)
            _ = new System.Windows.Application
            {
                ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown,
            };
    }

    private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(
        IntPtr hwnd,
        int attr,
        ref int pvAttribute,
        int cbAttribute
    );
}
