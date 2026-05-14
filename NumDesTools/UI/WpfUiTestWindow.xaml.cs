using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

public partial class WpfUiTestWindow : Window
{
    public WpfUiTestWindow()
    {
        MergeWpfUiTheme(this);
        InitializeComponent();

        if (Resources["ApplicationBackgroundBrush"] is System.Windows.Media.Brush bg)
            Background = bg;

        Loaded += (_, _) => ApplyDarkTitleBar(this);
    }

    /// <summary>
    /// 强制深色标题栏，避免非激活状态变白。
    /// 所有使用 WPF-UI 深色主题的 Window 都应在 Loaded 后调用。
    /// </summary>
    internal static void ApplyDarkTitleBar(Window window)
    {
        var hwnd = new WindowInteropHelper(window).Handle;
        if (hwnd == IntPtr.Zero)
            return;
        int dark = 1;
        DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref dark, sizeof(int));
    }

    internal static void MergeWpfUiTheme(FrameworkElement target)
    {
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
                PluginLog.Write($"[WpfUiTheme] 加载 {uri} 失败: {ex.Message}");
            }
        }
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
