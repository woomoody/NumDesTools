using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Threading;
using ExcelDna.Integration;
using MahApps.Metro.Controls;
using Clipboard = System.Windows.Clipboard;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class PluginLogWindow : MetroWindow
{
    private static PluginLogWindow? _instance;
    private bool _autoScroll = true;
    private readonly DispatcherTimer _drainTimer;
    private IntPtr _hwnd;
    private HwndSource? _hwndSource;

    // ── 静态入口 ──────────────────────────────────────────────────────────

    public static void EnsureOpen()
    {
        if (_instance is { IsLoaded: true })
        {
            _instance.Activate();
            ForceForegroundAndFocus(_instance._hwnd);
            return;
        }
        _instance = new PluginLogWindow(); // EnsureInitialized/SetExcelOwner 在构造函数内
        _instance.Show();
    }

    public static void CloseWindow() => _instance?.Close();

    // ── 实例 ──────────────────────────────────────────────────────────────

    public PluginLogWindow()
    {
        MahAppsHelper.EnsureInitialized(); // 必须在 InitializeComponent 之前
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        LogList.ItemsSource = PluginLog.Lines;

        // 把文件里的历史日志先刷一次
        PluginLog.DrainPendingToUi();
        UpdateStatus();
        if (PluginLog.Lines.Count > 0)
            LogList.ScrollIntoView(PluginLog.Lines[^1]);

        // 每 200ms 排空 pending 队列到 UI，完全避免跨线程操作 ObservableCollection
        _drainTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(200) };
        _drainTimer.Tick += (_, _) =>
        {
            var countBefore = PluginLog.Lines.Count;
            PluginLog.DrainPendingToUi();
            if (PluginLog.Lines.Count != countBefore)
            {
                UpdateStatus();
                if (_autoScroll && PluginLog.Lines.Count > 0)
                    LogList.ScrollIntoView(PluginLog.Lines[^1]);
            }
        };
        _drainTimer.Start();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        _hwnd = new WindowInteropHelper(this).Handle;
        _hwndSource = HwndSource.FromHwnd(_hwnd);
        _hwndSource?.AddHook(WndProcHook);

        // DispatcherPriority.Input 确保在所有输入处理完成后再抢焦点
        ForceForegroundAndFocus(_hwnd);
        Dispatcher.BeginInvoke(
            DispatcherPriority.Input,
            new System.Action(() => Keyboard.Focus(FilterBox))
        );
    }

    // ── Win32 焦点强制 ────────────────────────────────────────────────────

    private static void ForceForegroundAndFocus(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero)
            return;
        MahAppsHelper.SetForegroundWindow(hwnd);
        MahAppsHelper.SetFocus(hwnd);
    }

    private IntPtr WndProcHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        // WM_ACTIVATE = 0x0006; wParam 低字节非零 = 正在激活
        const int WmActivate = 0x0006;
        if (msg == WmActivate && (wParam.ToInt32() & 0xFFFF) != 0)
        {
            // 窗口被激活时再次强制 SetFocus，防止 Excel 吞键盘消息
            MahAppsHelper.SetFocus(hwnd);
        }
        return IntPtr.Zero;
    }

    private void UpdateStatus()
    {
        var view = CollectionViewSource.GetDefaultView(PluginLog.Lines);
        var visible = view?.Cast<string>().Count() ?? PluginLog.Lines.Count;
        StatusText.Text =
            visible == PluginLog.Lines.Count
                ? $"{PluginLog.Lines.Count} 行"
                : $"{visible}/{PluginLog.Lines.Count} 行（已过滤）";
    }

    private void FilterBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        // 点击输入框时先激活窗口 + Win32 强制焦点，防止 Excel 拦截键盘输入
        Activate();
        ForceForegroundAndFocus(_hwnd);
        Dispatcher.BeginInvoke(
            DispatcherPriority.Input,
            new System.Action(() => Keyboard.Focus(FilterBox))
        );
    }

    private void FilterBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        var view = CollectionViewSource.GetDefaultView(PluginLog.Lines);
        if (view == null)
            return;

        var text = FilterBox.Text.Trim();
        if (string.IsNullOrEmpty(text))
        {
            view.Filter = null;
        }
        else
        {
            var terms = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            view.Filter = item =>
            {
                var s = item?.ToString() ?? "";
                return terms.All(t => s.Contains(t, StringComparison.OrdinalIgnoreCase));
            };
        }
        UpdateStatus();
        if (_autoScroll && PluginLog.Lines.Count > 0)
            LogList.ScrollIntoView(PluginLog.Lines[^1]);
    }

    // ── 按钮 ──────────────────────────────────────────────────────────────

    private void BtnAutoScroll_Click(object sender, RoutedEventArgs e)
    {
        _autoScroll = !_autoScroll;
        BtnAutoScroll.Content = _autoScroll ? "自动滚动: 开" : "自动滚动: 关";
        BtnAutoScroll.Background = _autoScroll
            ? System.Windows.Media.Brushes.DarkGreen
            : System.Windows.Media.Brushes.DimGray;
    }

    private void BtnCopy_Click(object sender, RoutedEventArgs e) => CopySelected();

    private void BtnCopyAll_Click(object sender, RoutedEventArgs e) =>
        Clipboard.SetText(string.Join(Environment.NewLine, PluginLog.Lines));

    private void BtnClear_Click(object sender, RoutedEventArgs e)
    {
        PluginLog.Lines.Clear();
        UpdateStatus();
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

    // ── 键盘 ──────────────────────────────────────────────────────────────

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }

    private void LogList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            CopySelected();
    }

    private void CopySelected()
    {
        var selected = LogList.SelectedItems.Cast<string>().ToList();
        var text =
            selected.Count > 0
                ? string.Join(Environment.NewLine, selected)
                : string.Join(Environment.NewLine, PluginLog.Lines);
        if (!string.IsNullOrEmpty(text))
            Clipboard.SetText(text);
    }

    // ── 关闭 ──────────────────────────────────────────────────────────────

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _drainTimer.Stop();
        _instance = null;
        // OnLogWindowClosed 会调 Excel COM（InvalidateControl），必须回到 Excel 主线程
        ExcelAsyncUtil.QueueAsMacro(NumDesAddIn.OnLogWindowClosed);
    }
}
