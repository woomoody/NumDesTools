using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelDna.Integration;
using MahApps.Metro.Controls;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class PluginLogWindow : MetroWindow
{
    private static PluginLogWindow? _instance;
    private bool _autoScroll = true;
    private readonly DispatcherTimer _drainTimer;

    // ── 静态入口 ──────────────────────────────────────────────────────────

    public static void EnsureOpen()
    {
        if (_instance is { IsLoaded: true })
        {
            _instance.Activate();
            return;
        }
        _instance = new PluginLogWindow();
        _instance.Show();
    }

    public static void CloseWindow() => _instance?.Close();

    // ── 实例 ──────────────────────────────────────────────────────────────

    public PluginLogWindow()
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        LogList.ItemsSource = PluginLog.Lines;

        PluginLog.DrainPendingToUi();
        UpdateStatus();
        if (PluginLog.Lines.Count > 0)
            LogList.ScrollIntoView(PluginLog.Lines[^1]);

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
    }

    private void UpdateStatus() => StatusText.Text = $"{PluginLog.Lines.Count} 行";

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
        ClipboardHelper.SetTextSafe(string.Join(Environment.NewLine, PluginLog.Lines));

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
            ClipboardHelper.SetTextSafe(text);
    }

    // ── 关闭 ──────────────────────────────────────────────────────────────

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        _drainTimer.Stop();
        _instance = null;
        ExcelAsyncUtil.QueueAsMacro(NumDesAddIn.OnLogWindowClosed);
    }
}
