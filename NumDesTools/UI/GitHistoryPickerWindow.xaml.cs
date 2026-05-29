using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MahApps.Metro.Controls;
using Button = System.Windows.Controls.Button;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using Style = System.Windows.Style;

namespace NumDesTools.UI;

public partial class GitHistoryPickerWindow : MetroWindow
{
    public record CommitEntry(string Sha, string Display);

    private readonly Func<int, int, List<CommitEntry>> _loadPage;
    private readonly ObservableCollection<CommitEntry> _items = [];
    private int _loadedCount;
    private bool _isLoading;
    private bool _hasMore = true;
    private const int PageSize = 30;

    // 调用方读结果
    public string? SelectedSha { get; private set; }
    public string? SelectedMode { get; private set; }

    // mode: "working" | "another" | "ok"（仅第二个选择器用）
    private readonly IReadOnlyList<string> _modes;

    public GitHistoryPickerWindow(
        string title,
        Func<int, int, List<CommitEntry>> loadPage,
        IReadOnlyList<string> modes,
        IReadOnlyList<CommitEntry>? preloaded = null,
        int initialIndex = 0
    )
    {
        MahAppsHelper.EnsureInitialized();
        MahAppsHelper.SetExcelOwner(this);
        InitializeComponent();

        Title = title;
        _loadPage = loadPage;
        _modes = modes;

        var screen = System.Windows.SystemParameters.WorkArea;
        Width = screen.Width * 0.65;
        Height = screen.Height * 0.72;

        CommitList.ItemsSource = _items;

        BuildButtons();

        if (preloaded is { Count: > 0 })
        {
            foreach (var e in preloaded)
                _items.Add(e);
            _loadedCount = preloaded.Count;
            _hasMore = preloaded.Count == PageSize;
            StatusText.Text = _hasMore
                ? $"已加载 {_loadedCount} 条，滚动到底加载更多"
                : $"共 {_loadedCount} 条，已全部加载";
            if (initialIndex >= 0 && initialIndex < _items.Count)
            {
                CommitList.SelectedIndex = initialIndex;
                CommitList.ScrollIntoView(_items[initialIndex]);
            }
        }

        Loaded += (_, _) =>
        {
            if (_items.Count == 0)
                LoadNextPage();
            CommitList.Focus();
        };
    }

    private void BuildButtons()
    {
        foreach (var mode in _modes)
        {
            var m = mode;
            var btn = new Button
            {
                Content = ModeLabel(m),
                Margin = new Thickness(0, 0, 8, 0),
                Padding = new Thickness(14, 6, 14, 6),
            };
            if (m == _modes[0])
                btn.Style = FindResource("MahApps.Styles.Button.Square.Accent") as Style;
            else
                btn.Style = FindResource("MahApps.Styles.Button.Square") as Style;
            btn.Click += (_, _) => Confirm(m);
            ButtonPanel.Children.Add(btn);
        }

        var cancelBtn = new Button
        {
            Content = "取消",
            Padding = new Thickness(14, 6, 14, 6),
            Style = FindResource("MahApps.Styles.Button.Square") as Style,
        };
        cancelBtn.Click += (_, _) => Close();
        ButtonPanel.Children.Add(cancelBtn);
    }

    private static string ModeLabel(string mode) =>
        mode switch
        {
            "working" => "与当前工作区对比",
            "another" => "与另一历史版本对比",
            "ok" => "开始对比",
            _ => mode,
        };

    private void Confirm(string mode)
    {
        if (CommitList.SelectedItem is not CommitEntry entry)
            return;
        SelectedSha = entry.Sha;
        SelectedMode = mode;
        DialogResult = true;
        Close();
    }

    private void LoadNextPage()
    {
        if (_isLoading || !_hasMore)
            return;
        _isLoading = true;
        StatusText.Text = "加载中…";

        int skip = _loadedCount;
        System.Threading.ThreadPool.QueueUserWorkItem(_ =>
        {
            List<CommitEntry> page;
            try
            {
                page = _loadPage(skip, PageSize);
            }
            catch
            {
                page = [];
            }

            Dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                (System.Action)(
                    () =>
                    {
                        foreach (var e in page)
                            _items.Add(e);
                        _loadedCount += page.Count;
                        _hasMore = page.Count == PageSize;
                        _isLoading = false;

                        if (CommitList.SelectedIndex < 0 && _items.Count > 0)
                            CommitList.SelectedIndex = 0;

                        StatusText.Text = _hasMore
                            ? $"已加载 {_loadedCount} 条，滚动到底加载更多"
                            : $"共 {_loadedCount} 条，已全部加载";
                    }
                )
            );
        });
    }

    private void CommitList_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (e.VerticalOffset + e.ViewportHeight >= e.ExtentHeight - 3)
            LoadNextPage();
    }

    private void CommitList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && _modes.Count > 0)
        {
            Confirm(_modes[0]);
            e.Handled = true;
        }
    }

    private void CommitList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (_modes.Count > 0)
            Confirm(_modes[0]);
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
    }

    public IReadOnlyList<CommitEntry> LoadedEntries => _items;
    public int LoadedCount => _loadedCount;
    public bool HasMore => _hasMore;
}
