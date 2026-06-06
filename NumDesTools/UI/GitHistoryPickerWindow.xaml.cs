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
    public record CommitEntry(string Sha, string Display, string Author = "");

    private readonly Func<int, int, List<CommitEntry>> _loadPage;
    private readonly List<CommitEntry> _allItems = [];
    private readonly ObservableCollection<CommitEntry> _displayItems = [];
    private int _loadedCount;
    private bool _isLoading;
    private bool _hasMore = true;
    private bool _filterLoading;
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

        CommitList.ItemsSource = _displayItems;

        BuildButtons();

        if (preloaded is { Count: > 0 })
        {
            foreach (var e in preloaded)
                _allItems.Add(e);
            _loadedCount = preloaded.Count;
            _hasMore = preloaded.Count == PageSize;
            ApplyFilter();
            if (initialIndex >= 0 && initialIndex < _displayItems.Count)
            {
                CommitList.SelectedIndex = initialIndex;
                CommitList.ScrollIntoView(_displayItems[initialIndex]);
            }
        }

        Loaded += (_, _) =>
        {
            if (_allItems.Count == 0)
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
                            _allItems.Add(e);
                        _loadedCount += page.Count;
                        _hasMore = page.Count == PageSize;
                        _isLoading = false;
                        ApplyFilter();
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

    private void ApplyFilter()
    {
        var author = AuthorFilterBox.SelectedItem as string;
        var filtered =
            string.IsNullOrEmpty(author) || author == "(全部)"
                ? _allItems
                : _allItems.Where(c => c.Author == author).ToList();

        _displayItems.Clear();
        foreach (var e in filtered)
            _displayItems.Add(e);

        if (CommitList.SelectedIndex < 0 && _displayItems.Count > 0)
            CommitList.SelectedIndex = 0;

        // 追加新出现的作者到 ComboBox（保留当前选中）
        var existing = AuthorFilterBox.Items.Cast<string>().ToHashSet();
        var newAuthors = _allItems
            .Select(c => c.Author)
            .Where(a => !string.IsNullOrEmpty(a) && !existing.Contains(a))
            .Distinct()
            .OrderBy(a => a)
            .ToList();
        if (newAuthors.Count > 0)
        {
            var prev = AuthorFilterBox.SelectedItem as string;
            _filterLoading = true;
            if (AuthorFilterBox.Items.Count == 0)
                AuthorFilterBox.Items.Add("(全部)");
            // 重建排序后的完整列表
            var allAuthors = existing
                .Where(a => a != "(全部)")
                .Concat(newAuthors)
                .Distinct()
                .OrderBy(a => a)
                .ToList();
            AuthorFilterBox.Items.Clear();
            AuthorFilterBox.Items.Add("(全部)");
            foreach (var a in allAuthors)
                AuthorFilterBox.Items.Add(a);
            AuthorFilterBox.SelectedItem =
                prev != null && AuthorFilterBox.Items.Contains(prev) ? prev : "(全部)";
            _filterLoading = false;
        }
        else if (AuthorFilterBox.Items.Count == 0)
        {
            _filterLoading = true;
            AuthorFilterBox.Items.Add("(全部)");
            AuthorFilterBox.SelectedIndex = 0;
            _filterLoading = false;
        }

        var suffix =
            filtered.Count < _allItems.Count ? $"，筛选后 {filtered.Count} 条" : "";
        StatusText.Text = _hasMore
            ? $"已加载 {_loadedCount} 条{suffix}，滚动到底加载更多"
            : $"共 {_loadedCount} 条{suffix}，已全部加载";
    }

    private void AuthorFilter_Changed(object sender, SelectionChangedEventArgs e)
    {
        if (_filterLoading)
            return;
        ApplyFilter();
    }

    public IReadOnlyList<CommitEntry> LoadedEntries => _allItems;
    public int LoadedCount => _loadedCount;
    public bool HasMore => _hasMore;
}
