using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NumDesTools.ExcelToLua;
using Border   = System.Windows.Controls.Border;
using CheckBox = System.Windows.Controls.CheckBox;
using Window   = System.Windows.Window;

namespace NumDesTools.UI;

public partial class GitExportSelectWindow : Window
{
    // 每个文件条目：路径 + 来源标签
    private record FileEntry(string Path, string Source);

    // ComboBox 显示用包装
    private record CommitItem(string Display, string Sha);

    private readonly string _repoBasePath;
    private readonly string _gitAuthor;
    private int _commitCount = 3;
    private List<FileEntry> _entries = [];

    // 调用方读取导出结果
    public List<string>? SelectedPaths { get; private set; }

    public GitExportSelectWindow(string repoBasePath, string gitAuthor)
    {
        InitializeComponent();
        _repoBasePath = repoBasePath;
        _gitAuthor    = gitAuthor;
        Loaded += (_, _) =>
        {
            AuthorText.Text = string.IsNullOrEmpty(gitAuthor) ? "（未知）" : gitAuthor;
            RefreshFileList();
        };
    }

    // ── 文件列表构建 ──────────────────────────────────────────────────────────

    private void RefreshFileList()
    {
        FileListPanel.Children.Clear();

        var entries = BuildEntries();
        _entries = entries;

        foreach (var entry in entries)
        {
            var row = MakeFileRow(entry);
            FileListPanel.Children.Add(row);
        }

        UpdateCountLabel();
    }

    private List<FileEntry> BuildEntries()
    {
        var seen  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var list  = new List<FileEntry>();

        // 工作区 + 暂存
        var gitFiles = SvnGitTools.GitDiffAndStagedFiles(_repoBasePath);
        foreach (var f in gitFiles)
        {
            if (IsExportable(f) && seen.Add(f))
                list.Add(new FileEntry(f, "变更/暂存"));
        }

        // 近期提交（仅在历史模式下）
        if (ModeWithHistory.IsChecked == true && !string.IsNullOrEmpty(_gitAuthor))
        {
            try
            {
                var historyFiles = SvnGitTools.GetRecentAuthorCommitFiles(_repoBasePath, _gitAuthor, _commitCount);
                foreach (var f in historyFiles)
                {
                    if (IsExportable(f))
                        list.Add(new FileEntry(f, seen.Add(f) ? $"历史×{_commitCount}" : $"历史×{_commitCount}(已含)"));
                }
            }
            catch (Exception ex)
            {
                StatusText.Text = $"历史查询失败：{ex.Message}";
            }
        }

        return list;
    }

    private static bool IsExportable(string path)
    {
        var name = System.IO.Path.GetFileName(path);
        return !name.Contains('#')
            && !name.Contains('~')
            && !path.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
            && !path.EndsWith(".xll",  StringComparison.OrdinalIgnoreCase)
            && (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
             || path.EndsWith(".xls",  StringComparison.OrdinalIgnoreCase));
    }

    private static Border MakeFileRow(FileEntry entry)
    {
        var isAlreadyContained = entry.Source.EndsWith("(已含)");

        var cb = new CheckBox
        {
            IsChecked = !isAlreadyContained,
            Tag       = entry.Path,
            Margin    = new Thickness(0),
        };

        var sourceColor = entry.Source.StartsWith("历史") ? "#88CCFF"
                        : entry.Source == "指定提交"     ? "#FFD080"
                        : "#88FF88";
        var sourceBrush = new SolidColorBrush(
            (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(sourceColor));

        var badge = new Border
        {
            Background    = new SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(
                    entry.Source.StartsWith("历史") ? "#1A3A6E"
                    : entry.Source == "指定提交"    ? "#3A2800"
                    : "#1A3A1A")),
            CornerRadius  = new CornerRadius(3),
            Padding       = new Thickness(4, 1, 4, 1),
            Margin        = new Thickness(6, 0, 0, 0),
        };
        badge.Child = new TextBlock
        {
            Text       = entry.Source,
            Foreground = sourceBrush,
            FontSize   = 9,
        };

        var nameText = new TextBlock
        {
            Text      = System.IO.Path.GetFileName(entry.Path),
            Foreground = isAlreadyContained
                ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x66, 0x66, 0x66))
                : System.Windows.Media.Brushes.White,
            FontSize  = 11,
            VerticalAlignment = VerticalAlignment.Center,
            ToolTip   = entry.Path,
            Margin    = new Thickness(6, 0, 0, 0),
            MaxWidth  = 360,
            TextTrimming = TextTrimming.CharacterEllipsis,
        };

        var panel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            Margin      = new Thickness(4, 2, 4, 2),
        };
        panel.Children.Add(cb);
        panel.Children.Add(nameText);
        panel.Children.Add(badge);

        var border = new Border
        {
            BorderBrush     = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x2A, 0x2A, 0x2A)),
            BorderThickness = new Thickness(0, 0, 0, 1),
            Child           = panel,
        };

        cb.Checked   += (_, _) => UpdateCount(border.Parent as StackPanel ?? FileListPanel_Static);
        cb.Unchecked += (_, _) => UpdateCount(border.Parent as StackPanel ?? FileListPanel_Static);

        return border;
    }

    // Panel 引用供静态 helper 用（因为 MakeFileRow 是 static）
    private static StackPanel? FileListPanel_Static;

    private void UpdateCountLabel()
    {
        FileListPanel_Static = FileListPanel;
        int total   = FileListPanel.Children.Count;
        int checked_ = CountChecked();
        FileCountText.Text = $"共 {total} 个文件，已选 {checked_} 个";
    }

    private static void UpdateCount(StackPanel? panel)
    {
        // 触发器：直接从 parent window 的 FileCountText 更新
        // 通过遍历 visual tree 找到 window
        if (panel == null) return;
        var win = Window.GetWindow(panel) as GitExportSelectWindow;
        win?.UpdateCountLabel();
    }

    private int CountChecked()
    {
        int count = 0;
        foreach (var child in FileListPanel.Children)
        {
            if (child is Border b && b.Child is StackPanel sp)
                foreach (var c in sp.Children)
                    if (c is CheckBox cb && cb.IsChecked == true)
                        count++;
        }
        return count;
    }

    private List<string> GetCheckedPaths()
    {
        var result = new List<string>();
        foreach (var child in FileListPanel.Children)
        {
            if (child is Border b && b.Child is StackPanel sp)
                foreach (var c in sp.Children)
                    if (c is CheckBox cb && cb.IsChecked == true && cb.Tag is string path)
                        result.Add(path);
        }
        return result;
    }

    // ── 事件处理 ──────────────────────────────────────────────────────────────

    private void Mode_Changed(object sender, RoutedEventArgs e)
    {
        if (HistoryPanel == null) return;
        bool historyOn = ModeWithHistory.IsChecked == true;
        bool commitOn  = ModeSpecificCommit.IsChecked == true;
        HistoryPanel.IsEnabled     = historyOn;
        HistoryAuthorRow.IsEnabled = historyOn;
        CommitPickPanel.IsEnabled  = commitOn;

        if (!commitOn) RefreshFileList();
        // 指定提交模式：等用户选择后再刷新，不自动刷新
    }

    private void LoadCommitList_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var commits = SvnGitTools.GetCommitList(_repoBasePath, 50);
            CommitCombo.ItemsSource = commits.Select(c =>
                new CommitItem(
                    $"{c.ShortSha}  {c.When:MM-dd HH:mm}  [{c.Author}]  {c.Message}",
                    c.Sha))
                .ToList();
            CommitCombo.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            StatusText.Text = $"加载提交列表失败：{ex.Message}";
        }
    }

    private void CommitCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (ModeSpecificCommit.IsChecked != true) return;
        if (CommitCombo.SelectedItem is not CommitItem item) return;
        try
        {
            var files = SvnGitTools.GetCommitFiles(_repoBasePath, item.Sha)
                .Where(IsExportable).ToList();

            FileListPanel.Children.Clear();
            _entries = files.Select(f => new FileEntry(f, "指定提交")).ToList();
            foreach (var entry in _entries)
                FileListPanel.Children.Add(MakeFileRow(entry));
            UpdateCountLabel();
        }
        catch (Exception ex)
        {
            StatusText.Text = $"读取提交文件失败：{ex.Message}";
        }
    }

    private void IncCommitCount_Click(object sender, RoutedEventArgs e)
    {
        if (_commitCount < 20) _commitCount++;
        CommitCountText.Text = _commitCount.ToString();
        if (ModeWithHistory.IsChecked == true) RefreshFileList();
    }

    private void DecCommitCount_Click(object sender, RoutedEventArgs e)
    {
        if (_commitCount > 1) _commitCount--;
        CommitCountText.Text = _commitCount.ToString();
        if (ModeWithHistory.IsChecked == true) RefreshFileList();
    }

    private void Refresh_Click(object sender, RoutedEventArgs e)
    {
        if (ModeSpecificCommit.IsChecked == true)
            CommitCombo_SelectionChanged(sender, null!);
        else
            RefreshFileList();
    }

    private void SelectAll_Checked(object sender, RoutedEventArgs e)   => SetAllChecked(true);
    private void SelectAll_Unchecked(object sender, RoutedEventArgs e) => SetAllChecked(false);

    private void SetAllChecked(bool value)
    {
        foreach (var child in FileListPanel.Children)
            if (child is Border b && b.Child is StackPanel sp)
                foreach (var c in sp.Children)
                    if (c is CheckBox cb) cb.IsChecked = value;
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        SelectedPaths = GetCheckedPaths();
        DialogResult  = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
    private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape) Close();
    }

}
