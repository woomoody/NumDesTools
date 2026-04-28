using System.IO;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Newtonsoft.Json;
using Key              = System.Windows.Input.Key;
using KeyEventArgs     = System.Windows.Input.KeyEventArgs;
using RoutedEventArgs  = System.Windows.RoutedEventArgs;
using Visibility       = System.Windows.Visibility;
using Window           = System.Windows.Window;
using MouseButtonEventArgs = System.Windows.Input.MouseButtonEventArgs;

namespace NumDesTools.UI;

public partial class ExcelFilePickerWindow : Window
{
    private record FileEntry(string FullPath, string RelPath);
    private record HistoryEntry(string Keyword, int Count);

    private const string HistoryKey = "FilePickerSearchHistory";
    private const int HistoryMaxRaw = 200;
    private const int HistoryMaxKept = 50;

    private List<FileEntry> _allFiles = [];
    private List<FileEntry> _filtered = [];
    private List<HistoryEntry> _history = [];

    public string? SelectedFile { get; private set; }

    private readonly string _rootDir;

    public ExcelFilePickerWindow(string rootDir)
    {
        InitializeComponent();
        _rootDir = rootDir;
        RootLabel.Text = rootDir;
        LoadHistory();
        Loaded += (_, _) =>
        {
            LoadFiles();
            SearchBox.Focus();
        };
    }

    // ── History ──────────────────────────────────────────────────────────────

    private void LoadHistory()
    {
        if (!NumDesAddIn.GlobalValue.Value.TryGetValue(HistoryKey, out var json) || string.IsNullOrEmpty(json))
            return;
        try { _history = JsonConvert.DeserializeObject<List<HistoryEntry>>(json) ?? []; }
        catch { _history = []; }
    }

    private void SaveHistory()
    {
        var json = JsonConvert.SerializeObject(_history);
        NumDesAddIn.GlobalValue.SaveValue(HistoryKey, json);
    }

    private void RecordSearch(string keyword)
    {
        if (string.IsNullOrWhiteSpace(keyword)) return;
        var idx = _history.FindIndex(h => string.Equals(h.Keyword, keyword, StringComparison.OrdinalIgnoreCase));
        if (idx >= 0)
            _history[idx] = _history[idx] with { Count = _history[idx].Count + 1 };
        else
            _history.Add(new HistoryEntry(keyword, 1));

        if (_history.Count > HistoryMaxRaw)
            _history = [.. _history.OrderByDescending(h => h.Count).Take(HistoryMaxKept)];

        SaveHistory();
    }

    private void ShowHistoryPopup()
    {
        if (_history.Count == 0) return;
        HistoryList.ItemsSource = _history
            .OrderByDescending(h => h.Count)
            .Select(h => h.Keyword)
            .ToList();
        HistoryPopup.IsOpen = true;
    }

    private void ApplyHistoryItem(string keyword)
    {
        HistoryPopup.IsOpen = false;
        SearchBox.Text = keyword;
        SearchBox.CaretIndex = keyword.Length;
        SearchBox.Focus();
    }

    private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(SearchBox.Text))
            ShowHistoryPopup();
    }

    private void SearchBox_LostFocus(object sender, RoutedEventArgs e)
    {
        // Popup StaysOpen=False handles close; only close if focus didn't go to HistoryList
        if (!HistoryList.IsKeyboardFocusWithin)
            HistoryPopup.IsOpen = false;
    }

    private void HistoryList_MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (HistoryList.SelectedItem is string kw)
            ApplyHistoryItem(kw);
    }

    private void HistoryList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && HistoryList.SelectedItem is string kw)
        {
            ApplyHistoryItem(kw);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            HistoryPopup.IsOpen = false;
            SearchBox.Focus();
            e.Handled = true;
        }
    }

    private void LoadFiles()
    {
        var all = new List<FileEntry>();
        try
        {
            foreach (var f in Directory.EnumerateFiles(_rootDir, "*.xls*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xlsm") continue;
                var rel = Path.GetRelativePath(_rootDir, f).Replace('\\', '/');
                all.Add(new FileEntry(f, rel));
            }
        }
        catch { }

        _allFiles = [.. all.OrderBy(x => x.RelPath)];
        ApplyFilter();
    }

    private void ApplyFilter()
    {
        if (SearchBox == null || FilterTilde == null || FilterHash == null || FilterXlsm == null || FileList == null) return;

        var filterTilde = FilterTilde.IsChecked == true;
        var filterHash  = FilterHash.IsChecked == true;
        var filterXlsm  = FilterXlsm.IsChecked == true;
        var keyword     = SearchBox.Text.Trim();

        var list = _allFiles.AsEnumerable();

        if (filterTilde)
            list = list.Where(f => !Path.GetFileName(f.FullPath).StartsWith('~'));

        if (filterHash)
            list = list.Where(f => !Path.GetFileName(f.FullPath).StartsWith('#') &&
                                   !f.RelPath.Split('/').Any(seg => seg.StartsWith('#')));

        if (filterXlsm)
            list = list.Where(f => !f.FullPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrEmpty(keyword))
            list = list
                .Select(f => (f, score: Score(f.RelPath, keyword)))
                .Where(x => x.score > 0)
                .OrderByDescending(x => x.score)
                .Select(x => x.f);

        _filtered = [.. list];
        FileList.ItemsSource = _filtered;
        CountLabel.Text = $"{_filtered.Count} 个文件";

        if (_filtered.Count > 0)
            FileList.SelectedIndex = 0;

        UpdateSelectedLabel();
    }

    // 简单评分：连续字符命中得高分，分散命中得低分
    private static int Score(string path, string keyword)
    {
        var name = Path.GetFileName(path).ToLowerInvariant();
        var kw   = keyword.ToLowerInvariant();

        if (name.Contains(kw)) return 100 + (name.StartsWith(kw) ? 50 : 0);

        int idx = 0;
        int score = 0;
        foreach (var ch in name)
        {
            if (idx < kw.Length && ch == kw[idx])
            {
                score++;
                idx++;
            }
        }
        return idx == kw.Length ? score : 0;
    }

    private void UpdateSelectedLabel()
    {
        if (FileList.SelectedItem is FileEntry fe)
            SelectedPathLabel.Text = fe.RelPath;
        else
            SelectedPathLabel.Text = string.Empty;
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        SearchPlaceholder.Visibility = string.IsNullOrEmpty(SearchBox.Text)
            ? Visibility.Visible : Visibility.Collapsed;
        if (!string.IsNullOrEmpty(SearchBox.Text))
            HistoryPopup.IsOpen = false;
        else if (SearchBox.IsFocused)
            ShowHistoryPopup();
        ApplyFilter();
    }

    private void Filter_Changed(object sender, RoutedEventArgs e) => ApplyFilter();

    private void FileList_SelectionChanged(object sender, SelectionChangedEventArgs e) => UpdateSelectedLabel();

    private void FileList_DoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (FileList.SelectedItem is FileEntry) Confirm();
    }

    private void SearchBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Down)
        {
            if (HistoryPopup.IsOpen && HistoryList.Items.Count > 0)
            {
                HistoryList.Focus();
                HistoryList.SelectedIndex = 0;
                (HistoryList.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem)?.Focus();
                e.Handled = true;
            }
            else if (_filtered.Count > 0)
            {
                FileList.Focus();
                FileList.SelectedIndex = Math.Max(0, FileList.SelectedIndex);
                (FileList.ItemContainerGenerator.ContainerFromIndex(FileList.SelectedIndex) as ListBoxItem)?.Focus();
                e.Handled = true;
            }
        }
        else if (e.Key == Key.Enter)
        {
            if (HistoryPopup.IsOpen) HistoryPopup.IsOpen = false;
            Confirm();
            e.Handled = true;
        }
        else if (e.Key == Key.Escape && HistoryPopup.IsOpen)
        {
            HistoryPopup.IsOpen = false;
            e.Handled = true;
        }
    }

    private void FileList_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter) { Confirm(); e.Handled = true; }
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) { DialogResult = false; Close(); }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)     => Confirm();
    private void Cancel_Click(object sender, RoutedEventArgs e) { DialogResult = false; Close(); }

    private void Confirm()
    {
        if (FileList.SelectedItem is not FileEntry fe) return;
        RecordSearch(SearchBox.Text.Trim());
        SelectedFile = fe.FullPath;
        DialogResult = true;
        Close();
    }
}
