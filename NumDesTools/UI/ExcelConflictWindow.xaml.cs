using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NumDesTools.ConflictResolver;
using Action     = System.Action;
using Border     = System.Windows.Controls.Border;
using Button     = System.Windows.Controls.Button;
using MessageBox = System.Windows.MessageBox;
using Window     = System.Windows.Window;
using WpfColor   = System.Windows.Media.Color;

namespace NumDesTools.UI;

public partial class ExcelConflictWindow : Window
{
    private FileDiff _diff;
    private readonly bool _autoGitAdd;
    private readonly Dictionary<string, ObservableCollection<RowConflict>> _sheetRows = new();

    public ExcelConflictWindow(FileDiff diff, string? outPath = null, bool autoGitAdd = true)
    {
        InitializeComponent();
        _diff       = diff;
        _autoGitAdd = autoGitAdd;
        _outPath    = outPath ?? diff.OursPath;

        FileNameText.Text = Path.GetFileName(diff.OursPath);
        ConflictRowItem.OnScrollOffsetChanged = offset => MetaScroll.ScrollToHorizontalOffset(offset);
        BuildSheetTabs();
        UpdateStats();
        ApplyButton.IsEnabled = diff.TotalConflictRows > 0;
    }

    private readonly string _outPath;
    private double[] _currentColWidths = Array.Empty<double>();

    // ── Sheet Tabs ───────────────────────────────────────────────────────────

    private void BuildSheetTabs()
    {
        SheetTabs.Items.Clear();
        _sheetRows.Clear();

        foreach (var sheet in _diff.Sheets)
        {
            var rows = new ObservableCollection<RowConflict>(sheet.Rows);
            _sheetRows[sheet.SheetName] = rows;

            var conflictCount = sheet.Rows.Count(r => r.DiffType != RowDiffType.Same);
            var header = conflictCount > 0
                ? $"● {sheet.SheetName} ({conflictCount})"
                : sheet.SheetName;

            var tab = new TabItem
            {
                Header     = header,
                Tag        = sheet.SheetName,
                Foreground = sheet.HasConflict
                    ? System.Windows.Media.Brushes.OrangeRed
                    : System.Windows.Media.Brushes.Gray
            };
            SheetTabs.Items.Add(tab);
        }

        if (SheetTabs.Items.Count > 0)
            SheetTabs.SelectedIndex = 0;
    }

    private void RefreshConflictList()
    {
        if (SheetTabs.SelectedItem is not TabItem tab) return;
        var sheetName = tab.Tag?.ToString() ?? string.Empty;
        if (!_sheetRows.TryGetValue(sheetName, out var allRows)) return;

        var filtered = OnlyConflictToggle.IsChecked == true
            ? allRows.Where(r => r.DiffType == RowDiffType.OnlyOurs
                              || r.DiffType == RowDiffType.OnlyTheirs
                              || (r.DiffType == RowDiffType.Modified && r.Cells.Count > 0))
            : (IEnumerable<RowConflict>)allRows;

        var sheetDiff = _diff.Sheets.FirstOrDefault(s => s.SheetName == sheetName);
        if (sheetDiff != null)
        {
            _currentColWidths = ComputeSheetColWidths(sheetDiff.AllColumns, sheetDiff.Rows);
            ConflictRowItem.CurrentSheetColWidths = _currentColWidths;
        }

        ConflictList.ItemsSource = new ObservableCollection<RowConflict>(filtered);

        if (sheetDiff != null)
        {
            BuildMetaHeader(sheetDiff);
            BuildColBatchBar(sheetDiff);
        }
    }

    // ── Meta 列头 ────────────────────────────────────────────────────────────

    private void BuildMetaHeader(SheetDiff sheet)
    {
        var cols = sheet.AllColumns;
        if (cols.Count == 0) { MetaScroll.Visibility = Visibility.Collapsed; return; }
        MetaScroll.Visibility = Visibility.Visible;

        BuildMetaGrid(MetaFieldGrid, cols, _currentColWidths,
            col => MakeMetaCell(col, "#5A9FDF", "#0D1A2A", bold: true, fontSize: 10));
        BuildMetaGrid(MetaLabelGrid, cols, _currentColWidths,
            col => MakeMetaCell(
                sheet.LabelRow.TryGetValue(col, out var v) ? v : string.Empty,
                "#999999", "#0A0A0A", bold: false, fontSize: 10));
    }

    private static void BuildMetaGrid(Grid grid, List<string> cols, double[] widths, Func<string, TextBlock> makeCell)
    {
        grid.ColumnDefinitions.Clear();
        grid.Children.Clear();
        for (int i = 0; i < cols.Count; i++)
        {
            var w = i < widths.Length ? widths[i] : 130.0;
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(w) });
            var tb = makeCell(cols[i]);
            tb.Width = w;
            Grid.SetColumn(tb, i);
            grid.Children.Add(tb);
        }
    }

    private static double[] ComputeSheetColWidths(List<string> cols, IEnumerable<RowConflict> rows)
    {
        const double min = 60, max = 220, pad = 12, fs = 11.0;
        var tf = new Typeface("Consolas");
        double Measure(string t)
        {
            if (string.IsNullOrEmpty(t)) return min;
            var ft = new FormattedText(t, CultureInfo.CurrentCulture,
                System.Windows.FlowDirection.LeftToRight, tf, fs, System.Windows.Media.Brushes.White, 1.0);
            return Math.Max(min, Math.Min(max, ft.Width + pad));
        }

        var widths = cols.Select(c => Measure(c)).ToArray();
        foreach (var rc in rows)
        {
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                if (rc.OursFullRow   != null && rc.OursFullRow.TryGetValue(col,   out var ov))
                    widths[i] = Math.Max(widths[i], Measure(ov?.ToString() ?? string.Empty));
                if (rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var tv))
                    widths[i] = Math.Max(widths[i], Measure(tv?.ToString() ?? string.Empty));
            }
        }
        return widths;
    }

    private static TextBlock MakeMetaCell(string text, string fg, string bg, bool bold, double fontSize)
    {
        var fgBrush = new SolidColorBrush((WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(fg));
        var bgBrush = new SolidColorBrush((WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(bg));
        return new TextBlock
        {
            Text         = text,
            Foreground   = fgBrush,
            Background   = bgBrush,
            Padding      = new Thickness(5, 2, 5, 2),
            FontSize     = fontSize,
            FontWeight   = bold ? FontWeights.Bold : FontWeights.Normal,
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip      = string.IsNullOrEmpty(text) ? null : text,
        };
    }

    // ── 详情面板 ─────────────────────────────────────────────────────────────

    internal void OnCellSelected(object sender, CellSelectedRoutedEventArgs e)
    {
        var rc = e.Row;
        DetailRowKey.Text = $"ID: {rc.RowKey}  [{rc.DiffTypeBadge}]";
        DetailHint.Text   = string.Empty;

        List<CellConflict> items;
        if (rc.DiffType == RowDiffType.Modified)
        {
            items = rc.Cells.ToList();
        }
        else
        {
            var src    = rc.OursFullRow ?? rc.TheirsFullRow;
            var isOurs = rc.OursFullRow != null;
            items = (src?.Keys ?? Enumerable.Empty<string>()).Select(col =>
            {
                src!.TryGetValue(col, out var val);
                return new CellConflict
                {
                    ColName     = col,
                    OursValue   = isOurs ? val : null,
                    TheirsValue = isOurs ? null : val,
                };
            }).ToList();
        }

        BuildDetailPanel(items);
    }

    private static readonly SolidColorBrush DetailFgOurs   = new(Color(0xA8, 0xC8, 0xFF));
    private static readonly SolidColorBrush DetailFgTheirs = new(Color(0xA8, 0xFF, 0xCA));
    // 字符级差异用黄色，与主列表红色背景区分
    private static readonly SolidColorBrush DetailFgDiff   = new(Color(0xFF, 0xD0, 0x40));
    private static readonly SolidColorBrush DetailBgDiff   = new(Color(0x35, 0x2A, 0x00));
    private static readonly SolidColorBrush DetailFgMuted  = new(Color(0x55, 0x55, 0x55));
    private static readonly SolidColorBrush DetailBgOurs   = new(Color(0x0A, 0x15, 0x25));
    private static readonly SolidColorBrush DetailBgTheirs = new(Color(0x0A, 0x1A, 0x0F));
    private static readonly SolidColorBrush DetailBgCol    = new(Color(0x1A, 0x2A, 0x1A));
    private static readonly SolidColorBrush DetailBorder   = new(Color(0x33, 0x33, 0x33));
    private static readonly SolidColorBrush DetailBg       = new(Color(0x1A, 0x1A, 0x1A));

    private static WpfColor Color(byte r, byte g, byte b) => WpfColor.FromRgb(r, g, b);

    private void BuildColBatchBar(SheetDiff sheet)
    {
        // 收集当前 sheet 所有 Modified 行里出现的冲突列
        var conflictCols = sheet.Rows
            .Where(r => r.DiffType == RowDiffType.Modified)
            .SelectMany(r => r.Cells.Select(c => c.ColName))
            .Distinct()
            .ToList();

        // 加上 OnlyTheirs 中 OURS 没有的新增列
        var newCols = sheet.Rows
            .Where(r => r.DiffType == RowDiffType.OnlyTheirs || r.DiffType == RowDiffType.OnlyOurs)
            .Take(1)
            .SelectMany(r => (r.TheirsFullRow ?? r.OursFullRow ?? new Dictionary<string, object?>()).Keys)
            .Except(sheet.Rows.Where(r => r.OursFullRow != null).SelectMany(r => r.OursFullRow!.Keys))
            .ToList();

        var allBatchCols = conflictCols.Concat(newCols).Distinct().ToList();

        if (allBatchCols.Count == 0)
        {
            ColBatchBar.Visibility = Visibility.Collapsed;
            return;
        }

        ColBatchBar.Visibility = Visibility.Visible;
        ColBatchPanel.Children.Clear();
        ColBatchPanel.Children.Add(new TextBlock
        {
            Text = "列批量:", Foreground = new SolidColorBrush(Color(0x88, 0x88, 0x88)),
            FontSize = 10, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0)
        });

        foreach (var col in allBatchCols)
        {
            var colCapture = col;
            // 短标签：列名超过8字符截断
            var label = col.Length > 8 ? col[..8] + "…" : col;
            ColBatchPanel.Children.Add(MakeColBatchBtn($"{label}↑我", Color(0x1A, 0x3A, 0x6E),
                () => SetColumnChoiceAndRefresh(colCapture, ConflictChoice.Ours)));
            ColBatchPanel.Children.Add(MakeColBatchBtn($"{label}↓他", Color(0x1A, 0x5C, 0x3A),
                () => SetColumnChoiceAndRefresh(colCapture, ConflictChoice.Theirs)));
        }
    }

    private void BuildDetailPanel(List<CellConflict> items)
    {
        DetailPanel.Children.Clear();

        foreach (var cell in items)
        {
            var border = new Border
            {
                Margin          = new Thickness(0, 2, 0, 2),
                Background      = DetailBg,
                BorderBrush     = DetailBorder,
                BorderThickness = new Thickness(1),
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // 列名 + 单列批量按钮
            var colPanel = new StackPanel { Background = DetailBgCol };
            colPanel.Children.Add(new TextBlock
            {
                Text = cell.ColName, Foreground = new SolidColorBrush(Color(0x88, 0xFF, 0x88)),
                FontSize = 11, FontWeight = FontWeights.Bold, TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(6, 4, 6, 2),
            });
            var colBtnRow = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, Margin = new Thickness(4, 0, 4, 4) };
            var cellCapture = cell;
            colBtnRow.Children.Add(MakeColBatchBtn("全取我的", Color(0x1A, 0x3A, 0x6E), () => { cellCapture.Choice = ConflictChoice.Ours; }));
            colBtnRow.Children.Add(MakeColBatchBtn("全取他的", Color(0x1A, 0x5C, 0x3A), () => { cellCapture.Choice = ConflictChoice.Theirs; }));
            colPanel.Children.Add(colBtnRow);
            var colBorder = new Border { Child = colPanel };
            Grid.SetColumn(colBorder, 0);
            grid.Children.Add(colBorder);

            // OURS
            var oursBorder = new Border
            {
                Background      = DetailBgOurs,
                Padding         = new Thickness(6, 4, 6, 4),
                BorderBrush     = new SolidColorBrush(Color(0x22, 0x22, 0x22)),
                BorderThickness = new Thickness(1, 0, 1, 0),
            };
            var oursPanel = new StackPanel();
            oursPanel.Children.Add(new TextBlock { Text = "我的", Foreground = DetailFgMuted, FontSize = 9, Margin = new Thickness(0,0,0,2) });
            oursPanel.Children.Add(MakeInlineDiffBlock(cell.OursDisplay, cell.TheirsDisplay, DetailFgOurs));
            oursBorder.Child = oursPanel;
            Grid.SetColumn(oursBorder, 1);
            grid.Children.Add(oursBorder);

            // THEIRS
            var theirsBorder = new Border { Background = DetailBgTheirs, Padding = new Thickness(6, 4, 6, 4) };
            var theirsPanel = new StackPanel();
            theirsPanel.Children.Add(new TextBlock { Text = "他的", Foreground = DetailFgMuted, FontSize = 9, Margin = new Thickness(0,0,0,2) });
            theirsPanel.Children.Add(MakeInlineDiffBlock(cell.TheirsDisplay, cell.OursDisplay, DetailFgTheirs));
            theirsBorder.Child = theirsPanel;
            Grid.SetColumn(theirsBorder, 2);
            grid.Children.Add(theirsBorder);

            border.Child = grid;
            DetailPanel.Children.Add(border);
        }
    }

    private static Button MakeColBatchBtn(string label, WpfColor bg, Action onClick)
    {
        var btn = new System.Windows.Controls.Button
        {
            Content         = label,
            FontSize        = 9,
            Padding         = new Thickness(5, 2, 5, 2),
            Margin          = new Thickness(0, 0, 4, 0),
            Background      = new SolidColorBrush(bg),
            Foreground      = System.Windows.Media.Brushes.White,
            BorderThickness = new Thickness(0),
            Cursor          = System.Windows.Input.Cursors.Hand,
        };
        btn.Click += (_, _) => onClick();
        return btn;
    }

    private void SetColumnChoiceAndRefresh(string colName, ConflictChoice choice)
    {
        if (SheetTabs.SelectedItem is not TabItem tab) return;
        var sheetName = tab.Tag?.ToString() ?? string.Empty;
        var sheetDiff = _diff.Sheets.FirstOrDefault(s => s.SheetName == sheetName);
        sheetDiff?.SetColumnChoice(colName, choice);
    }

    // 构建带字符级高亮的 TextBlock：公共前缀/后缀正常色，差异段黄色加粗
    private TextBlock MakeInlineDiffBlock(string text, string other, SolidColorBrush normalFg)
    {
        var tb = new TextBlock { FontSize = 11, TextWrapping = TextWrapping.Wrap };

        // 空值直接显示
        if (string.IsNullOrEmpty(text) && string.IsNullOrEmpty(other))
        {
            tb.Foreground = normalFg;
            tb.Text = "(空)";
            return tb;
        }

        var (prefixLen, suffixLen) = FindCommonPrefixSuffix(text, other);

        // 完全相同
        if (prefixLen == text.Length && prefixLen == other.Length)
        {
            tb.Foreground = normalFg;
            tb.Text = text;
            return tb;
        }

        int diffStart = prefixLen;
        int diffEnd   = text.Length - suffixLen;

        if (diffStart > 0)
            tb.Inlines.Add(new System.Windows.Documents.Run(text[..diffStart]) { Foreground = normalFg });

        if (diffStart < diffEnd)
            tb.Inlines.Add(new System.Windows.Documents.Run(text[diffStart..diffEnd])
            {
                Foreground = DetailFgDiff,
                FontWeight = FontWeights.Bold,
                Background = DetailBgDiff,
            });

        if (diffEnd < text.Length)
            tb.Inlines.Add(new System.Windows.Documents.Run(text[diffEnd..]) { Foreground = normalFg });

        return tb;
    }

    private static (int prefix, int suffix) FindCommonPrefixSuffix(string a, string b)
    {
        int maxPrefix = Math.Min(a.Length, b.Length);
        int prefix = 0;
        while (prefix < maxPrefix && a[prefix] == b[prefix]) prefix++;

        int maxSuffix = Math.Min(a.Length - prefix, b.Length - prefix);
        int suffix = 0;
        while (suffix < maxSuffix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix]) suffix++;

        return (prefix, suffix);
    }

    // ── 按钮事件 ─────────────────────────────────────────────────────────────

    private void SheetTabs_SelectionChanged(object sender, SelectionChangedEventArgs e) =>
        RefreshConflictList();

    private void FilterToggle_Changed(object sender, RoutedEventArgs e) =>
        RefreshConflictList();

    private void AllOurs_Click(object sender, RoutedEventArgs e)
    {
        SetAll(ConflictChoice.Ours);
        RefreshConflictList();
    }

    private void AllTheirs_Click(object sender, RoutedEventArgs e)
    {
        SetAll(ConflictChoice.Theirs);
        RefreshConflictList();
    }

    private void SetAll(ConflictChoice choice)
    {
        foreach (var sheet in _diff.Sheets)
        foreach (var row in sheet.Rows)
        {
            row.SetAllCells(choice);
            row.RowChoice = choice;
        }
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var savePath = _outPath;
            if (_outPath == _diff.OursPath && !_autoGitAdd)
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title            = "保存合并结果",
                    Filter           = "Excel 文件|*.xlsx",
                    FileName         = Path.GetFileName(_outPath),
                    InitialDirectory = Path.GetDirectoryName(_outPath)
                };
                if (dlg.ShowDialog() != true) return;
                savePath = dlg.FileName;
            }

            ConflictApplier.Apply(_diff, savePath, _autoGitAdd);

            MessageBox.Show(
                _autoGitAdd
                    ? $"已写回并执行 git add。\n{savePath}"
                    : $"已写回。\n{savePath}",
                "完成", MessageBoxButton.OK, MessageBoxImage.Information);

            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"写回失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    private void UpdateStats()
    {
        StatRows.Text  = _diff.TotalConflictRows.ToString();
        StatCells.Text = _diff.TotalConflictCells.ToString();
    }
}
