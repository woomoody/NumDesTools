using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MahApps.Metro.Controls;
using NumDesTools.ConflictResolver;
using Action = System.Action;
using Border = System.Windows.Controls.Border;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
using Window = System.Windows.Window;
using WpfColor = System.Windows.Media.Color;

namespace NumDesTools.UI;

public partial class ExcelConflictWindow : MetroWindow
{
    private FileDiff _diff;
    private readonly bool _autoGitAdd;
    private readonly Dictionary<string, ObservableCollection<RowConflict>> _sheetRows = new();

    public ExcelConflictWindow(FileDiff diff, string? outPath = null, bool autoGitAdd = true)
    {
        MahAppsHelper.EnsureInitialized();
        _suppressRefresh = true;
        InitializeComponent();
        _diff = diff;
        _autoGitAdd = autoGitAdd;
        _outPath = outPath ?? diff.OursPath;

        FileNameText.Text = Path.GetFileName(diff.OursPath);
        // 冲突行（Modified）：驱动冲突行列头 + 冲突行滚动条
        ConflictRowItem.OnConflictScrollOffsetChanged = offset =>
        {
            MetaScroll.ScrollToHorizontalOffset(offset);
            BatchScroll.ScrollToHorizontalOffset(offset);
            SharedHScrollBar.Value = offset;
        };
        // 新增/删除行（OnlyOurs/OnlyTheirs）：驱动各自列头 + 新增/删除滚动条
        ConflictRowItem.OnRowsScrollOffsetChanged = offset =>
        {
            OursMetaScroll.ScrollToHorizontalOffset(offset);
            TheirsMetaScroll.ScrollToHorizontalOffset(offset);
            RowsHScrollBar.Value = offset;
        };
        BuildSheetTabs();
        UpdateStats();
        ApplyButton.IsEnabled = diff.TotalConflictRows > 0;

        // 拖拽多选 + 单行切换选中
        ConflictList.PreviewMouseLeftButtonDown += ConflictList_DragStart;
        ConflictList.PreviewMouseMove           += ConflictList_DragMove;
        ConflictList.PreviewMouseLeftButtonUp   += ConflictList_DragEnd;
        ConflictList.PreviewMouseDown           += ConflictList_MiddleClick;

        // 三个列表点击任意位置都触发详情（兜底：路由事件冒泡失效时仍有效）
        ConflictList.PreviewMouseLeftButtonDown   += ListBox_ShowDetail;
        OnlyOursList.PreviewMouseLeftButtonDown   += ListBox_ShowDetail;
        OnlyTheirsList.PreviewMouseLeftButtonDown += ListBox_ShowDetail;

        // 把首次列表渲染推到窗口显示后，避免构造函数阻塞 ShowDialog
        Loaded += (_, _) =>
            Dispatcher.BeginInvoke(
                System.Windows.Threading.DispatcherPriority.Background,
                (System.Action)RefreshConflictList
            );

        // 窗口/控件尺寸变化时重新计算滚动条范围（解决首次布局 ActualWidth=0 问题）
        SizeChanged += (_, _) =>
        {
            ApplyScrollBarRange(SharedHScrollBar, _conflictTotalWidth);
            UpdateRowsScrollBarVisibility();
        };
    }

    private enum ViewMode
    {
        ConflictOnly,
        Context,
        All
    }

    private readonly string _outPath;
    private double[] _currentColWidths = Array.Empty<double>();
    private double[] _allColWidths = Array.Empty<double>();
    private HashSet<string> _conflictColSet = new(StringComparer.Ordinal);
    private bool _suppressRefresh;

    // section 展开状态（VSCode 风格：折叠=28px标题栏，展开=记忆高度）
    private bool _sectionOnlyOursExpanded = false;
    private bool _sectionOnlyTheirsExpanded = false;
    private bool _detailExpanded = true;
    private double _oursExpandedHeight = 200;
    private double _theirsExpandedHeight = 200;
    private double _detailExpandedHeight = 180;

    // 滚动条总宽度（列宽之和），SizeChanged 时重新计算 Maximum/Viewport
    private double _conflictTotalWidth = 0;
    private double _rowsTotalWidth = 0;

    // 拖拽多选
    private bool _isDragging;
    private int _dragStartIndex = -1;
    private List<RowConflict> _selectedRows = [];

    // 全量模式分批加载
    private const int PageSize = 200;
    private const int ConflictContext = 5; // 冲突行上下各 N 行
    private List<RowConflict> _pendingRows = []; // 待加载的完整列表
    private int _loadedCount;
    private ViewMode _viewMode = ViewMode.ConflictOnly;

    // ── Sheet Tabs ───────────────────────────────────────────────────────────

    private void BuildSheetTabs()
    {
        _suppressRefresh = true;
        SheetTabs.Items.Clear();
        _sheetRows.Clear();

        foreach (var sheet in _diff.Sheets)
        {
            var rows = new ObservableCollection<RowConflict>(sheet.Rows);
            _sheetRows[sheet.SheetName] = rows;

            var conflictCount = sheet.Rows.Count(r => r.DiffType != RowDiffType.Same);
            var header =
                conflictCount > 0 ? $"● {sheet.SheetName} ({conflictCount})" : sheet.SheetName;

            var tab = new TabItem
            {
                Header = header,
                Tag = sheet.SheetName,
                Foreground = sheet.HasConflict
                    ? System.Windows.Media.Brushes.OrangeRed
                    : System.Windows.Media.Brushes.Gray
            };
            SheetTabs.Items.Add(tab);
        }

        if (SheetTabs.Items.Count > 0)
            SheetTabs.SelectedIndex = 0;
        _suppressRefresh = false;
    }

    private void RefreshConflictList()
    {
        if (SheetTabs.SelectedItem is not TabItem tab)
            return;
        var sheetName = tab.Tag?.ToString() ?? string.Empty;
        if (!_sheetRows.TryGetValue(sheetName, out var allRows))
            return;

        var sheetDiff = _diff.Sheets.FirstOrDefault(s => s.SheetName == sheetName);

        _conflictColSet =
            sheetDiff
                ?.Rows.Where(r => r.DiffType == RowDiffType.Modified)
                .SelectMany(r => r.Cells.Select(c => c.ColName))
                .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);

        bool hideNonConflict = HideNoConflictCols.IsChecked == true;
        if (sheetDiff != null)
        {
            List<string>? visibleCols =
                (hideNonConflict && _conflictColSet.Count > 0)
                    ? sheetDiff.AllColumns.Where(c => _conflictColSet.Contains(c)).ToList()
                    : null;
            ConflictRowItem.VisibleColumns = visibleCols;

            _allColWidths = ComputeSheetColWidths(
                sheetDiff.AllColumns,
                sheetDiff.Rows,
                _conflictColSet
            );
            ConflictRowItem.AllSheetColWidths = _allColWidths;

            var colsForWidths = visibleCols ?? sheetDiff.AllColumns;
            _currentColWidths =
                visibleCols != null
                    ? ComputeSheetColWidths(colsForWidths, sheetDiff.Rows, _conflictColSet)
                    : _allColWidths;
            ConflictRowItem.CurrentSheetColWidths = _currentColWidths;

            // 列宽确定后一次性设置两条滚动条，后续行渲染不再重复触发
            _conflictTotalWidth = _currentColWidths.Sum();
            _rowsTotalWidth = _allColWidths.Sum();
            ApplyScrollBarRange(SharedHScrollBar, _conflictTotalWidth);
            UpdateRowsScrollBarVisibility();
        }

        // Split rows into three sections
        var modifiedRows = allRows
            .Where(r => r.DiffType == RowDiffType.Modified && r.Cells.Count > 0)
            .ToList();
        var onlyOursRows = allRows.Where(r => r.DiffType == RowDiffType.OnlyOurs).ToList();
        var onlyTheirsRows = allRows.Where(r => r.DiffType == RowDiffType.OnlyTheirs).ToList();

        // For Context/All modes, include Same rows in the modified section only
        List<RowConflict> displayModified;
        switch (_viewMode)
        {
            case ViewMode.Context:
            {
                var allList = allRows.ToList();
                _pendingRows = allList;
                _loadedCount = 0;
                var contextRows = BuildConflictContextRows(allList)
                    .Where(r =>
                        r.DiffType != RowDiffType.OnlyOurs && r.DiffType != RowDiffType.OnlyTheirs
                    )
                    .ToList();
                LoadMoreBar.Visibility =
                    contextRows.Count
                    < allList.Count(r =>
                        r.DiffType != RowDiffType.OnlyOurs && r.DiffType != RowDiffType.OnlyTheirs
                    )
                        ? Visibility.Visible
                        : Visibility.Collapsed;
                LoadMoreStatus.Text = $"冲突±{ConflictContext}行：已显示 {contextRows.Count}";
                displayModified = contextRows;
                break;
            }
            case ViewMode.All:
            {
                var allList = allRows
                    .Where(r =>
                        r.DiffType != RowDiffType.OnlyOurs && r.DiffType != RowDiffType.OnlyTheirs
                    )
                    .ToList();
                if (allList.Count <= PageSize)
                {
                    _pendingRows = [];
                    _loadedCount = 0;
                    LoadMoreBar.Visibility = Visibility.Collapsed;
                    displayModified = allList;
                }
                else
                {
                    _pendingRows = allList;
                    _loadedCount = PageSize;
                    LoadMoreBar.Visibility = Visibility.Visible;
                    LoadMoreStatus.Text = $"已显示 {PageSize}/{allList.Count} 行";
                    displayModified = allList.Take(PageSize).ToList();
                }
                break;
            }
            default: // ConflictOnly
                _pendingRows = [];
                _loadedCount = 0;
                LoadMoreBar.Visibility = Visibility.Collapsed;
                displayModified = modifiedRows;
                break;
        }

        ConflictList.ItemsSource = new ObservableCollection<RowConflict>(displayModified);
        OnlyOursList.ItemsSource = new ObservableCollection<RowConflict>(onlyOursRows);
        OnlyTheirsList.ItemsSource = new ObservableCollection<RowConflict>(onlyTheirsRows);

        SectionOnlyOursCount.Text = $"({onlyOursRows.Count})";
        SectionOnlyTheirsCount.Text = $"({onlyTheirsRows.Count})";

        // 新增/删除 section：无数据时折叠且隐藏（行高=0）
        if (onlyOursRows.Count == 0)
        {
            RowOnlyOurs.MinHeight = 0;
            RowOnlyOurs.Height = new GridLength(0);
        }
        else
        {
            SetSectionExpanded(
                RowOnlyOurs,
                OnlyOursList,
                OursHeaderBorder,
                SectionOnlyOursChevron,
                _sectionOnlyOursExpanded,
                ref _oursExpandedHeight
            );
        }

        if (onlyTheirsRows.Count == 0)
        {
            RowOnlyTheirs.MinHeight = 0;
            RowOnlyTheirs.Height = new GridLength(0);
        }
        else
        {
            SetSectionExpanded(
                RowOnlyTheirs,
                OnlyTheirsList,
                TheirsHeaderBorder,
                SectionOnlyTheirsChevron,
                _sectionOnlyTheirsExpanded,
                ref _theirsExpandedHeight
            );
        }

        if (sheetDiff != null)
        {
            BuildMetaHeader(sheetDiff);
            BuildSectionMetaHeader(
                sheetDiff,
                OursMetaFieldGrid,
                OursMetaLabelGrid,
                sheetDiff.AllColumns,
                _allColWidths
            );
            BuildSectionMetaHeader(
                sheetDiff,
                TheirsMetaFieldGrid,
                TheirsMetaLabelGrid,
                sheetDiff.AllColumns,
                _allColWidths
            );
            BuildColBatchBar(sheetDiff);
            BuildChangeNav(sheetDiff);
        }

        void OnLayout(object? s, EventArgs _)
        {
            ConflictList.LayoutUpdated -= OnLayout;
            ScrollToFirstConflictColumn();
        }
        ConflictList.LayoutUpdated += OnLayout;
    }

    // VSCode 风格：折叠 = RowDefinition.Height 设为 28（标题栏高），展开 = 恢复记忆高度
    private void SetSectionExpanded(
        RowDefinition row,
        System.Windows.Controls.ListBox list,
        Border? colHeader,
        TextBlock chevron,
        bool expanded,
        ref double savedHeight
    )
    {
        if (expanded)
        {
            row.MinHeight = 28;
            row.Height = new GridLength(Math.Max(savedHeight, 120));
        }
        else
        {
            // 记住当前高度再折叠
            if (row.ActualHeight > 28)
                savedHeight = row.ActualHeight;
            row.MinHeight = 28;
            row.Height = new GridLength(28);
        }
        list.Visibility = expanded ? Visibility.Visible : Visibility.Collapsed;
        if (colHeader != null)
            colHeader.Visibility = expanded ? Visibility.Visible : Visibility.Collapsed;
        chevron.Text = expanded ? "▼ " : "▶ ";
        UpdateRowsScrollBarVisibility();
    }

    private void UpdateRowsScrollBarVisibility()
    {
        if (_sectionOnlyOursExpanded || _sectionOnlyTheirsExpanded)
            ApplyScrollBarRange(RowsHScrollBar, _rowsTotalWidth);
        else
            RowsHScrollBar.Visibility = Visibility.Collapsed;
    }

    private void SectionHeader_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (sender == SectionOnlyOursHeader)
        {
            _sectionOnlyOursExpanded = !_sectionOnlyOursExpanded;
            SetSectionExpanded(
                RowOnlyOurs,
                OnlyOursList,
                OursHeaderBorder,
                SectionOnlyOursChevron,
                _sectionOnlyOursExpanded,
                ref _oursExpandedHeight
            );
        }
        else if (sender == SectionOnlyTheirsHeader)
        {
            _sectionOnlyTheirsExpanded = !_sectionOnlyTheirsExpanded;
            SetSectionExpanded(
                RowOnlyTheirs,
                OnlyTheirsList,
                TheirsHeaderBorder,
                SectionOnlyTheirsChevron,
                _sectionOnlyTheirsExpanded,
                ref _theirsExpandedHeight
            );
        }
    }

    private void DetailHeader_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        _detailExpanded = !_detailExpanded;
        if (_detailExpanded)
        {
            RowDetail.MinHeight = 28;
            RowDetail.Height = new GridLength(Math.Max(_detailExpandedHeight, 120));
        }
        else
        {
            if (RowDetail.ActualHeight > 28)
                _detailExpandedHeight = RowDetail.ActualHeight;
            RowDetail.MinHeight = 28;
            RowDetail.Height = new GridLength(28);
        }
        DetailContent.Visibility = _detailExpanded ? Visibility.Visible : Visibility.Collapsed;
        DetailChevron.Text = _detailExpanded ? "▼ " : "▶ ";
    }

    private void ScrollToFirstConflictColumn()
    {
        if (_currentColWidths.Length == 0)
            return;
        if (SheetTabs.SelectedItem is not TabItem tab)
            return;
        var sheetName = tab.Tag?.ToString() ?? string.Empty;
        var sheetDiff = _diff.Sheets.FirstOrDefault(s => s.SheetName == sheetName);
        if (sheetDiff == null)
            return;

        var cols = sheetDiff.AllColumns;
        bool hasRowConflict = sheetDiff.Rows.Any(r =>
            r.DiffType == RowDiffType.OnlyOurs || r.DiffType == RowDiffType.OnlyTheirs
        );

        double offset = 0;
        bool found = false;
        for (int i = 0; i < cols.Count && i < _currentColWidths.Length; i++)
        {
            if (_conflictColSet.Contains(cols[i]) || (hasRowConflict && i == 0))
            {
                found = true;
                break;
            }
            offset += _currentColWidths[i];
        }

        if (!found)
            return;

        const double margin = 40;
        var scrollTo = Math.Max(0, offset - margin);
        ConflictRowItem.InitialScrollOffset = scrollTo;
        ConflictRowItem.SetGlobalScrollOffset(scrollTo);
        MetaScroll.ScrollToHorizontalOffset(scrollTo);
        BatchScroll.ScrollToHorizontalOffset(scrollTo);
        SharedHScrollBar.Value = scrollTo;
        // 新增/删除行从列头第0列开始
        ConflictRowItem.InitialRowsScrollOffset = 0;
        ConflictRowItem.SetGlobalRowsScrollOffset(0);
    }

    private static List<RowConflict> BuildConflictContextRows(List<RowConflict> rows)
    {
        var indices = new HashSet<int>();
        for (int i = 0; i < rows.Count; i++)
        {
            if (rows[i].DiffType == RowDiffType.Same)
                continue;
            for (int d = -ConflictContext; d <= ConflictContext; d++)
            {
                var idx = i + d;
                if (idx >= 0 && idx < rows.Count)
                    indices.Add(idx);
            }
        }
        return rows.Where((_, i) => indices.Contains(i)).ToList();
    }

    private void LoadMore_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingRows.Count == 0)
            return;
        // 追加模式：把 _loadedCount 视为"从全量 pending 加载了多少"
        // 首次点加载更多时先从头加载 PageSize（因为当前显示的是 context 子集）
        if (_loadedCount == 0)
        {
            // 切换到分页模式：用前 PageSize 行替换当前 context 视图
            _loadedCount = PageSize;
            ConflictList.ItemsSource = new ObservableCollection<RowConflict>(
                _pendingRows.Take(PageSize)
            );
        }
        else
        {
            var next = _pendingRows.Skip(_loadedCount).Take(PageSize).ToList();
            if (next.Count == 0)
            {
                LoadMoreBar.Visibility = Visibility.Collapsed;
                return;
            }
            var current = (ObservableCollection<RowConflict>)ConflictList.ItemsSource;
            foreach (var row in next)
                current.Add(row);
            _loadedCount += next.Count;
        }

        if (_loadedCount >= _pendingRows.Count)
            LoadMoreStatus.Text = $"已显示全部 {_loadedCount} 行";
        else
            LoadMoreStatus.Text = $"已显示 {_loadedCount}/{_pendingRows.Count} 行";
    }

    private void LoadAll_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingRows.Count == 0)
            return;
        _loadedCount = _pendingRows.Count;
        ConflictList.ItemsSource = new ObservableCollection<RowConflict>(_pendingRows);
        LoadMoreBar.Visibility = Visibility.Collapsed;
    }

    private void ApplyScrollBarRange(
        System.Windows.Controls.Primitives.ScrollBar bar,
        double totalWidth
    )
    {
        if (totalWidth <= 0)
        {
            bar.Visibility = Visibility.Collapsed;
            return;
        }
        // 先 Visible，让 WPF 完成布局，再读 ActualWidth
        bar.Visibility = Visibility.Visible;
        bar.UpdateLayout();
        var viewport = bar.ActualWidth > 0 ? bar.ActualWidth : ActualWidth;
        bar.ViewportSize = viewport;
        bar.Maximum = Math.Max(0, totalWidth - viewport);
        if (totalWidth <= viewport)
            bar.Visibility = Visibility.Collapsed;
    }

    private void SharedHScrollBar_Scroll(
        object sender,
        System.Windows.Controls.Primitives.ScrollEventArgs e
    )
    {
        var offset = e.NewValue;
        ConflictRowItem.SetGlobalScrollOffset(offset);
        MetaScroll.ScrollToHorizontalOffset(offset);
        BatchScroll.ScrollToHorizontalOffset(offset);
    }

    private void RowsHScrollBar_Scroll(
        object sender,
        System.Windows.Controls.Primitives.ScrollEventArgs e
    )
    {
        var offset = e.NewValue;
        ConflictRowItem.SetGlobalRowsScrollOffset(offset);
        OursMetaScroll.ScrollToHorizontalOffset(offset);
        TheirsMetaScroll.ScrollToHorizontalOffset(offset);
    }

    // ── Meta 列头 ────────────────────────────────────────────────────────────

    private static readonly SolidColorBrush MetaFieldFg = new(Color(0x5A, 0x9F, 0xDF));
    private static readonly SolidColorBrush MetaFieldBg = new(Color(0x0D, 0x1A, 0x2A));
    private static readonly SolidColorBrush MetaLabelFg = new(Color(0x99, 0x99, 0x99));
    private static readonly SolidColorBrush MetaLabelBg = new(Color(0x0A, 0x0A, 0x0A));

    private void BuildMetaHeader(SheetDiff sheet)
    {
        var cols = ConflictRowItem.VisibleColumns ?? sheet.AllColumns;
        if (cols.Count == 0)
        {
            MetaScroll.Visibility = Visibility.Collapsed;
            return;
        }
        MetaScroll.Visibility = Visibility.Visible;

        BuildMetaGrid(
            MetaFieldGrid,
            cols,
            _currentColWidths,
            col => MakeMetaCell(col, MetaFieldFg, MetaFieldBg, bold: true)
        );
        BuildMetaGrid(
            MetaLabelGrid,
            cols,
            _currentColWidths,
            col =>
                MakeMetaCell(
                    sheet.LabelRow.TryGetValue(col, out var v) ? v : string.Empty,
                    MetaLabelFg,
                    MetaLabelBg,
                    bold: false
                )
        );
    }

    private void BuildSectionMetaHeader(
        SheetDiff sheet,
        Grid fieldGrid,
        Grid labelGrid,
        List<string> cols,
        double[] widths
    )
    {
        if (cols.Count == 0)
            return;
        BuildMetaGrid(
            fieldGrid,
            cols,
            widths,
            col => MakeMetaCell(col, MetaFieldFg, MetaFieldBg, bold: true)
        );
        BuildMetaGrid(
            labelGrid,
            cols,
            widths,
            col =>
                MakeMetaCell(
                    sheet.LabelRow.TryGetValue(col, out var v) ? v : string.Empty,
                    MetaLabelFg,
                    MetaLabelBg,
                    bold: false
                )
        );
    }

    private static void BuildMetaGrid(
        Grid grid,
        List<string> cols,
        double[] widths,
        Func<string, TextBlock> makeCell
    )
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

    private static bool IsRemarkCol(string col)
    {
        var lo = col.ToLowerInvariant();
        return lo.Contains("备注")
            || lo.Contains("remark")
            || lo.Contains("note")
            || lo.Contains("desc");
    }

    private static double[] ComputeSheetColWidths(
        List<string> cols,
        IEnumerable<RowConflict> rows,
        HashSet<string> conflictCols
    )
    {
        const double min = 40,
            maxConflict = 280,
            maxRemark = 160,
            maxNormal = 60,
            maxRow = 120,
            pad = 16,
            charPx = 7.5;
        double Cap(string col) =>
            conflictCols.Contains(col) ? maxConflict : (IsRemarkCol(col) ? maxRemark : maxNormal);
        double Measure(string t, double cap) =>
            string.IsNullOrEmpty(t) ? min : Math.Max(min, Math.Min(cap, t.Length * charPx + pad));

        var widths = cols.Select(c => Measure(c, Cap(c))).ToArray();
        foreach (var rc in rows.Take(100))
        {
            bool isRowConflict = rc.DiffType is RowDiffType.OnlyOurs or RowDiffType.OnlyTheirs;
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                double cap = conflictCols.Contains(col)
                    ? maxConflict
                    : (isRowConflict ? maxRow : (IsRemarkCol(col) ? maxRemark : maxNormal));
                if (rc.OursFullRow != null && rc.OursFullRow.TryGetValue(col, out var ov))
                    widths[i] = Math.Max(widths[i], Measure(ov?.ToString() ?? string.Empty, cap));
                if (rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var tv))
                    widths[i] = Math.Max(widths[i], Measure(tv?.ToString() ?? string.Empty, cap));
            }
        }
        return widths;
    }

    private static TextBlock MakeMetaCell(
        string text,
        SolidColorBrush fg,
        SolidColorBrush bg,
        bool bold
    ) =>
        new()
        {
            Text = text,
            Foreground = fg,
            Background = bg,
            Padding = new Thickness(5, 2, 5, 2),
            FontSize = 10,
            FontWeight = bold ? FontWeights.Bold : FontWeights.Normal,
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = string.IsNullOrEmpty(text) ? null : text,
        };

    // ── 详情面板 ─────────────────────────────────────────────────────────────

    internal void OnCellSelected(object sender, CellSelectedRoutedEventArgs e) =>
        ShowDetailForRow(e.Row);

    private void ListBox_ShowDetail(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        // 从点击位置向上找最近的 ConflictRowItem，取其 DataContext
        var el = e.OriginalSource as DependencyObject;
        while (el != null)
        {
            if (el is ConflictRowItem item && item.DataContext is RowConflict rc)
            {
                ShowDetailForRow(rc);
                return;
            }
            el = VisualTreeHelper.GetParent(el);
        }
    }

    private static readonly SolidColorBrush DetailFgOurs = new(Color(0xA8, 0xC8, 0xFF));
    private static readonly SolidColorBrush DetailFgTheirs = new(Color(0xA8, 0xFF, 0xCA));

    // 字符级差异用黄色，与主列表红色背景区分
    private static readonly SolidColorBrush DetailFgDiff = new(Color(0xFF, 0xD0, 0x40));
    private static readonly SolidColorBrush DetailBgDiff = new(Color(0x35, 0x2A, 0x00));
    private static readonly SolidColorBrush DetailFgMuted = new(Color(0x55, 0x55, 0x55));
    private static readonly SolidColorBrush DetailBgOurs = new(Color(0x0A, 0x15, 0x25));
    private static readonly SolidColorBrush DetailBgTheirs = new(Color(0x0A, 0x1A, 0x0F));
    private static readonly SolidColorBrush DetailBgCol = new(Color(0x1A, 0x2A, 0x1A));
    private static readonly SolidColorBrush DetailBorder = new(Color(0x33, 0x33, 0x33));
    private static readonly SolidColorBrush DetailBg = new(Color(0x1A, 0x1A, 0x1A));

    private static WpfColor Color(byte r, byte g, byte b) => WpfColor.FromRgb(r, g, b);

    private void BuildColBatchBar(SheetDiff sheet)
    {
        // 与 Modified 行保持一致：隐藏无冲突列时用 VisibleColumns
        var cols = ConflictRowItem.VisibleColumns ?? sheet.AllColumns;
        if (cols.Count == 0 || _currentColWidths.Length == 0 || !cols.Any(_conflictColSet.Contains))
        {
            ColBatchBar.Visibility = Visibility.Collapsed;
            return;
        }

        ColBatchBar.Visibility = Visibility.Visible;

        var grid = new Grid { Height = 30 };
        for (int i = 0; i < cols.Count; i++)
        {
            var w = i < _currentColWidths.Length ? _currentColWidths[i] : 130.0;
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(w) });
        }

        for (int i = 0; i < cols.Count; i++)
        {
            var col = cols[i];
            if (!_conflictColSet.Contains(col))
                continue;

            var colCapture = col;
            var cell = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
            };
            cell.Children.Add(
                MakeColBatchBtn(
                    "↑我",
                    Color(0x1A, 0x3A, 0x6E),
                    () => SetColumnChoiceAndRefresh(colCapture, ConflictChoice.Ours)
                )
            );
            cell.Children.Add(
                MakeColBatchBtn(
                    "↓他",
                    Color(0x1A, 0x5C, 0x3A),
                    () => SetColumnChoiceAndRefresh(colCapture, ConflictChoice.Theirs)
                )
            );
            Grid.SetColumn(cell, i);
            grid.Children.Add(cell);
        }

        ColBatchPanel.Children.Clear();
        ColBatchPanel.Children.Add(grid);
    }

    private void BuildDetailPanel(List<CellConflict> items)
    {
        DetailPanel.Children.Clear();

        foreach (var cell in items)
        {
            var border = new Border
            {
                Margin = new Thickness(0, 2, 0, 2),
                Background = DetailBg,
                BorderBrush = DetailBorder,
                BorderThickness = new Thickness(1),
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
            );
            grid.ColumnDefinitions.Add(
                new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
            );

            // 列名 + 单列批量按钮
            var colPanel = new StackPanel { Background = DetailBgCol };
            colPanel.Children.Add(
                new TextBlock
                {
                    Text = cell.ColName,
                    Foreground = new SolidColorBrush(Color(0x88, 0xFF, 0x88)),
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(6, 4, 6, 2),
                }
            );
            var colBtnRow = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                Margin = new Thickness(4, 0, 4, 4)
            };
            var cellCapture = cell;
            colBtnRow.Children.Add(
                MakeColBatchBtn(
                    "全取我的",
                    Color(0x1A, 0x3A, 0x6E),
                    () =>
                    {
                        cellCapture.Choice = ConflictChoice.Ours;
                    }
                )
            );
            colBtnRow.Children.Add(
                MakeColBatchBtn(
                    "全取他的",
                    Color(0x1A, 0x5C, 0x3A),
                    () =>
                    {
                        cellCapture.Choice = ConflictChoice.Theirs;
                    }
                )
            );
            colPanel.Children.Add(colBtnRow);
            var colBorder = new Border { Child = colPanel };
            Grid.SetColumn(colBorder, 0);
            grid.Children.Add(colBorder);

            // OURS
            var oursBorder = new Border
            {
                Background = DetailBgOurs,
                Padding = new Thickness(6, 4, 6, 4),
                BorderBrush = new SolidColorBrush(Color(0x22, 0x22, 0x22)),
                BorderThickness = new Thickness(1, 0, 1, 0),
            };
            var oursPanel = new StackPanel();
            oursPanel.Children.Add(
                new TextBlock
                {
                    Text = "我的",
                    Foreground = DetailFgMuted,
                    FontSize = 9,
                    Margin = new Thickness(0, 0, 0, 2)
                }
            );
            oursPanel.Children.Add(
                MakeDetailValueBox(cell.OursDisplay, cell.TheirsDisplay, DetailFgOurs, DetailBgOurs)
            );
            oursBorder.Child = oursPanel;
            Grid.SetColumn(oursBorder, 1);
            grid.Children.Add(oursBorder);

            // THEIRS
            var theirsBorder = new Border
            {
                Background = DetailBgTheirs,
                Padding = new Thickness(6, 4, 6, 4)
            };
            var theirsPanel = new StackPanel();
            theirsPanel.Children.Add(
                new TextBlock
                {
                    Text = "他的",
                    Foreground = DetailFgMuted,
                    FontSize = 9,
                    Margin = new Thickness(0, 0, 0, 2)
                }
            );
            theirsPanel.Children.Add(
                MakeDetailValueBox(
                    cell.TheirsDisplay,
                    cell.OursDisplay,
                    DetailFgTheirs,
                    DetailBgTheirs
                )
            );
            theirsBorder.Child = theirsPanel;
            Grid.SetColumn(theirsBorder, 2);
            grid.Children.Add(theirsBorder);

            border.Child = grid;
            DetailPanel.Children.Add(border);
        }
    }

    private static Button MakeColBatchBtn(string label, WpfColor bg, Action onClick)
    {
        var btn = new Button
        {
            Content = label,
            FontSize = 9,
            Padding = new Thickness(5, 2, 5, 2),
            Margin = new Thickness(0, 0, 4, 0),
            Background = new SolidColorBrush(bg),
            Foreground = System.Windows.Media.Brushes.White,
            BorderThickness = new Thickness(0),
            Cursor = System.Windows.Input.Cursors.Hand,
        };
        btn.Click += (_, _) => onClick();
        return btn;
    }

    private void SetColumnChoiceAndRefresh(string colName, ConflictChoice choice)
    {
        if (SheetTabs.SelectedItem is not TabItem tab)
            return;
        var sheetName = tab.Tag?.ToString() ?? string.Empty;
        var sheetDiff = _diff.Sheets.FirstOrDefault(s => s.SheetName == sheetName);
        sheetDiff?.SetColumnChoice(colName, choice);
        RefreshConflictList();
    }

    // 字符级高亮 RichTextBox（只读可选中），差异段黄色加粗直接显示
    private System.Windows.Controls.RichTextBox MakeDetailValueBox(
        string text,
        string other,
        SolidColorBrush fg,
        SolidColorBrush bg
    )
    {
        var diffTb = MakeInlineDiffBlock(text, other, fg);
        var para = new System.Windows.Documents.Paragraph { Margin = new Thickness(0) };
        foreach (var inline in diffTb.Inlines.ToList())
        {
            diffTb.Inlines.Remove(inline);
            para.Inlines.Add(inline);
        }
        // 如果 MakeInlineDiffBlock 走了纯文本路径（tb.Text 非空）
        if (!para.Inlines.Any() && !string.IsNullOrEmpty(diffTb.Text))
            para.Inlines.Add(new System.Windows.Documents.Run(diffTb.Text) { Foreground = fg });

        var doc = new System.Windows.Documents.FlowDocument(para)
        {
            PagePadding = new Thickness(0),
            LineHeight = 16,
            PageWidth = 10000, // 禁止自动换行
        };

        var rtb = new System.Windows.Controls.RichTextBox(doc)
        {
            IsReadOnly = true,
            Background = bg,
            Foreground = fg,
            FontSize = 11,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(0),
            MaxHeight = 36,
            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled,
            HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
            Cursor = System.Windows.Input.Cursors.IBeam,
            ToolTip = string.IsNullOrEmpty(text) ? null : text,
        };
        return rtb;
    }

    // 构建带字符级高亮的 TextBlock：公共前缀/后缀正常色，差异段黄色加粗
    private TextBlock MakeInlineDiffBlock(string text, string other, SolidColorBrush normalFg)
    {
        var tb = new TextBlock
        {
            FontSize = 11,
            TextWrapping = TextWrapping.NoWrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = string.IsNullOrEmpty(text) ? null : text
        };

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
        int diffEnd = text.Length - suffixLen;

        if (diffStart > 0)
            tb.Inlines.Add(
                new System.Windows.Documents.Run(text[..diffStart]) { Foreground = normalFg }
            );

        if (diffStart < diffEnd)
            tb.Inlines.Add(
                new System.Windows.Documents.Run(text[diffStart..diffEnd])
                {
                    Foreground = DetailFgDiff,
                    FontWeight = FontWeights.Bold,
                    Background = DetailBgDiff,
                }
            );

        if (diffEnd < text.Length)
            tb.Inlines.Add(
                new System.Windows.Documents.Run(text[diffEnd..]) { Foreground = normalFg }
            );

        return tb;
    }

    private static (int prefix, int suffix) FindCommonPrefixSuffix(string a, string b)
    {
        int maxPrefix = Math.Min(a.Length, b.Length);
        int prefix = 0;
        while (prefix < maxPrefix && a[prefix] == b[prefix])
            prefix++;

        int maxSuffix = Math.Min(a.Length - prefix, b.Length - prefix);
        int suffix = 0;
        while (suffix < maxSuffix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix])
            suffix++;

        return (prefix, suffix);
    }

    // ── 变更导航 ─────────────────────────────────────────────────────────────

    private void BuildChangeNav(SheetDiff sheet)
    {
        ChangeNavList.Items.Clear();

        var modified = sheet
            .Rows.Where(r => r.DiffType == RowDiffType.Modified && r.Cells.Count > 0)
            .ToList();
        var onlyOurs = sheet.Rows.Where(r => r.DiffType == RowDiffType.OnlyOurs).ToList();
        var onlyTheirs = sheet.Rows.Where(r => r.DiffType == RowDiffType.OnlyTheirs).ToList();

        if (modified.Count > 0)
            AddNavGroup(
                "冲突列",
                "#FFD080",
                "#3A2A00",
                modified,
                RowDiffType.Modified,
                ConflictChoice.Ours
            );
        if (onlyOurs.Count > 0)
            AddNavGroup(
                "仅我有",
                "#FF8888",
                "#3A1A1A",
                onlyOurs,
                RowDiffType.OnlyOurs,
                ConflictChoice.Ours
            );
        if (onlyTheirs.Count > 0)
            AddNavGroup(
                "仅他有",
                "#88FF88",
                "#1A3A1A",
                onlyTheirs,
                RowDiffType.OnlyTheirs,
                ConflictChoice.Theirs
            );
    }

    private void AddNavGroup(
        string label,
        string fgHex,
        string bgHex,
        List<RowConflict> rows,
        RowDiffType type,
        ConflictChoice defaultChoice
    )
    {
        bool collapsed = false;
        var groupItems = new List<ListBoxItem>();

        // ── 组头 ──
        var headerItem = new ListBoxItem
        {
            Padding = new Thickness(0),
            IsEnabled = true,
            Background = System.Windows.Media.Brushes.Transparent
        };
        var fg = new SolidColorBrush(
            (WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(fgHex)
        );
        var bg = new SolidColorBrush(
            (WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(bgHex)
        );

        var headerBorder = new Border { Background = bg, Padding = new Thickness(4, 3, 4, 3) };
        var headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(
            new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }
        );
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var labelPanel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal
        };
        var chevron = new TextBlock
        {
            Text = "▼ ",
            Foreground = fg,
            FontSize = 10,
            VerticalAlignment = VerticalAlignment.Center
        };
        labelPanel.Children.Add(chevron);
        labelPanel.Children.Add(
            new TextBlock
            {
                Text = label,
                Foreground = fg,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            }
        );
        labelPanel.Children.Add(
            new TextBlock
            {
                Text = $" ({rows.Count})",
                Foreground = new SolidColorBrush(WpfColor.FromRgb(0x88, 0x88, 0x88)),
                FontSize = 10,
                VerticalAlignment = VerticalAlignment.Center
            }
        );
        Grid.SetColumn(labelPanel, 0);
        headerGrid.Children.Add(labelPanel);

        // 全选按钮
        var btnPanel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal
        };
        btnPanel.Children.Add(
            MakeColBatchBtn(
                "全取我的",
                Color(0x1A, 0x3A, 0x6E),
                () =>
                {
                    foreach (var r in rows)
                    {
                        r.SetAllCells(ConflictChoice.Ours);
                        r.RowChoice = ConflictChoice.Ours;
                    }
                }
            )
        );
        btnPanel.Children.Add(
            MakeColBatchBtn(
                "全取他的",
                Color(0x1A, 0x5C, 0x3A),
                () =>
                {
                    foreach (var r in rows)
                    {
                        r.SetAllCells(ConflictChoice.Theirs);
                        r.RowChoice = ConflictChoice.Theirs;
                    }
                }
            )
        );
        Grid.SetColumn(btnPanel, 1);
        headerGrid.Children.Add(btnPanel);
        headerBorder.Child = headerGrid;
        headerItem.Content = headerBorder;

        ChangeNavList.Items.Add(headerItem);

        // ── 子项 ──
        foreach (var rc in rows)
        {
            var tooltip = string.IsNullOrEmpty(rc.DisplayName)
                ? rc.RowKey
                : $"{rc.RowKey}  {rc.DisplayName}";
            var item = new ListBoxItem
            {
                Tag = rc,
                Padding = new Thickness(0),
                ToolTip = tooltip
            };
            var panel = new StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                Margin = new Thickness(16, 1, 4, 1)
            };
            panel.Children.Add(
                new TextBlock
                {
                    Text = rc.RowKey,
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 11,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    MaxWidth = 80,
                    VerticalAlignment = VerticalAlignment.Center,
                }
            );
            if (!string.IsNullOrEmpty(rc.DisplayName))
            {
                panel.Children.Add(
                    new TextBlock
                    {
                        Text = $"  {rc.DisplayName}",
                        Foreground = new SolidColorBrush(WpfColor.FromRgb(0xAA, 0xCC, 0xFF)),
                        FontSize = 10,
                        VerticalAlignment = VerticalAlignment.Center,
                        TextTrimming = TextTrimming.CharacterEllipsis,
                        MaxWidth = 140,
                    }
                );
            }
            if (type == RowDiffType.Modified)
            {
                panel.Children.Add(
                    new TextBlock
                    {
                        Text = $" +{rc.Cells.Count}",
                        Foreground = new SolidColorBrush(WpfColor.FromRgb(0xFF, 0xD0, 0x80)),
                        FontSize = 9,
                        VerticalAlignment = VerticalAlignment.Center,
                    }
                );
            }
            item.Content = panel;
            groupItems.Add(item);
            ChangeNavList.Items.Add(item);
        }

        // 折叠/展开：挂在 Border 上（PreviewMouseDown 穿透 ListBoxItem 捕获）
        headerBorder.PreviewMouseLeftButtonDown += (_, e) =>
        {
            // 如果点的是批量按钮区域则不折叠
            if (
                e.OriginalSource is System.Windows.FrameworkElement fe
                && fe.IsDescendantOf(btnPanel)
            )
                return;
            collapsed = !collapsed;
            chevron.Text = collapsed ? "▶ " : "▼ ";
            foreach (var it in groupItems)
                it.Visibility = collapsed ? Visibility.Collapsed : Visibility.Visible;
            e.Handled = true;
        };
    }

    private void ChangeNavList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (ChangeNavList.SelectedItem is not ListBoxItem { Tag: RowConflict rc })
            return;
        ShowDetailForRow(rc);

        // 在上方 ConflictList 中定位到该行
        if (ConflictList.ItemsSource is System.Collections.IEnumerable items)
        {
            foreach (var item in items)
            {
                if (item is RowConflict row && row.RowKey == rc.RowKey)
                {
                    ConflictList.SelectedItem = row;
                    ConflictList.ScrollIntoView(row);
                    break;
                }
            }
        }
    }

    private void ShowDetailForRow(RowConflict rc)
    {
        // 详情面板折叠时自动展开
        if (!_detailExpanded)
        {
            _detailExpanded = true;
            RowDetail.MinHeight = 28;
            RowDetail.Height = new GridLength(Math.Max(_detailExpandedHeight, 120));
            DetailContent.Visibility = Visibility.Visible;
            DetailChevron.Text = "▼ ";
        }

        DetailRowKey.Text = $"ID: {rc.RowKey}  [{rc.DiffTypeBadge}]";
        DetailHint.Text = string.Empty;

        List<CellConflict> items;
        if (rc.DiffType == RowDiffType.Modified)
        {
            items = rc.Cells.ToList();
        }
        else
        {
            // OnlyOurs: 显示我方全部字段（他方为空）; OnlyTheirs: 显示他方全部字段（我方为空）
            // 两侧都显示非空方的值，让用户能看清楚内容
            var src = rc.OursFullRow ?? rc.TheirsFullRow;
            var isOurs = rc.DiffType == RowDiffType.OnlyOurs;
            items = (src?.Keys ?? Enumerable.Empty<string>())
                .Select(col =>
                {
                    src!.TryGetValue(col, out var val);
                    return new CellConflict
                    {
                        ColName = col,
                        OursValue = isOurs ? val : null,
                        TheirsValue = isOurs ? null : val,
                    };
                })
                .Where(c =>
                    !string.IsNullOrEmpty(c.OursValue?.ToString())
                    || !string.IsNullOrEmpty(c.TheirsValue?.ToString())
                )
                .ToList();
        }
        BuildDetailPanel(items);
    }

    // ── 按钮事件 ─────────────────────────────────────────────────────────────

    private void SheetTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressRefresh)
            return;
        ClearRowSelection();
        UpdateSelectionButtons();
        RefreshConflictList();
    }

    private static void ApplyToggleStyle(
        System.Windows.Controls.Primitives.ToggleButton btn,
        bool active,
        WpfColor activeColor
    )
    {
        btn.Background = active
            ? new SolidColorBrush(activeColor)
            : System.Windows.Media.Brushes.Transparent;
        btn.Foreground = active
            ? System.Windows.Media.Brushes.White
            : new SolidColorBrush(Color(0xAA, 0xAA, 0xAA));
    }

    private void HideNoConflictCols_Changed(object sender, RoutedEventArgs e)
    {
        if (_suppressRefresh)
            return;
        RefreshConflictList();
    }

    private void ViewMode_Changed(object sender, RoutedEventArgs e)
    {
        if (_suppressRefresh)
            return;
        if (sender is System.Windows.Controls.Primitives.ToggleButton btn && btn.IsChecked == true)
        {
            _suppressRefresh = true;
            ViewConflictOnly.IsChecked = ReferenceEquals(btn, ViewConflictOnly);
            ViewContext.IsChecked = ReferenceEquals(btn, ViewContext);
            ViewAll.IsChecked = ReferenceEquals(btn, ViewAll);
            ApplyToggleStyle(
                ViewConflictOnly,
                ViewConflictOnly.IsChecked == true,
                Color(0x1A, 0x3A, 0x6E)
            );
            ApplyToggleStyle(ViewContext, ViewContext.IsChecked == true, Color(0x1A, 0x5C, 0x3A));
            ApplyToggleStyle(ViewAll, ViewAll.IsChecked == true, Color(0x3A, 0x2A, 0x00));
            _viewMode =
                ViewConflictOnly.IsChecked == true
                    ? ViewMode.ConflictOnly
                    : ViewAll.IsChecked == true
                        ? ViewMode.All
                        : ViewMode.Context;
            _suppressRefresh = false;
            RefreshConflictList();
        }
    }

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

    // ── 筛选选中 ─────────────────────────────────────────────────────────────

    private void FilterBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
            FilterSelect_Click(sender, e);
    }

    private void FilterSelect_Click(object sender, RoutedEventArgs e)
    {
        var keyword = FilterBox.Text.Trim();
        if (string.IsNullOrEmpty(keyword)) return;

        var items = ConflictList.ItemsSource as System.Collections.IList;
        if (items == null) return;

        ClearRowSelection();

        bool isNumeric = keyword.All(char.IsDigit);

        foreach (var item in items)
        {
            if (item is not RowConflict rc) continue;

            bool match = isNumeric
                ? rc.RowKey.Contains(keyword, StringComparison.Ordinal)
                : rc.RowKey.Contains(keyword, StringComparison.OrdinalIgnoreCase)
                  || rc.DisplayName.Contains(keyword, StringComparison.OrdinalIgnoreCase);

            if (!match) continue;
            rc.IsSelected = true;
            _selectedRows.Add(rc);
        }

        FilterClearBtn.Visibility = Visibility.Visible;
        UpdateSelectionButtons();
    }

    private void FilterClear_Click(object sender, RoutedEventArgs e)
    {
        FilterBox.Text = string.Empty;
        FilterClearBtn.Visibility = Visibility.Collapsed;
        ClearRowSelection();
        UpdateSelectionButtons();
    }

    private void OnRowDeSelected(object sender, RowDeSelectedRoutedEventArgs e)
    {
        var rc = e.Row;
        rc.IsSelected = false;
        _selectedRows.Remove(rc);
        UpdateSelectionButtons();
    }

    // ── 拖拽多选 ─────────────────────────────────────────────────────────────

    private void ConflictList_DragStart(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        // 点击按钮时不触发拖拽（避免干扰 ✕ 剔除按钮及行内其他按钮）
        if (e.OriginalSource is DependencyObject src)
        {
            var hit = src;
            while (hit != null)
            {
                if (hit is System.Windows.Controls.Button) return;
                hit = System.Windows.Media.VisualTreeHelper.GetParent(hit);
                if (hit is System.Windows.Controls.ListBox) break;
            }
        }

        _dragStartIndex = GetRowIndexAtPoint(e.GetPosition(ConflictList));
        if (_dragStartIndex < 0) return;

        // Ctrl+点击：切换单行选中，不清除其他行
        if ((System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) != 0)
        {
            ToggleRowAt(_dragStartIndex);
            UpdateSelectionButtons();
            return;
        }

        _isDragging = true;
        ClearRowSelection();
        SetRangeSelected(_dragStartIndex, _dragStartIndex, true);
        ConflictList.CaptureMouse();
    }

    private void ConflictList_MiddleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.ChangedButton != System.Windows.Input.MouseButton.Middle) return;
        if (e.OriginalSource is DependencyObject src)
        {
            var hit = src;
            while (hit != null)
            {
                if (hit is System.Windows.Controls.Button) return;
                hit = System.Windows.Media.VisualTreeHelper.GetParent(hit);
                if (hit is System.Windows.Controls.ListBox) break;
            }
        }
        var idx = GetRowIndexAtPoint(e.GetPosition(ConflictList));
        if (idx < 0) return;
        ToggleRowAt(idx);
        UpdateSelectionButtons();
        e.Handled = true;
    }

    private void ToggleRowAt(int idx)
    {
        var items = ConflictList.ItemsSource as System.Collections.IList;
        if (items == null || idx >= items.Count) return;
        if (items[idx] is not RowConflict rc) return;
        if (rc.IsSelected)
        {
            rc.IsSelected = false;
            _selectedRows.Remove(rc);
        }
        else
        {
            rc.IsSelected = true;
            _selectedRows.Add(rc);
        }
    }

    private void ConflictList_DragMove(object sender, System.Windows.Input.MouseEventArgs e)
    {
        if (!_isDragging || e.LeftButton != System.Windows.Input.MouseButtonState.Pressed) return;
        var idx = GetRowIndexAtPoint(e.GetPosition(ConflictList));
        if (idx < 0) return;
        ClearRowSelection();
        int from = Math.Min(_dragStartIndex, idx);
        int to   = Math.Max(_dragStartIndex, idx);
        SetRangeSelected(from, to, true);
    }

    private void ConflictList_DragEnd(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (!_isDragging) return;
        _isDragging = false;
        ConflictList.ReleaseMouseCapture();
        UpdateSelectionButtons();
    }

    private int GetRowIndexAtPoint(System.Windows.Point pt)
    {
        // 用 VisualTreeHelper 找到鼠标下的 ListBoxItem，再查索引
        var hit = ConflictList.InputHitTest(pt) as DependencyObject;
        while (hit != null)
        {
            if (hit is ListBoxItem lbi)
            {
                var items = ConflictList.ItemsSource as System.Collections.IList;
                if (items == null) return -1;
                return items.IndexOf(lbi.DataContext);
            }
            hit = System.Windows.Media.VisualTreeHelper.GetParent(hit);
        }
        return -1;
    }

    private void SetRangeSelected(int from, int to, bool selected)
    {
        var items = ConflictList.ItemsSource as System.Collections.IList;
        if (items == null) return;
        _selectedRows.Clear();
        for (int i = from; i <= to && i < items.Count; i++)
        {
            if (items[i] is RowConflict rc)
            {
                rc.IsSelected = selected;
                if (selected) _selectedRows.Add(rc);
            }
        }
    }

    private void ClearRowSelection()
    {
        foreach (var rc in _selectedRows)
            rc.IsSelected = false;
        _selectedRows.Clear();
    }

    private void UpdateSelectionButtons()
    {
        var hasSelection = _selectedRows.Count > 0;
        SelectionActionBar.Visibility = hasSelection ? Visibility.Visible : Visibility.Collapsed;
        SelectionCountText.Text = hasSelection ? $"已选 {_selectedRows.Count} 行" : string.Empty;
    }

    private void SelectedOurs_Click(object sender, RoutedEventArgs e)
    {
        foreach (var rc in _selectedRows)
        {
            rc.SetAllCells(ConflictChoice.Ours);
            rc.RowChoice = ConflictChoice.Ours;
        }
    }

    private void SelectedTheirs_Click(object sender, RoutedEventArgs e)
    {
        foreach (var rc in _selectedRows)
        {
            rc.SetAllCells(ConflictChoice.Theirs);
            rc.RowChoice = ConflictChoice.Theirs;
        }
    }

    private void SelectedClear_Click(object sender, RoutedEventArgs e)
    {
        ClearRowSelection();
        UpdateSelectionButtons();
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        // 检查未处理的冲突行
        var unresolved = _diff.Sheets
            .SelectMany(s => s.Rows)
            .Where(r => r.DiffType != RowDiffType.Same && !r.IsResolved)
            .ToList();

        if (unresolved.Count > 0)
        {
            var names = unresolved
                .Take(5)
                .Select(r => string.IsNullOrEmpty(r.DisplayName) ? r.RowKey : $"{r.RowKey} {r.DisplayName}");
            var preview = string.Join("\n  ", names);
            var more = unresolved.Count > 5 ? $"\n  …共 {unresolved.Count} 行" : string.Empty;

            var result = MessageBox.Show(
                $"以下 {unresolved.Count} 行尚未做出版本选择：\n  {preview}{more}\n\n确定要继续写回吗？（未选择的行将保留[我的]版本）",
                "存在未处理冲突",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning
            );
            if (result != MessageBoxResult.OK)
            {
                // 跳转到第一个未处理行
                ScrollToRow(unresolved[0]);
                return;
            }
        }

        try
        {
            var savePath = _outPath;
            if (_outPath == _diff.OursPath && !_autoGitAdd)
            {
                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "保存合并结果",
                    Filter = "Excel 文件|*.xlsx",
                    FileName = Path.GetFileName(_outPath),
                    InitialDirectory = Path.GetDirectoryName(_outPath)
                };
                if (dlg.ShowDialog() != true)
                    return;
                savePath = dlg.FileName;
            }

            ConflictApplier.Apply(_diff, savePath, _autoGitAdd);

            MessageBox.Show(
                _autoGitAdd ? $"已写回并执行 git add。\n{savePath}" : $"已写回。\n{savePath}",
                "完成",
                MessageBoxButton.OK,
                MessageBoxImage.Information
            );

            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"写回失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ScrollToRow(RowConflict target)
    {
        // 切到对应 sheet tab
        foreach (TabItem tab in SheetTabs.Items)
        {
            if (tab.Tag?.ToString() == target.SheetName)
            {
                SheetTabs.SelectedItem = tab;
                break;
            }
        }

        // 等布局完成后滚动
        Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, (Action)(() =>
        {
            var items = ConflictList.ItemsSource as System.Collections.IList;
            if (items == null) return;
            var idx = items.IndexOf(target);
            if (idx < 0) return;
            ConflictList.ScrollIntoView(target);
            // 高亮未处理行，方便用户找到
            target.IsSelected = true;
            if (!_selectedRows.Contains(target))
                _selectedRows.Add(target);
            UpdateSelectionButtons();
        }));
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void UpdateStats()
    {
        StatRows.Text = _diff.TotalConflictRows.ToString();
        StatCells.Text = _diff.TotalConflictCells.ToString();
    }

    private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape)
        {
            Close();
            e.Handled = true;
            return;
        }

        // ↑↓ 全局驱动 ChangeNavList，跳过组头（IsEnabled=false 的项）
        if (e.Key != System.Windows.Input.Key.Up && e.Key != System.Windows.Input.Key.Down)
            return;
        if (ChangeNavList.Items.Count == 0)
            return;

        // 如果焦点已在 ChangeNavList 里，让它自己处理
        if (ChangeNavList.IsKeyboardFocusWithin)
            return;

        int current = ChangeNavList.SelectedIndex;
        int next = current;
        int delta = e.Key == System.Windows.Input.Key.Down ? 1 : -1;

        for (int i = current + delta; i >= 0 && i < ChangeNavList.Items.Count; i += delta)
        {
            if (ChangeNavList.Items[i] is ListBoxItem { IsEnabled: true, Tag: RowConflict })
            {
                next = i;
                break;
            }
        }

        if (next != current && next >= 0)
        {
            ChangeNavList.SelectedIndex = next;
            ChangeNavList.ScrollIntoView(ChangeNavList.Items[next]);
        }
        e.Handled = true;
    }
}
