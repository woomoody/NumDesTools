using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NumDesTools.ConflictResolver;
using Border = System.Windows.Controls.Border;
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using HAlign = System.Windows.HorizontalAlignment;
using RadioButton = System.Windows.Controls.RadioButton;
using UserControl = System.Windows.Controls.UserControl;
using VAlign = System.Windows.VerticalAlignment;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfColorConverter = System.Windows.Media.ColorConverter;

namespace NumDesTools.UI;

public partial class ConflictRowItem : UserControl
{
    // 过滤后的冲突列宽（Modified行用）
    public static double[]? CurrentSheetColWidths { get; set; }

    // 全量列宽（OnlyOurs/OnlyTheirs行用，不受隐藏无冲突列影响）
    public static double[]? AllSheetColWidths { get; set; }

    // null = 显示全部列；有值 = 仅显示这些列（隐藏无冲突列模式）
    public static List<string>? VisibleColumns { get; set; }

    // 冲突行（Modified）滚动回调
    public static Action<double>? OnConflictScrollOffsetChanged { get; set; }
    public static Action<double>? OnConflictTotalWidthChanged { get; set; }

    // 新增/删除行（OnlyOurs/OnlyTheirs）滚动回调
    public static Action<double>? OnRowsScrollOffsetChanged { get; set; }
    public static Action<double>? OnRowsTotalWidthChanged { get; set; }

    // 兼容旧调用（同时触发两路）
    public static Action<double>? OnScrollOffsetChanged
    {
        set
        {
            OnConflictScrollOffsetChanged = value;
            OnRowsScrollOffsetChanged = value;
        }
    }
    public static Action<double>? OnTotalWidthChanged
    {
        set
        {
            OnConflictTotalWidthChanged = value;
            OnRowsTotalWidthChanged = value;
        }
    }

    // 活跃实例集合，按类型分开（修改行 vs 新增/删除行）
    private static readonly List<ConflictRowItem> _conflictItems = [];
    private static readonly List<ConflictRowItem> _rowItems = [];
    private bool _isModifiedRow;

    // 当前sheet的默认滚动位置
    public static double InitialScrollOffset { get; set; }
    public static double InitialRowsScrollOffset { get; set; }

    public static void SetGlobalScrollOffset(double offset)
    {
        InitialScrollOffset = offset;
        foreach (var item in _conflictItems)
            item.ApplyScrollOffset(offset);
    }

    public static void SetGlobalRowsScrollOffset(double offset)
    {
        InitialRowsScrollOffset = offset;
        foreach (var item in _rowItems)
            item.ApplyScrollOffset(offset);
    }

    private void ApplyScrollOffset(double offset)
    {
        if (OursScroll.ScrollableWidth > 0)
        {
            OursScroll.ScrollToHorizontalOffset(offset);
            TheirsScroll.ScrollToHorizontalOffset(offset);
        }
        else
        {
            // Not yet measured — queue via LayoutUpdated
            void OnLayout(object? s, EventArgs _)
            {
                OursScroll.LayoutUpdated -= OnLayout;
                OursScroll.ScrollToHorizontalOffset(offset);
                TheirsScroll.ScrollToHorizontalOffset(offset);
            }
            OursScroll.LayoutUpdated += OnLayout;
        }
    }

    private static readonly SolidColorBrush BgOursNormal = Brush("#0A1525");
    private static readonly SolidColorBrush BgTheirsNormal = Brush("#0A1A0F");
    private static readonly SolidColorBrush BgDiff = Brush("#5A1A1A");
    private static readonly SolidColorBrush BgOnlyOurs = Brush("#3A1A1A");
    private static readonly SolidColorBrush BgOnlyTheirs = Brush("#1A3A1A");
    private static readonly SolidColorBrush BgChosenOurs = Brush("#0A3A6E");
    private static readonly SolidColorBrush BgChosenTheirs = Brush("#0A4A2A");
    private static readonly SolidColorBrush BgRejected = Brush("#2A2A2A");
    private static readonly SolidColorBrush FgOurs = Brush("#A8C8FF");
    private static readonly SolidColorBrush FgTheirs = Brush("#A8FFCA");
    private static readonly SolidColorBrush FgDiff = Brush("#FF8888");
    private static readonly SolidColorBrush FgRejected = Brush("#555555");

    public static readonly RoutedEvent CellSelectedEvent = EventManager.RegisterRoutedEvent(
        "CellSelected",
        RoutingStrategy.Bubble,
        typeof(CellSelectedEventHandler),
        typeof(ConflictRowItem)
    );

    public static readonly RoutedEvent RowDeSelectedEvent = EventManager.RegisterRoutedEvent(
        "RowDeSelected",
        RoutingStrategy.Bubble,
        typeof(RowDeSelectedEventHandler),
        typeof(ConflictRowItem)
    );

    public static void AddCellSelectedHandler(DependencyObject d, CellSelectedEventHandler h) =>
        (d as UIElement)?.AddHandler(CellSelectedEvent, h);

    public static void RemoveCellSelectedHandler(DependencyObject d, CellSelectedEventHandler h) =>
        (d as UIElement)?.RemoveHandler(CellSelectedEvent, h);

    public static void AddRowDeSelectedHandler(DependencyObject d, RowDeSelectedEventHandler h) =>
        (d as UIElement)?.AddHandler(RowDeSelectedEvent, h);

    public static void RemoveRowDeSelectedHandler(
        DependencyObject d,
        RowDeSelectedEventHandler h
    ) => (d as UIElement)?.RemoveHandler(RowDeSelectedEvent, h);

    public event CellSelectedEventHandler CellSelected
    {
        add => AddHandler(CellSelectedEvent, value);
        remove => RemoveHandler(CellSelectedEvent, value);
    }

    private void DeSelect_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is RowConflict rc)
            RaiseEvent(new RowDeSelectedRoutedEventArgs(RowDeSelectedEvent, this, rc));
        e.Handled = true;
    }

    public ConflictRowItem()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Loaded += (_, _) =>
        {
            var list = _isModifiedRow ? _conflictItems : _rowItems;
            if (!list.Contains(this))
                list.Add(this);
        };
        Unloaded += (_, _) =>
        {
            _conflictItems.Remove(this);
            _rowItems.Remove(this);
            if (_currentRc != null)
                _currentRc.PropertyChanged -= OnRcPropertyChanged;
        };

        // 注册一次，用实例字段 _isModifiedRow 和 _syncGuard 防重入
        OursScroll.ScrollChanged += OnOursScrollChanged;
        TheirsScroll.ScrollChanged += OnTheirsScrollChanged;
    }

    private bool _syncGuard;

    private void OnOursScrollChanged(object s, ScrollChangedEventArgs ev)
    {
        if (ev.HorizontalChange == 0 || _syncGuard)
            return;
        _syncGuard = true;
        TheirsScroll.ScrollToHorizontalOffset(ev.HorizontalOffset);
        _syncGuard = false;
        if (_isModifiedRow)
            OnConflictScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
        else
            OnRowsScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
    }

    private void OnTheirsScrollChanged(object s, ScrollChangedEventArgs ev)
    {
        if (ev.HorizontalChange == 0 || _syncGuard)
            return;
        _syncGuard = true;
        OursScroll.ScrollToHorizontalOffset(ev.HorizontalOffset);
        _syncGuard = false;
        if (_isModifiedRow)
            OnConflictScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
        else
            OnRowsScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
    }

    private void ApplyInitialScrollAfterLayout()
    {
        var offset = _isModifiedRow ? InitialScrollOffset : InitialRowsScrollOffset;
        if (offset > 0)
            ApplyScrollOffset(offset);
    }

    private RowConflict? _currentRc;

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (_currentRc != null)
            _currentRc.PropertyChanged -= OnRcPropertyChanged;

        if (DataContext is not RowConflict rc)
            return;

        _currentRc = rc;
        rc.PropertyChanged += OnRcPropertyChanged;
        Render(rc);
    }

    private void OnRcPropertyChanged(
        object? sender,
        System.ComponentModel.PropertyChangedEventArgs e
    )
    {
        if (e.PropertyName == nameof(RowConflict.IsSelected) && _currentRc != null)
        {
            UpdateSelectionHighlight(_currentRc);
            DeSelectBtn.Visibility = _currentRc.IsSelected
                ? Visibility.Visible
                : Visibility.Collapsed;
        }
    }

    private void Render(RowConflict rc)
    {
        _isModifiedRow =
            rc.DiffType != RowDiffType.OnlyOurs && rc.DiffType != RowDiffType.OnlyTheirs;
        // 更新列表归属（DataContext 复用时 DiffType 可能变化）
        if (IsLoaded)
        {
            _conflictItems.Remove(this);
            _rowItems.Remove(this);
            var list = _isModifiedRow ? _conflictItems : _rowItems;
            if (!list.Contains(this))
                list.Add(this);
        }
        RowKeyText.Text = rc.RowKey;
        DiffTypeBadge.Text = rc.DiffTypeBadge;

        // # 备注列分段显示
        RowHashCols.Children.Clear();
        var hashVals = rc.HashColValues;
        for (int hi = 0; hi < hashVals.Count; hi++)
        {
            if (hi > 0)
                RowHashCols.Children.Add(
                    new TextBlock
                    {
                        Text = " | ",
                        Foreground = Brush("#555555"),
                        FontSize = 11,
                        VerticalAlignment = VerticalAlignment.Center
                    }
                );
            RowHashCols.Children.Add(
                new TextBlock
                {
                    Text = hashVals[hi].Val,
                    ToolTip = $"{hashVals[hi].Col}: {hashVals[hi].Val}",
                    Foreground = Brush("#AACCFF"),
                    FontSize = 11,
                    MaxWidth = 200,
                    TextWrapping = TextWrapping.NoWrap,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                }
            );
        }

        var isModified = rc.DiffType == RowDiffType.Modified;
        var isOnlyOurs = rc.DiffType == RowDiffType.OnlyOurs;
        var isOnlyTheirs = rc.DiffType == RowDiffType.OnlyTheirs;

        UpdateSelectionHighlight(rc);
        DeSelectBtn.Visibility = rc.IsSelected ? Visibility.Visible : Visibility.Collapsed;

        BadgeBorder.Background = isModified
            ? Brush("#5A4A00")
            : isOnlyOurs
                ? Brush("#5A1A1A")
                : Brush("#1A5A2A");
        DiffTypeBadge.Foreground = isModified
            ? Brush("#FFD080")
            : isOnlyOurs
                ? Brush("#FF8888")
                : Brush("#88FF88");

        BatchButtons.Visibility = isModified ? Visibility.Visible : Visibility.Collapsed;
        RowChoicePanel.Visibility =
            isOnlyOurs || isOnlyTheirs ? Visibility.Visible : Visibility.Collapsed;

        if (isOnlyOurs || isOnlyTheirs)
        {
            var bindOurs = new System.Windows.Data.Binding(nameof(RowConflict.RowChoiceOurs))
            {
                Source = rc,
                Mode = System.Windows.Data.BindingMode.TwoWay
            };
            var bindTheirs = new System.Windows.Data.Binding(nameof(RowConflict.RowChoiceTheirs))
            {
                Source = rc,
                Mode = System.Windows.Data.BindingMode.TwoWay
            };
            RowChoiceOursRb.SetBinding(CheckBox.IsCheckedProperty, bindOurs);
            RowChoiceTheirsRb.SetBinding(CheckBox.IsCheckedProperty, bindTheirs);
        }

        var allCols = rc.AllColumns.Count > 0 ? rc.AllColumns : DeriveColumns(rc);
        List<string> cols;
        double[] colWidths;

        if (isOnlyOurs || isOnlyTheirs)
        {
            // 新增/删除行：始终显示全量列，用全量列宽
            cols = allCols;
            var widths = AllSheetColWidths ?? CurrentSheetColWidths;
            colWidths =
                (widths != null && widths.Length >= cols.Count)
                    ? widths
                    : FallbackColWidths(cols, rc);
        }
        else
        {
            // Modified / Same 行：受隐藏无冲突列影响，与列头对齐
            if (VisibleColumns != null && CurrentSheetColWidths != null)
            {
                cols = VisibleColumns;
                colWidths = CurrentSheetColWidths;
            }
            else
            {
                cols = allCols;
                colWidths = CurrentSheetColWidths ?? FallbackColWidths(cols, rc);
            }
        }

        if (cols.Count == 0)
        {
            OursScroll.Visibility = TheirsScroll.Visibility = Visibility.Collapsed;
            return;
        }
        var diffCols = new HashSet<string>(rc.Cells.Select(c => c.ColName), StringComparer.Ordinal);

        // OURS 行
        OursScroll.Visibility = Visibility.Visible;
        BuildGridColumns(OursGrid, colWidths);
        OursGrid.Children.Clear();
        var rowBgOurs = isOnlyOurs ? BgOnlyOurs : BgOursNormal;
        for (int i = 0; i < cols.Count; i++)
        {
            var col = cols[i];
            var val =
                rc.OursFullRow != null && rc.OursFullRow.TryGetValue(col, out var v)
                    ? v?.ToString() ?? ""
                    : "";
            var diff = diffCols.Contains(col);
            if (isModified && diff)
            {
                var cell = rc.Cells.FirstOrDefault(c => c.ColName == col);
                if (cell != null)
                {
                    var container = MakeCellContainer(cell, isOurs: true, val, colWidths[i], rc);
                    Grid.SetColumn(container, i);
                    OursGrid.Children.Add(container);
                    continue;
                }
            }
            var tb = MakeCell(val, diff ? FgDiff : FgOurs, diff ? BgDiff : rowBgOurs, colWidths[i]);
            tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
            Grid.SetColumn(tb, i);
            OursGrid.Children.Add(tb);
        }

        // THEIRS 行
        if (isOnlyOurs)
        {
            TheirsScroll.Visibility = Visibility.Collapsed;
        }
        else
        {
            TheirsScroll.Visibility = Visibility.Visible;
            BuildGridColumns(TheirsGrid, colWidths);
            TheirsGrid.Children.Clear();
            var rowBgTheirs = isOnlyTheirs ? BgOnlyTheirs : BgTheirsNormal;
            for (int i = 0; i < cols.Count; i++)
            {
                var col = cols[i];
                var val =
                    rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var v)
                        ? v?.ToString() ?? ""
                        : "";
                var diff = diffCols.Contains(col);
                if (isModified && diff)
                {
                    var cell = rc.Cells.FirstOrDefault(c => c.ColName == col);
                    if (cell != null)
                    {
                        var container = MakeCellContainer(
                            cell,
                            isOurs: false,
                            val,
                            colWidths[i],
                            rc
                        );
                        Grid.SetColumn(container, i);
                        TheirsGrid.Children.Add(container);
                        continue;
                    }
                }
                var tb = MakeCell(
                    val,
                    diff ? FgDiff : FgTheirs,
                    diff ? BgDiff : rowBgTheirs,
                    colWidths[i]
                );
                tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
                Grid.SetColumn(tb, i);
                TheirsGrid.Children.Add(tb);
            }
        }

        ApplyInitialScrollAfterLayout();
    }

    private void RaiseDetailEvent(RowConflict rc) =>
        RaiseEvent(new CellSelectedRoutedEventArgs(CellSelectedEvent, this, rc));

    // 兜底：单行自测量（CurrentSheetColWidths 未设置时）
    private static double[] FallbackColWidths(List<string> cols, RowConflict rc)
    {
        const double min = 60,
            max = 220,
            pad = 12,
            fs = 11.0;
        var tf = new Typeface("Consolas");
        double M(string t)
        {
            if (string.IsNullOrEmpty(t))
                return min;
            var ft = new FormattedText(
                t,
                System.Globalization.CultureInfo.CurrentCulture,
                System.Windows.FlowDirection.LeftToRight,
                tf,
                fs,
                WpfBrushes.White,
                1.0
            );
            return Math.Max(min, Math.Min(max, ft.Width + pad));
        }
        return cols.Select(col =>
            {
                var w = M(col);
                if (rc.OursFullRow != null && rc.OursFullRow.TryGetValue(col, out var ov))
                    w = Math.Max(w, M(ov?.ToString() ?? ""));
                if (rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var tv))
                    w = Math.Max(w, M(tv?.ToString() ?? ""));
                return w;
            })
            .ToArray();
    }

    private static void BuildGridColumns(Grid g, double[] widths)
    {
        g.ColumnDefinitions.Clear();
        foreach (var w in widths)
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(w) });
    }

    private static TextBlock MakeCell(
        string text,
        SolidColorBrush fg,
        SolidColorBrush bg,
        double colWidth
    ) =>
        new()
        {
            Text = text,
            Foreground = fg,
            Background = bg,
            Padding = new Thickness(5, 3, 5, 3),
            FontSize = 11,
            TextWrapping = TextWrapping.NoWrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            ToolTip = string.IsNullOrEmpty(text) ? null : text,
            VerticalAlignment = VerticalAlignment.Center,
            Width = colWidth,
            Height = 24,
            Cursor = System.Windows.Input.Cursors.Hand,
        };

    // 冲突格容器：未选时显示选择按钮，已选后背景变色并显示×撤销按钮
    // 点击文本区域或×按钮均触发详情面板
    private Grid MakeCellContainer(
        CellConflict cell,
        bool isOurs,
        string val,
        double colWidth,
        RowConflict rc
    )
    {
        var container = new Grid { Width = colWidth };

        void Refresh()
        {
            container.Children.Clear();
            bool chosen = cell.IsExplicit;
            bool thisWon =
                chosen
                && (
                    isOurs
                        ? cell.Choice == ConflictChoice.Ours
                        : cell.Choice == ConflictChoice.Theirs
                );
            bool otherWon = chosen && !thisWon;

            var bgBrush = thisWon
                ? (isOurs ? BgChosenOurs : BgChosenTheirs)
                : otherWon
                    ? BgRejected
                    : BgDiff;
            var fgBrush = otherWon ? FgRejected : (isOurs ? FgOurs : FgTheirs);

            var tb = new TextBlock
            {
                Text = val,
                Foreground = fgBrush,
                Background = bgBrush,
                Padding = new Thickness(5, 3, 5, 3),
                FontSize = 11,
                TextWrapping = TextWrapping.NoWrap,
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = string.IsNullOrEmpty(val) ? null : val,
                VerticalAlignment = VerticalAlignment.Center,
                Width = colWidth,
                Height = 24,
                Cursor = System.Windows.Input.Cursors.Hand,
            };
            tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
            container.Children.Add(tb);

            if (!chosen)
            {
                var btn = new Button
                {
                    Content = isOurs ? "我的" : "他的",
                    Foreground = isOurs ? FgOurs : FgTheirs,
                    FontSize = 9,
                    HorizontalAlignment = HAlign.Right,
                    VerticalAlignment = VAlign.Bottom,
                    Margin = new Thickness(0, 0, 2, 1),
                    Background = WpfBrushes.Transparent,
                    BorderBrush = isOurs ? FgOurs : FgTheirs,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(2, 0, 2, 0),
                    Cursor = System.Windows.Input.Cursors.Hand,
                };
                btn.Click += (_, _) =>
                {
                    cell.Choice = isOurs ? ConflictChoice.Ours : ConflictChoice.Theirs;
                    cell.IsExplicit = true;
                    RaiseDetailEvent(rc);
                };
                container.Children.Add(btn);
            }
            else if (thisWon)
            {
                var xBtn = new Button
                {
                    Content = "×",
                    Foreground = WpfBrushes.White,
                    FontSize = 9,
                    HorizontalAlignment = HAlign.Right,
                    VerticalAlignment = VAlign.Top,
                    Margin = new Thickness(0, 1, 2, 0),
                    Background = WpfBrushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Padding = new Thickness(2, 0, 2, 0),
                    Cursor = System.Windows.Input.Cursors.Hand,
                };
                xBtn.Click += (_, _) =>
                {
                    cell.ClearChoice();
                    RaiseDetailEvent(rc);
                };
                container.Children.Add(xBtn);
            }
        }

        Refresh();
        cell.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName is nameof(CellConflict.IsExplicit) or nameof(CellConflict.Choice))
                Refresh();
        };
        return container;
    }

    private static List<string> DeriveColumns(RowConflict rc)
    {
        var set = new LinkedList<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        void Add(string c)
        {
            if (seen.Add(c))
                set.AddLast(c);
        }
        if (rc.OursFullRow != null)
            foreach (var k in rc.OursFullRow.Keys)
                Add(k);
        if (rc.TheirsFullRow != null)
            foreach (var k in rc.TheirsFullRow.Keys)
                Add(k);
        foreach (var c in rc.Cells)
            Add(c.ColName);
        return set.ToList();
    }

    private static readonly SolidColorBrush BgSelected =
        new(WpfColor.FromArgb(180, 0x3A, 0x60, 0xA0));

    private void UpdateSelectionHighlight(RowConflict rc)
    {
        if (rc.IsSelected)
        {
            HeaderGrid.Background = BgSelected;
        }
        else
        {
            HeaderGrid.Background =
                rc.DiffType == RowDiffType.OnlyOurs
                    ? BgOnlyOurs
                    : rc.DiffType == RowDiffType.OnlyTheirs
                        ? BgOnlyTheirs
                        : Brush("#2A2A2A");
        }
    }

    private void BatchChoice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && DataContext is RowConflict rc)
        {
            var choice =
                btn.Tag?.ToString() == "Theirs" ? ConflictChoice.Theirs : ConflictChoice.Ours;
            rc.SetAllCells(choice);
        }
    }

    private static SolidColorBrush Brush(string hex)
    {
        var c = (WpfColor)WpfColorConverter.ConvertFromString(hex);
        return new SolidColorBrush(c);
    }
}

public delegate void RowDeSelectedEventHandler(object sender, RowDeSelectedRoutedEventArgs e);

public delegate void CellSelectedEventHandler(object sender, CellSelectedRoutedEventArgs e);

public class CellSelectedRoutedEventArgs(RoutedEvent routedEvent, object source, RowConflict row)
    : RoutedEventArgs(routedEvent, source)
{
    public RowConflict Row { get; } = row;
}

public class RowDeSelectedRoutedEventArgs(RoutedEvent routedEvent, object source, RowConflict row)
    : RoutedEventArgs(routedEvent, source)
{
    public RowConflict Row { get; } = row;
}
