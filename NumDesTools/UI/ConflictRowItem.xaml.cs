using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NumDesTools.ConflictResolver;
using Border           = System.Windows.Controls.Border;
using Button           = System.Windows.Controls.Button;
using CheckBox         = System.Windows.Controls.CheckBox;
using RadioButton      = System.Windows.Controls.RadioButton;
using UserControl      = System.Windows.Controls.UserControl;
using WpfBrushes       = System.Windows.Media.Brushes;
using WpfColor         = System.Windows.Media.Color;
using WpfColorConverter = System.Windows.Media.ColorConverter;
using HAlign           = System.Windows.HorizontalAlignment;
using VAlign           = System.Windows.VerticalAlignment;

namespace NumDesTools.UI;

public partial class ConflictRowItem : UserControl
{
    // 由 ExcelConflictWindow 在刷新列表前设置，所有行共享同一套列宽
    public static double[]? CurrentSheetColWidths { get; set; }

    // 主窗口注入：行内滚动时同步 MetaScroll + SharedHScrollBar
    public static Action<double>? OnScrollOffsetChanged { get; set; }

    // 主窗口注入：列宽总量变化时更新共享滚动条 Maximum
    public static Action<double>? OnTotalWidthChanged { get; set; }

    // 活跃实例集合，供主窗口共享滚动条驱动所有行
    private static readonly List<ConflictRowItem> _activeItems = [];

    // 当前sheet的默认滚动位置，新加载的行自动定位到冲突列
    public static double InitialScrollOffset { get; set; }

    public static void SetGlobalScrollOffset(double offset)
    {
        InitialScrollOffset = offset;
        foreach (var item in _activeItems)
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

    private static readonly SolidColorBrush BgOursNormal   = Brush("#0A1525");
    private static readonly SolidColorBrush BgTheirsNormal = Brush("#0A1A0F");
    private static readonly SolidColorBrush BgDiff         = Brush("#5A1A1A");
    private static readonly SolidColorBrush BgOnlyOurs     = Brush("#3A1A1A");
    private static readonly SolidColorBrush BgOnlyTheirs   = Brush("#1A3A1A");
    private static readonly SolidColorBrush BgChosenOurs   = Brush("#0A3A6E");
    private static readonly SolidColorBrush BgChosenTheirs = Brush("#0A4A2A");
    private static readonly SolidColorBrush BgRejected     = Brush("#2A2A2A");
    private static readonly SolidColorBrush FgOurs         = Brush("#A8C8FF");
    private static readonly SolidColorBrush FgTheirs       = Brush("#A8FFCA");
    private static readonly SolidColorBrush FgDiff         = Brush("#FF8888");
    private static readonly SolidColorBrush FgRejected     = Brush("#555555");

    public static readonly RoutedEvent CellSelectedEvent =
        EventManager.RegisterRoutedEvent("CellSelected", RoutingStrategy.Bubble,
            typeof(CellSelectedEventHandler), typeof(ConflictRowItem));

    public event CellSelectedEventHandler CellSelected
    {
        add    => AddHandler(CellSelectedEvent, value);
        remove => RemoveHandler(CellSelectedEvent, value);
    }

    public ConflictRowItem()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Loaded   += (_, _) => { if (!_activeItems.Contains(this)) _activeItems.Add(this); };
        Unloaded += (_, _) => _activeItems.Remove(this);
    }

    private void ApplyInitialScrollAfterLayout()
    {
        if (InitialScrollOffset > 0)
            ApplyScrollOffset(InitialScrollOffset);
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (DataContext is not RowConflict rc) return;
        Render(rc);
    }

    private void Render(RowConflict rc)
    {
        RowKeyText.Text = rc.RowKey;
        DiffTypeBadge.Text = rc.DiffTypeBadge;

        // # 备注列分段显示
        RowHashCols.Children.Clear();
        var hashVals = rc.HashColValues;
        for (int hi = 0; hi < hashVals.Count; hi++)
        {
            if (hi > 0)
                RowHashCols.Children.Add(new TextBlock { Text = " | ", Foreground = Brush("#555555"), FontSize = 11, VerticalAlignment = VerticalAlignment.Center });
            RowHashCols.Children.Add(new TextBlock
            {
                Text = hashVals[hi].Val,
                ToolTip = $"{hashVals[hi].Col}: {hashVals[hi].Val}",
                Foreground = Brush("#AACCFF"), FontSize = 11,
                MaxWidth = 200, TextWrapping = TextWrapping.NoWrap,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
            });
        }

        var isModified   = rc.DiffType == RowDiffType.Modified;
        var isOnlyOurs   = rc.DiffType == RowDiffType.OnlyOurs;
        var isOnlyTheirs = rc.DiffType == RowDiffType.OnlyTheirs;

        HeaderGrid.Background = isOnlyOurs   ? BgOnlyOurs
                              : isOnlyTheirs ? BgOnlyTheirs
                              :                Brush("#2A2A2A");
        BadgeBorder.Background = isModified   ? Brush("#5A4A00")
                               : isOnlyOurs   ? Brush("#5A1A1A")
                               :                Brush("#1A5A2A");
        DiffTypeBadge.Foreground = isModified   ? Brush("#FFD080")
                                 : isOnlyOurs   ? Brush("#FF8888")
                                 :                Brush("#88FF88");

        BatchButtons.Visibility   = isModified                  ? Visibility.Visible : Visibility.Collapsed;
        RowChoicePanel.Visibility = isOnlyOurs || isOnlyTheirs ? Visibility.Visible : Visibility.Collapsed;

        if (isOnlyOurs || isOnlyTheirs)
        {
            var bindOurs = new System.Windows.Data.Binding(nameof(RowConflict.RowChoiceOurs))
                { Source = rc, Mode = System.Windows.Data.BindingMode.TwoWay };
            var bindTheirs = new System.Windows.Data.Binding(nameof(RowConflict.RowChoiceTheirs))
                { Source = rc, Mode = System.Windows.Data.BindingMode.TwoWay };
            RowChoiceOursRb.SetBinding(CheckBox.IsCheckedProperty,   bindOurs);
            RowChoiceTheirsRb.SetBinding(CheckBox.IsCheckedProperty, bindTheirs);
        }

        var cols = rc.AllColumns.Count > 0 ? rc.AllColumns : DeriveColumns(rc);
        if (cols.Count == 0)
        {
            OursScroll.Visibility = TheirsScroll.Visibility = Visibility.Collapsed;
            return;
        }

        var sharedWidths = CurrentSheetColWidths;
        var colWidths = (sharedWidths != null && sharedWidths.Length >= cols.Count)
            ? sharedWidths
            : FallbackColWidths(cols, rc);
        var diffCols = new HashSet<string>(rc.Cells.Select(c => c.ColName), StringComparer.Ordinal);

        OnTotalWidthChanged?.Invoke(colWidths.Sum());
        SyncScroll(OursScroll, TheirsScroll);

        // OURS 行
        OursScroll.Visibility = Visibility.Visible;
        BuildGridColumns(OursGrid, colWidths);
        OursGrid.Children.Clear();
        var rowBgOurs = isOnlyOurs ? BgOnlyOurs : BgOursNormal;
        for (int i = 0; i < cols.Count; i++)
        {
            var col  = cols[i];
            var val  = rc.OursFullRow != null && rc.OursFullRow.TryGetValue(col, out var v) ? v?.ToString() ?? "" : "";
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
                var col  = cols[i];
                var val  = rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var v) ? v?.ToString() ?? "" : "";
                var diff = diffCols.Contains(col);
                if (isModified && diff)
                {
                    var cell = rc.Cells.FirstOrDefault(c => c.ColName == col);
                    if (cell != null)
                    {
                        var container = MakeCellContainer(cell, isOurs: false, val, colWidths[i], rc);
                        Grid.SetColumn(container, i);
                        TheirsGrid.Children.Add(container);
                        continue;
                    }
                }
                var tb = MakeCell(val, diff ? FgDiff : FgTheirs, diff ? BgDiff : rowBgTheirs, colWidths[i]);
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
        const double min = 60, max = 220, pad = 12, fs = 11.0;
        var tf = new Typeface("Consolas");
        double M(string t)
        {
            if (string.IsNullOrEmpty(t)) return min;
            var ft = new FormattedText(t, System.Globalization.CultureInfo.CurrentCulture,
                System.Windows.FlowDirection.LeftToRight, tf, fs, WpfBrushes.White, 1.0);
            return Math.Max(min, Math.Min(max, ft.Width + pad));
        }
        return cols.Select(col =>
        {
            var w = M(col);
            if (rc.OursFullRow   != null && rc.OursFullRow.TryGetValue(col,   out var ov)) w = Math.Max(w, M(ov?.ToString() ?? ""));
            if (rc.TheirsFullRow != null && rc.TheirsFullRow.TryGetValue(col, out var tv)) w = Math.Max(w, M(tv?.ToString() ?? ""));
            return w;
        }).ToArray();
    }

    private static void BuildGridColumns(Grid g, double[] widths)
    {
        g.ColumnDefinitions.Clear();
        foreach (var w in widths)
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(w) });
    }

    private static TextBlock MakeCell(string text, SolidColorBrush fg, SolidColorBrush bg, double colWidth)
        => new()
        {
            Text              = text,
            Foreground        = fg,
            Background        = bg,
            Padding           = new Thickness(5, 3, 5, 3),
            FontSize          = 11,
            TextWrapping      = TextWrapping.NoWrap,
            TextTrimming      = TextTrimming.CharacterEllipsis,
            ToolTip           = string.IsNullOrEmpty(text) ? null : text,
            VerticalAlignment = VerticalAlignment.Center,
            Width             = colWidth,
            Height            = 24,
            Cursor            = System.Windows.Input.Cursors.Hand,
        };

    // 冲突格容器：未选时显示选择按钮，已选后背景变色并显示×撤销按钮
    // 点击文本区域或×按钮均触发详情面板
    private Grid MakeCellContainer(CellConflict cell, bool isOurs, string val, double colWidth, RowConflict rc)
    {
        var container = new Grid { Width = colWidth };

        void Refresh()
        {
            container.Children.Clear();
            bool chosen   = cell.IsExplicit;
            bool thisWon  = chosen && (isOurs ? cell.Choice == ConflictChoice.Ours : cell.Choice == ConflictChoice.Theirs);
            bool otherWon = chosen && !thisWon;

            var bgBrush = thisWon  ? (isOurs ? BgChosenOurs : BgChosenTheirs)
                        : otherWon ? BgRejected
                        :            BgDiff;
            var fgBrush = otherWon ? FgRejected : (isOurs ? FgOurs : FgTheirs);

            var tb = new TextBlock
            {
                Text              = val,
                Foreground        = fgBrush,
                Background        = bgBrush,
                Padding           = new Thickness(5, 3, 5, 3),
                FontSize          = 11,
                TextWrapping      = TextWrapping.NoWrap,
                TextTrimming      = TextTrimming.CharacterEllipsis,
                ToolTip           = string.IsNullOrEmpty(val) ? null : val,
                VerticalAlignment = VerticalAlignment.Center,
                Width             = colWidth,
                Height            = 24,
                Cursor            = System.Windows.Input.Cursors.Hand,
            };
            tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
            container.Children.Add(tb);

            if (!chosen)
            {
                var btn = new Button
                {
                    Content             = isOurs ? "我的" : "他的",
                    Foreground          = isOurs ? FgOurs : FgTheirs,
                    FontSize            = 9,
                    HorizontalAlignment = HAlign.Right,
                    VerticalAlignment   = VAlign.Bottom,
                    Margin              = new Thickness(0, 0, 2, 1),
                    Background          = WpfBrushes.Transparent,
                    BorderBrush         = isOurs ? FgOurs : FgTheirs,
                    BorderThickness     = new Thickness(1),
                    Padding             = new Thickness(2, 0, 2, 0),
                    Cursor              = System.Windows.Input.Cursors.Hand,
                };
                btn.Click += (_, _) =>
                {
                    cell.Choice     = isOurs ? ConflictChoice.Ours : ConflictChoice.Theirs;
                    cell.IsExplicit = true;
                    RaiseDetailEvent(rc);
                };
                container.Children.Add(btn);
            }
            else if (thisWon)
            {
                var xBtn = new Button
                {
                    Content             = "×",
                    Foreground          = WpfBrushes.White,
                    FontSize            = 9,
                    HorizontalAlignment = HAlign.Right,
                    VerticalAlignment   = VAlign.Top,
                    Margin              = new Thickness(0, 1, 2, 0),
                    Background          = WpfBrushes.Transparent,
                    BorderThickness     = new Thickness(0),
                    Padding             = new Thickness(2, 0, 2, 0),
                    Cursor              = System.Windows.Input.Cursors.Hand,
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

    private static void SyncScroll(ScrollViewer ours, ScrollViewer theirs)
    {
        theirs.ScrollChanged -= OnTheirsScrollChanged;
        theirs.ScrollChanged += OnTheirsScrollChanged;
        ours.ScrollChanged   -= OnOursScrollChanged;
        ours.ScrollChanged   += OnOursScrollChanged;

        void OnTheirsScrollChanged(object s, ScrollChangedEventArgs ev)
        {
            ours.ScrollToHorizontalOffset(ev.HorizontalOffset);
            OnScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
        }
        void OnOursScrollChanged(object s, ScrollChangedEventArgs ev)
        {
            theirs.ScrollToHorizontalOffset(ev.HorizontalOffset);
            OnScrollOffsetChanged?.Invoke(ev.HorizontalOffset);
        }
    }

    private static List<string> DeriveColumns(RowConflict rc)
    {
        var set  = new LinkedList<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        void Add(string c) { if (seen.Add(c)) set.AddLast(c); }
        if (rc.OursFullRow   != null) foreach (var k in rc.OursFullRow.Keys)   Add(k);
        if (rc.TheirsFullRow != null) foreach (var k in rc.TheirsFullRow.Keys) Add(k);
        foreach (var c in rc.Cells) Add(c.ColName);
        return set.ToList();
    }

    private void BatchChoice_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && DataContext is RowConflict rc)
        {
            var choice = btn.Tag?.ToString() == "Theirs" ? ConflictChoice.Theirs : ConflictChoice.Ours;
            rc.SetAllCells(choice);
        }
    }

    private static SolidColorBrush Brush(string hex)
    {
        var c = (WpfColor)WpfColorConverter.ConvertFromString(hex);
        return new SolidColorBrush(c);
    }
}

public delegate void CellSelectedEventHandler(object sender, CellSelectedRoutedEventArgs e);

public class CellSelectedRoutedEventArgs(RoutedEvent routedEvent, object source, RowConflict row)
    : RoutedEventArgs(routedEvent, source)
{
    public RowConflict Row { get; } = row;
}
