using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using NumDesTools.ConflictResolver;
using Button           = System.Windows.Controls.Button;
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

    // 主窗口注入：TheirsScroll 滚动时同步 MetaScroll
    public static Action<double>? OnScrollOffsetChanged { get; set; }

    private static readonly SolidColorBrush BgOursNormal   = Brush("#0A1525");
    private static readonly SolidColorBrush BgTheirsNormal = Brush("#0A1A0F");
    private static readonly SolidColorBrush BgDiff         = Brush("#5A1A1A");
    private static readonly SolidColorBrush BgOnlyOurs     = Brush("#3A1A1A");
    private static readonly SolidColorBrush BgOnlyTheirs   = Brush("#1A3A1A");
    private static readonly SolidColorBrush FgOurs         = Brush("#A8C8FF");
    private static readonly SolidColorBrush FgTheirs       = Brush("#A8FFCA");
    private static readonly SolidColorBrush FgDiff         = Brush("#FF8888");

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
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (DataContext is not RowConflict rc) return;
        Render(rc);
    }

    private void Render(RowConflict rc)
    {
        RowKeyText.Text    = rc.RowKey;
        DiffTypeBadge.Text = rc.DiffTypeBadge;

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

        BatchButtons.Visibility   = isModified ? Visibility.Visible   : Visibility.Collapsed;
        RowChoicePanel.Visibility = isModified ? Visibility.Collapsed : Visibility.Visible;

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
        var diffCols  = new HashSet<string>(rc.Cells.Select(c => c.ColName), StringComparer.Ordinal);

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
            var tb   = MakeCell(val, diff ? FgDiff : FgOurs, diff ? BgDiff : rowBgOurs, colWidths[i]);
            tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
            Grid.SetColumn(tb, i);
            OursGrid.Children.Add(tb);

            if (isModified && diff)
            {
                var cell = rc.Cells.FirstOrDefault(c => c.ColName == col);
                if (cell != null) OursGrid.Children.Add(MakeCellRadio(cell, isOurs: true, i));
            }
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
                var tb   = MakeCell(val, diff ? FgDiff : FgTheirs, diff ? BgDiff : rowBgTheirs, colWidths[i]);
                tb.MouseLeftButtonDown += (_, _) => RaiseDetailEvent(rc);
                Grid.SetColumn(tb, i);
                TheirsGrid.Children.Add(tb);

                if (isModified && diff)
                {
                    var cell = rc.Cells.FirstOrDefault(c => c.ColName == col);
                    if (cell != null) TheirsGrid.Children.Add(MakeCellRadio(cell, isOurs: false, i));
                }
            }
        }
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
            TextTrimming      = TextTrimming.CharacterEllipsis,
            ToolTip           = string.IsNullOrEmpty(text) ? null : text,
            VerticalAlignment = VerticalAlignment.Stretch,
            Width             = colWidth,
            Cursor            = System.Windows.Input.Cursors.Hand,
        };

    private static RadioButton MakeCellRadio(CellConflict cell, bool isOurs, int colIdx)
    {
        var rb = new RadioButton
        {
            GroupName           = $"cell_{cell.ColName}_{(isOurs ? "o" : "t")}",
            Content             = isOurs ? "我的" : "他的",
            Foreground          = isOurs ? FgOurs : FgTheirs,
            FontSize            = 9,
            HorizontalAlignment = HAlign.Right,
            VerticalAlignment   = VAlign.Bottom,
            Margin              = new Thickness(0, 0, 2, 1),
            Background          = WpfBrushes.Transparent,
            BorderThickness     = new Thickness(0),
            Padding             = new Thickness(2, 0, 2, 0),
        };
        var binding = new System.Windows.Data.Binding(isOurs ? nameof(CellConflict.ChoiceOurs) : nameof(CellConflict.ChoiceTheirs))
            { Source = cell, Mode = System.Windows.Data.BindingMode.TwoWay };
        rb.SetBinding(RadioButton.IsCheckedProperty, binding);
        Grid.SetColumn(rb, colIdx);
        return rb;
    }

    private static void SyncScroll(ScrollViewer ours, ScrollViewer theirs)
    {
        theirs.ScrollChanged -= OnTheirsScrollChanged;
        theirs.ScrollChanged += OnTheirsScrollChanged;

        void OnTheirsScrollChanged(object s, ScrollChangedEventArgs ev)
        {
            ours.ScrollToHorizontalOffset(ev.HorizontalOffset);
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
