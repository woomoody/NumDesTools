namespace NumDesTools;

/// <summary>
/// 视口范围工具：union 当前窗口所有窗格的可见区域，再裁剪到 UsedRange。
/// 供 CellHighlighter 和 CellSpotlightHighlighter 共用。
/// </summary>
internal static class ViewportHelper
{
    public static Range GetViewportRange(Worksheet ws)
    {
        var usedRange = ws.UsedRange;
        if (usedRange is null)
            return null;

        var win = AppServices.App.ActiveWindow;
        if (win == null)
            return usedRange;

        Range visible = null;
        for (int i = 1; i <= win.Panes.Count; i++)
        {
            try
            {
                var pr = win.Panes[i].VisibleRange;
                visible = visible == null ? pr : AppServices.App.Union(visible, pr);
            }
            catch { }
        }

        if (visible == null)
            return usedRange;

        return AppServices.App.Intersect(usedRange, visible) ?? usedRange;
    }
}

/// <summary>
/// 用 Interior.Color 直接染色实现同值高亮，不使用条件格式，不影响单元格编辑。
/// 染色前记录每格原始颜色，清除时精确还原，不破坏用户已有背景色。
/// 搜索范围限制在当前视口，避免大表卡顿。
/// </summary>
internal static class CellHighlighter
{
    private const int HighlightColor = 0xFFFF00;

    private static readonly List<(Range Cell, int ColorIndex, object Color)> _lastHighlighted = [];
    private static Worksheet _lastSheet;

    public static void Highlight(Worksheet ws, Range target)
    {
        App.ScreenUpdating = false;
        try
        {
            ClearPrevious(ws);

            var searchValue = target.Cells[1, 1].Text as string;
            if (string.IsNullOrEmpty(searchValue))
                return;

            var searchRange = ViewportHelper.GetViewportRange(ws);
            if (searchRange is null)
                return;

            var first = searchRange.Find(
                searchValue,
                LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlWhole,
                MatchCase: true
            );
            if (first == null)
                return;

            var firstAddress = first.Address;
            var current = first;
            do
            {
                var colorIndex = (int)current.Interior.ColorIndex;
                var originalColor = colorIndex == (int)XlColorIndex.xlColorIndexNone
                    ? null
                    : current.Interior.Color;
                current.Interior.Color = HighlightColor;
                _lastHighlighted.Add((current, colorIndex, originalColor));
                current = searchRange.FindNext(current);
            } while (current != null && current.Address != firstAddress);

            _lastSheet = ws;
        }
        finally
        {
            App.ScreenUpdating = true;
        }
    }

    public static void ClearAll()
    {
        App.ScreenUpdating = false;
        try
        {
            ClearPrevious(null);
        }
        finally
        {
            App.ScreenUpdating = true;
        }
    }

    private static void ClearPrevious(Worksheet ws)
    {
        if (_lastHighlighted.Count == 0)
            return;

        if (ws != null && _lastSheet != null && _lastSheet != ws)
        {
            _lastHighlighted.Clear();
            _lastSheet = null;
            return;
        }

        foreach (var (cell, colorIndex, originalColor) in _lastHighlighted)
        {
            try
            {
                if (colorIndex == (int)XlColorIndex.xlColorIndexNone)
                    cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                else
                    cell.Interior.Color = originalColor;
            }
            catch
            {
                // 单元格可能已失效，忽略
            }
        }

        _lastHighlighted.Clear();
        _lastSheet = null;
    }

    private static Application App => AppServices.App;
}

internal static class CellSpotlightHighlighter
{
    private const int RowColor = 0xFFC832;
    private const int ColColor = 0x32A0FF;

    private static readonly List<(Range Cell, int ColorIndex, object Color)> _lastHighlighted = [];
    private static Worksheet _lastSheet;

    public static void Highlight(Worksheet ws, Range target)
    {
        App.ScreenUpdating = false;
        try
        {
            ClearPrevious(ws);

            var scope = ViewportHelper.GetViewportRange(ws);
            if (scope is null)
                return;

            var anchor = target.Cells[1, 1];
            var rowRange = App.Intersect(anchor.EntireRow, scope);
            var colRange = App.Intersect(anchor.EntireColumn, scope);

            SaveColors(rowRange);
            SaveColors(colRange);
            if (rowRange != null) rowRange.Interior.Color = RowColor;
            if (colRange != null) colRange.Interior.Color = ColColor;

            _lastSheet = ws;
        }
        finally
        {
            App.ScreenUpdating = true;
        }
    }

    private static void SaveColors(Range range)
    {
        if (range is null)
            return;
        foreach (Range cell in range)
        {
            var colorIndex = (int)cell.Interior.ColorIndex;
            var originalColor = colorIndex == (int)XlColorIndex.xlColorIndexNone
                ? null
                : cell.Interior.Color;
            _lastHighlighted.Add((cell, colorIndex, originalColor));
        }
    }

    public static void ClearAll()
    {
        App.ScreenUpdating = false;
        try
        {
            ClearPrevious(null);
        }
        finally
        {
            App.ScreenUpdating = true;
        }
    }

    private static void ClearPrevious(Worksheet ws)
    {
        if (_lastHighlighted.Count == 0)
            return;

        if (ws != null && _lastSheet != null && _lastSheet != ws)
        {
            _lastHighlighted.Clear();
            _lastSheet = null;
            return;
        }

        foreach (var (cell, colorIndex, originalColor) in _lastHighlighted)
        {
            try
            {
                if (colorIndex == (int)XlColorIndex.xlColorIndexNone)
                    cell.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                else
                    cell.Interior.Color = originalColor;
            }
            catch
            {
                // 单元格可能已失效，忽略
            }
        }

        _lastHighlighted.Clear();
        _lastSheet = null;
    }

    private static Application App => AppServices.App;
}

internal static class CellHighlightController
{
    private static Application _app;
    private static bool _active;

    public static void Enable(Application app)
    {
        if (_active)
            return;
        _active = true;
        _app = app;
        app.SheetSelectionChange += OnSelectionChange;
        app.SheetDeactivate += OnSheetDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
        app.WorkbookBeforeClose += OnWorkbookBeforeClose;
    }

    public static void Disable()
    {
        if (!_active || _app is null)
            return;
        _active = false;
        _app.SheetSelectionChange -= OnSelectionChange;
        _app.SheetDeactivate -= OnSheetDeactivate;
        _app.WorkbookDeactivate -= OnWorkbookDeactivate;
        _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;
        CellHighlighter.ClearAll();
        _app = null;
    }

    private static void OnSelectionChange(object sh, Range target)
    {
        if (AppServices.App.CutCopyMode != 0)
            return;
        if (sh is Worksheet ws)
            CellHighlighter.Highlight(ws, target);
    }

    private static void OnSheetDeactivate(object sh) => CellHighlighter.ClearAll();

    private static void OnWorkbookDeactivate(object wb) => CellHighlighter.ClearAll();

    private static void OnWorkbookBeforeClose(Workbook wb, ref bool cancel) =>
        CellHighlighter.ClearAll();
}
