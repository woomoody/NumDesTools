using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace NumDesTools;

/// <summary>
/// 全屏透明叠加窗口，绘制十字聚光灯色条。
/// </summary>
internal sealed class CrosslightOverlay : Form
{
    private const int  GWL_EXSTYLE       = -20;
    private const int  WS_EX_LAYERED     = 0x00080000;
    private const int  WS_EX_TRANSPARENT = 0x00000020;
    private const int  WS_EX_TOOLWINDOW  = 0x00000080;
    private const int  WS_EX_NOACTIVATE  = 0x08000000;
    private const uint LWA_COLORKEY      = 0x00000001;
    private const uint LWA_ALPHA         = 0x00000002;

    [DllImport("user32.dll")] static extern int  GetWindowLong(IntPtr h, int i);
    [DllImport("user32.dll")] static extern int  SetWindowLong(IntPtr h, int i, int v);
    [DllImport("user32.dll")] static extern bool SetLayeredWindowAttributes(
        IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    private static readonly Color BackKey  = Color.FromArgb(1, 1, 1);
    private static readonly Color RowColor = Color.FromArgb(255, 200, 50);
    private static readonly Color ColColor = Color.FromArgb(50,  160, 255);
    private const byte OverlayAlpha = 75;

    private System.Drawing.Rectangle _rowBand;
    private System.Drawing.Rectangle _colBand;
    private bool _hasBands;

    private static CrosslightOverlay? _instance;
    public  static CrosslightOverlay  Instance => _instance ??= new CrosslightOverlay();

    private CrosslightOverlay()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar   = false;
        TopMost         = true;
        BackColor       = BackKey;

        var vs = SystemInformation.VirtualScreen;
        SetBounds(vs.Left, vs.Top, vs.Width, vs.Height);

        SetStyle(ControlStyles.OptimizedDoubleBuffer
               | ControlStyles.AllPaintingInWmPaint
               | ControlStyles.UserPaint, true);

        Show();

        int ex = GetWindowLong(Handle, GWL_EXSTYLE);
        ex |= WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
        SetWindowLong(Handle, GWL_EXSTYLE, ex);

        SetLayeredWindowAttributes(Handle,
            (uint)(BackKey.B | (BackKey.G << 8) | (BackKey.R << 16)),
            OverlayAlpha,
            LWA_COLORKEY | LWA_ALPHA);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
            return cp;
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackKey);
        if (!_hasBands) return;

        using var rowBrush = new SolidBrush(RowColor);
        using var colBrush = new SolidBrush(ColColor);
        e.Graphics.FillRectangle(rowBrush, _rowBand);
        e.Graphics.FillRectangle(colBrush, _colBand);
    }

    public void UpdateBands(System.Drawing.Rectangle rowBand, System.Drawing.Rectangle colBand)
    {
        _rowBand  = rowBand;
        _colBand  = colBand;
        _hasBands = true;
        SafeInvalidate();
    }

    public void ClearBands()
    {
        _hasBands = false;
        SafeInvalidate();
    }

    private void SafeInvalidate()
    {
        if (IsDisposed) return;
        if (InvokeRequired) BeginInvoke((System.Action)Invalidate);
        else Invalidate();
    }

    public static void DisposeInstance()
    {
        if (_instance is { IsDisposed: false })
        {
            _instance.Close();
            _instance.Dispose();
        }
        _instance = null;
    }
}

/// <summary>
/// Win32：找 EXCEL7 窗口并获取其屏幕客户区原点（即单元格网格区域左上角，含行号/列标头）。
/// </summary>
internal static class ExcelGridWin32
{
    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hWnd, ref POINT pt);

    [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }

    /// <summary>
    /// win.Hwnd 本身就是 EXCEL7 窗口句柄，直接 ClientToScreen(0,0) 即可。
    /// EXCEL7 客户区左上角的屏幕坐标。
    /// </summary>
    public static (int X, int Y)? GetExcel7Origin(IntPtr excel7Hwnd)
    {
        var pt = new POINT();
        return ClientToScreen(excel7Hwnd, ref pt) ? (pt.X, pt.Y) : null;
    }
}

/// <summary>
/// 聚光灯控制器。
/// 坐标方案：
///   xlfGetCell(42~45) 在宏上下文中返回单元格相对于 EXCEL7 客户区左上角的 points 偏移。
///   ExcelGridWin32.GetExcel7Origin 用 Win32 ClientToScreen 取 EXCEL7 客户区左上角屏幕坐标。
///   pixels-per-point 通过两次 PointsToScreenPixelsX 调用之差倒推（含 DPI + Zoom）。
///   三者共享 EXCEL7 客户区作为原点，叠加即为单元格屏幕绝对坐标。
/// </summary>
internal static class CrosslightController
{
    private static Application? _app;
    private static bool _active;

    public static bool IsActive => _active;

    public static void Enable(Application app)
    {
        if (_active) return;
        _active = true;
        _app    = app;

        app.SheetSelectionChange += OnSelectionChange;
        app.WorkbookActivate     += OnWorkbookActivate;
        app.WorkbookDeactivate   += OnWorkbookDeactivate;
        app.WindowActivate       += OnWindowActivate;
        app.WindowDeactivate     += OnWindowDeactivate;

        QueueRefresh();
    }

    public static void Disable()
    {
        if (!_active || _app == null) return;
        _active = false;

        _app.SheetSelectionChange -= OnSelectionChange;
        _app.WorkbookActivate     -= OnWorkbookActivate;
        _app.WorkbookDeactivate   -= OnWorkbookDeactivate;
        _app.WindowActivate       -= OnWindowActivate;
        _app.WindowDeactivate     -= OnWindowDeactivate;

        CrosslightOverlay.Instance.ClearBands();
        _app = null;
    }

    private static void OnSelectionChange(object sh, Range target) => QueueRefresh();
    private static void OnWorkbookActivate(object wb)              => QueueRefresh();
    private static void OnWorkbookDeactivate(object wb)            => CrosslightOverlay.Instance.ClearBands();
    private static void OnWindowActivate(object wb, object wn)     => QueueRefresh();
    private static void OnWindowDeactivate(object wb, object wn)   => CrosslightOverlay.Instance.ClearBands();

    private static void QueueRefresh() => ExcelAsyncUtil.QueueAsMacro(TryRefresh);

    private static void TryRefresh()
    {
        if (_app == null) return;
        try
        {
            var win = _app.ActiveWindow;
            if (win == null) { CrosslightOverlay.Instance.ClearBands(); return; }

            var sel = _app.Selection as Range;
            if (sel == null) { CrosslightOverlay.Instance.ClearBands(); return; }

            // xlfGetCell(42~45) 返回单元格相对于当前窗口内容区左上角的 points 偏移。
            // PointsToScreenPixelsX/Y 以相同的内容区左上角为原点，将 points 转为屏幕像素。
            // 两者原点完全一致，直接传入即可，无需任何额外换算。
            var sheetRef = (ExcelReference)XlCall.Excel(XlCall.xlSheetId);
            var cellRef  = new ExcelReference(
                sel.Row    - 1, sel.Row    - 1 + sel.Rows.Count    - 1,
                sel.Column - 1, sel.Column - 1 + sel.Columns.Count - 1,
                sheetRef.SheetId);

            if (XlCall.Excel(XlCall.xlfGetCell, 42, cellRef) is not double ptLeft   ||
                XlCall.Excel(XlCall.xlfGetCell, 43, cellRef) is not double ptTop    ||
                XlCall.Excel(XlCall.xlfGetCell, 44, cellRef) is not double ptRight  ||
                XlCall.Excel(XlCall.xlfGetCell, 45, cellRef) is not double ptBottom)
            {
                CrosslightOverlay.Instance.ClearBands();
                return;
            }

            int screenTop    = win.PointsToScreenPixelsY((int)ptTop);
            int screenBottom = win.PointsToScreenPixelsY((int)ptBottom);
            int screenLeft   = win.PointsToScreenPixelsX((int)ptLeft);
            int screenRight  = win.PointsToScreenPixelsX((int)ptRight);

            var vs = SystemInformation.VirtualScreen;

            System.Diagnostics.Debug.WriteLine(
                $"[Crosslight] pt L={ptLeft:F1} T={ptTop:F1} R={ptRight:F1} B={ptBottom:F1}" +
                $" | screen L={screenLeft} T={screenTop} R={screenRight} B={screenBottom}" +
                $" | vs L={vs.Left} T={vs.Top} W={vs.Width} H={vs.Height}" +
                $" | bandRow y={screenTop - vs.Top} h={screenBottom - screenTop}" +
                $" | bandCol x={screenLeft - vs.Left} w={screenRight - screenLeft}");

            CrosslightOverlay.Instance.UpdateBands(
                new System.Drawing.Rectangle(0, screenTop - vs.Top, vs.Width, Math.Max(2, screenBottom - screenTop)),
                new System.Drawing.Rectangle(screenLeft - vs.Left, 0, Math.Max(2, screenRight - screenLeft), vs.Height));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Crosslight] {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}");
            CrosslightOverlay.Instance.ClearBands();
        }
    }
}
