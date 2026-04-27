using System.Runtime.InteropServices;
using System.Windows.Forms;

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
/// 聚光灯控制器：挂接 Excel 事件，计算色条位置并驱动 CrosslightOverlay。
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

        TryRefresh();
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

    private static void OnSelectionChange(object sh, Range target) => TryRefresh();
    private static void OnWorkbookActivate(object wb)              => TryRefresh();
    private static void OnWorkbookDeactivate(object wb)            => CrosslightOverlay.Instance.ClearBands();
    private static void OnWindowActivate(object wb, object wn)     => TryRefresh();
    private static void OnWindowDeactivate(object wb, object wn)   => CrosslightOverlay.Instance.ClearBands();

    private static void TryRefresh()
    {
        if (_app == null) return;
        try
        {
            var win = _app.ActiveWindow;
            if (win == null) { CrosslightOverlay.Instance.ClearBands(); return; }

            var sel = _app.Selection as Range;
            if (sel == null) { CrosslightOverlay.Instance.ClearBands(); return; }

            var ws = (Worksheet)_app.ActiveSheet;

            // 冻结行/列的分割位置（0 = 无冻结）
            int splitRow = win.FreezePanes ? (int)win.SplitRow : 0;
            int splitCol = win.FreezePanes ? (int)win.SplitColumn : 0;

            // 冻结区域的物理尺寸（points）= 第一可滚动行/列的绝对坐标
            double frozenH = splitRow > 0 ? ((Range)ws.Cells[splitRow + 1, 1]).Top  : 0.0;
            double frozenW = splitCol > 0 ? ((Range)ws.Cells[1, splitCol + 1]).Left : 0.0;

            // 可滚动窗格第一可见行/列的绝对坐标
            double scrollTop  = ((Range)ws.Cells[win.ScrollRow, 1]).Top;
            double scrollLeft = ((Range)ws.Cells[1, win.ScrollColumn]).Left;

            // PointsToScreenPixelsX/Y(0) 对应可滚动窗格的屏幕原点（冻结区之下/右）。
            // 冻结区内的单元格偏移量 = sel.Top - frozenH（负值，在原点之上）。
            // 可滚动区的单元格偏移量 = sel.Top - scrollTop。
            double viewTop, viewBottom, viewLeft, viewRight;

            if (sel.Row <= splitRow)
            {
                viewTop    = sel.Top              - frozenH;
                viewBottom = sel.Top + sel.Height - frozenH;
            }
            else
            {
                viewTop    = sel.Top              - scrollTop;
                viewBottom = sel.Top + sel.Height - scrollTop;
            }

            if (sel.Column <= splitCol)
            {
                viewLeft  = sel.Left             - frozenW;
                viewRight = sel.Left + sel.Width - frozenW;
            }
            else
            {
                viewLeft  = sel.Left             - scrollLeft;
                viewRight = sel.Left + sel.Width - scrollLeft;
            }

            int rowTop    = win.PointsToScreenPixelsY((int)viewTop);
            int rowBottom = win.PointsToScreenPixelsY((int)viewBottom);
            int colLeft   = win.PointsToScreenPixelsX((int)viewLeft);
            int colRight  = win.PointsToScreenPixelsX((int)viewRight);

            var vs = SystemInformation.VirtualScreen;

            CrosslightOverlay.Instance.UpdateBands(
                new System.Drawing.Rectangle(0, rowTop - vs.Top, vs.Width, Math.Max(2, rowBottom - rowTop)),
                new System.Drawing.Rectangle(colLeft - vs.Left, 0, Math.Max(2, colRight - colLeft), vs.Height));
        }
        catch
        {
            CrosslightOverlay.Instance.ClearBands();
        }
    }
}
