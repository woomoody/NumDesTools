using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Timer = System.Windows.Forms.Timer;

namespace NumDesTools;

/// <summary>
/// 聚光灯：两条细长透明窗口绘制十字线，坐标机制与 CellSelectChangeTip 完全相同（Cursor.Position）。
/// </summary>
internal sealed class CrosslightOverlay : IDisposable
{
    private readonly LineForm _hLine;
    private readonly LineForm _vLine;
    private readonly Timer _scrollTimer;
    private readonly Timer _focusTimer;
    private int _lastScrollRow;
    private int _lastScrollCol;

    private static readonly Color _rowColor = Color.FromArgb(255, 200, 50);
    private static readonly Color _colColor = Color.FromArgb(50, 160, 255);
    private const byte LineAlpha = 160;
    private const int Thickness = 2;

    private static CrosslightOverlay? _instance;
    public static CrosslightOverlay Instance => _instance ??= new CrosslightOverlay();

    private CrosslightOverlay()
    {
        _hLine = new LineForm(_rowColor, LineAlpha);
        _vLine = new LineForm(_colColor, LineAlpha);

        _scrollTimer = new Timer { Interval = 150 };
        _scrollTimer.Tick += OnScrollCheck;

        _focusTimer = new Timer { Interval = 300 };
        _focusTimer.Tick += OnFocusCheck;
        _focusTimer.Start();
    }

    public void UpdateCross()
    {
        var cursor = Cursor.Position;

        // Excel 主窗口范围（同 CellSelectChangeTip 的 Screen.FromPoint(cursor).WorkingArea 思路）
        GetWindowRect((IntPtr)NumDesAddIn.App.Hwnd, out var rc);

        int left = rc.Left;
        int top = rc.Top;
        int right = rc.Right;
        int bottom = rc.Bottom;

        _hLine.Place(left, cursor.Y, right - left, Thickness);
        _vLine.Place(cursor.X, top, Thickness, bottom - top);

        ShowWindow(_hLine.Handle, SW_SHOWNOACTIVATE);
        ShowWindow(_vLine.Handle, SW_SHOWNOACTIVATE);

        try
        {
            var win = NumDesAddIn.App.ActiveWindow;
            if (win != null)
            {
                _lastScrollRow = win.ScrollRow;
                _lastScrollCol = win.ScrollColumn;
            }
        }
        catch
        { /* ignore */
        }
        _scrollTimer.Start();
    }

    public void ClearCross()
    {
        _scrollTimer.Stop();
        if (!_hLine.IsDisposed)
            _hLine.Hide();
        if (!_vLine.IsDisposed)
            _vLine.Hide();
    }

    private void OnScrollCheck(object? sender, EventArgs e)
    {
        try
        {
            var win = NumDesAddIn.App.ActiveWindow;
            if (win == null)
                return;
            if (win.ScrollRow != _lastScrollRow || win.ScrollColumn != _lastScrollCol)
                ClearCross();
        }
        catch
        {
            ClearCross();
        }
    }

    private void OnFocusCheck(object? sender, EventArgs e)
    {
        if (_hLine.IsDisposed || (!_hLine.Visible && !_vLine.Visible))
            return;
        try
        {
            var fg = GetForegroundWindow();
            GetWindowThreadProcessId(fg, out uint fgPid);
            GetWindowThreadProcessId((IntPtr)NumDesAddIn.App.Hwnd, out uint excelPid);
            if (fgPid != excelPid)
                ClearCross();
        }
        catch
        { /* ignore */
        }
    }

    public void Dispose()
    {
        _scrollTimer.Dispose();
        _focusTimer.Dispose();
        _hLine.Dispose();
        _vLine.Dispose();
    }

    public static void DisposeInstance()
    {
        _instance?.Dispose();
        _instance = null;
    }

    private const int SW_SHOWNOACTIVATE = 4;

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr h, int cmd);

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    [DllImport("user32.dll")]
    static extern bool GetWindowRect(IntPtr h, out RECT r);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left,
            Top,
            Right,
            Bottom;
    }

    private sealed class LineForm : Form
    {
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_LAYERED = 0x00080000;
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        [DllImport("user32.dll")]
        static extern int GetWindowLong(IntPtr h, int i);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr h, int i, int v);

        [DllImport("user32.dll")]
        static extern bool SetLayeredWindowAttributes(IntPtr h, uint key, byte alpha, uint flags);

        public LineForm(Color color, byte alpha)
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = color;
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.Manual;

            int ex = GetWindowLong(Handle, GWL_EXSTYLE);
            SetWindowLong(
                Handle,
                GWL_EXSTYLE,
                ex | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
            );
            SetLayeredWindowAttributes(
                Handle,
                0,
                alpha,
                0x00000002 /* LWA_ALPHA */
            );
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |=
                    WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
                return cp;
            }
        }

        public void Place(int x, int y, int w, int h)
        {
            Location = new Point(x, y);
            ClientSize = new Size(w, h);
        }
    }
}

internal static class CrosslightController
{
    private static Application? _app;
    private static bool _active;
    private static bool _fillMode;

    public static bool IsActive => _active;

    public static void Enable(Application app, bool fillMode = false)
    {
        if (_active)
        {
            bool wasFill = _fillMode;
            _fillMode = fillMode;
            // clear whichever mode was running before switching
            if (wasFill)
                CellSpotlightHighlighter.ClearAll();
            else
                CrosslightOverlay.Instance.ClearCross();
            TriggerCurrent();
            return;
        }
        _fillMode = fillMode;
        _active = true;
        _app = app;

        app.SheetSelectionChange += OnSelectionChange;
        app.WindowDeactivate += OnWindowDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
        app.WindowActivate += OnWindowActivate;
        app.SheetDeactivate += OnSheetDeactivate;
        app.WorkbookBeforeClose += OnWorkbookBeforeClose;

        TriggerCurrent();
    }

    public static void Disable()
    {
        if (!_active || _app == null)
            return;
        _active = false;

        _app.SheetSelectionChange -= OnSelectionChange;
        _app.WindowDeactivate -= OnWindowDeactivate;
        _app.WorkbookDeactivate -= OnWorkbookDeactivate;
        _app.WindowActivate -= OnWindowActivate;
        _app.SheetDeactivate -= OnSheetDeactivate;
        _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;

        CrosslightOverlay.Instance.ClearCross();
        CellSpotlightHighlighter.ClearAll();
        _app = null;
    }

    private static void TriggerCurrent()
    {
        if (_fillMode)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var ws = NumDesAddIn.App.ActiveSheet as Worksheet;
                    var sel = NumDesAddIn.App.Selection as Range;
                    if (ws != null && sel != null)
                        CellSpotlightHighlighter.Highlight(ws, sel);
                }
                catch { }
            });
        }
        else
        {
            ExcelAsyncUtil.QueueAsMacro(CrosslightOverlay.Instance.UpdateCross);
        }
    }

    private static void OnSelectionChange(object sh, Range target)
    {
        if (_fillMode)
        {
            if (sh is Worksheet ws)
                ExcelAsyncUtil.QueueAsMacro(() => CellSpotlightHighlighter.Highlight(ws, target));
        }
        else
        {
            ExcelAsyncUtil.QueueAsMacro(CrosslightOverlay.Instance.UpdateCross);
        }
    }

    private static void OnWindowDeactivate(object wb, object wn)
    {
        CrosslightOverlay.Instance.ClearCross();
        CellSpotlightHighlighter.ClearAll();
    }

    private static void OnWorkbookDeactivate(object wb)
    {
        CrosslightOverlay.Instance.ClearCross();
        CellSpotlightHighlighter.ClearAll();
    }

    private static void OnWindowActivate(object wb, object wn) =>
        TriggerCurrent();

    private static void OnSheetDeactivate(object sh) =>
        CellSpotlightHighlighter.ClearAll();

    private static void OnWorkbookBeforeClose(Workbook wb, ref bool cancel) =>
        CellSpotlightHighlighter.ClearAll();
}
