using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Rectangle = System.Drawing.Rectangle;
using Timer = System.Windows.Forms.Timer;

namespace NumDesTools;

/// <summary>
/// 聚光灯 overlay：在独立 STA 线程上绘制整行/整列半透明色条，
/// 坐标通过 ActivePane.PointsToScreenPixelsX/Y 精确对齐单元格，键盘导航也能跟随。
/// </summary>
internal sealed class CrosslightOverlay : IDisposable
{
    private static readonly Color RowColor = Color.FromArgb(255, 200, 50);
    private static readonly Color ColColor = Color.FromArgb(50, 160, 255);
    private const byte BandAlpha = 80;

    private static CrosslightOverlay? _instance;
    public static CrosslightOverlay Instance => _instance ??= new CrosslightOverlay();

    private Thread? _staThread;
    private RowColBandForm? _bandForm;
    private readonly object _staLock = new();
    private string? _lastAddress;

    private CrosslightOverlay()
    {
        EnsureStaThread();
    }

    private void EnsureStaThread()
    {
        lock (_staLock)
        {
            if (_staThread is { IsAlive: true })
                return;

            _staThread = new Thread(() =>
            {
                _bandForm = new RowColBandForm();
                System.Windows.Forms.Application.Run();
            })
            {
                IsBackground = true,
                Name = "CrosslightOverlay-STA",
            };
            _staThread.SetApartmentState(ApartmentState.STA);
            _staThread.Start();

            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (_bandForm is null && sw.ElapsedMilliseconds < 2000)
                Thread.Sleep(10);
        }
    }

    public void UpdateCross(Range target, bool forced = false)
    {
        EnsureStaThread();

        // Bug-Scroll：只有选中地址变化（或强制刷新）时才重绘，
        // 避免滚轮滚动触发 WindowActivate/WM_SIZE 使 overlay 随 cell 坐标漂移。
        // forced=true 供窗口移动/缩放时调用，确保 overlay 跟随窗口。
        var address = target.Address[true, true, XlReferenceStyle.xlA1, true];
        if (!forced && address == _lastAddress)
            return;
        _lastAddress = address;

        var cellRect = CellPositionProbe.GetCellScreenRect(target);
        if (cellRect == Rectangle.Empty)
            return;

        var excelHwnd = (IntPtr)AppServices.App.Hwnd;
        var currentScreen = Screen.FromHandle(excelHwnd);
        var screenBounds = currentScreen.Bounds;
        if (
            cellRect.Width > screenBounds.Width * 0.9
            || cellRect.Height > screenBounds.Height * 0.9
        )
            return;

        var win = AppServices.App.ActiveWindow;

        // gridLeft/gridTop 来自 Panes[1]（左上冻结角或唯一窗格）
        dynamic pane1 = win.Panes[1];
        Range firstCell = ((Range)pane1.VisibleRange).Cells[1, 1];
        int gridLeft = (int)pane1.PointsToScreenPixelsX((double)firstCell.Left);
        int gridTop = (int)pane1.PointsToScreenPixelsY((double)firstCell.Top);

        // Bug-FrozenPane：gridRight/gridBottom 必须来自最后一个窗格（可滚动区域末尾），
        // 不能用 ActivePane——若选中的是冻结区域，ActivePane 就是 Panes[1]，
        // 其 VisibleRange 只有冻结行，导致 gridBottom 严重偏小。
        int paneCount = win.Panes.Count;
        dynamic lastPane = win.Panes[paneCount];
        Range lastPaneVisible = lastPane.VisibleRange;
        int lastRow = lastPaneVisible.Rows.Count;
        int lastCol = lastPaneVisible.Columns.Count;
        Range lastCell = lastPaneVisible.Cells[lastRow, lastCol];
        int gridRight = (int)
            lastPane.PointsToScreenPixelsX((double)(lastCell.Left + lastCell.Width));
        int gridBottom = (int)
            lastPane.PointsToScreenPixelsY((double)(lastCell.Top + lastCell.Height));

        // Bug2：用 EXCEL7 子窗口客户区对四边做 clamp，消除末格半露导致的越界
        var gridHwnd = FindExcelGridHwnd(excelHwnd);
        if (gridHwnd != IntPtr.Zero)
        {
            GetClientRect(gridHwnd, out var cr);
            var ptLT = new POINT { X = 0, Y = 0 };
            // Bug2/3：裁掉右侧垂直滚动条和底部水平滚动条占用的像素
            var ptRB = new POINT
            {
                X = cr.Right - GetSystemMetrics(SmCxvscroll),
                Y = cr.Bottom - GetSystemMetrics(SmCyhscroll),
            };
            ClientToScreen(gridHwnd, ref ptLT);
            ClientToScreen(gridHwnd, ref ptRB);
            gridLeft = Math.Max(gridLeft, ptLT.X);
            gridTop = Math.Max(gridTop, ptLT.Y);
            gridRight = Math.Min(gridRight, ptRB.X);
            gridBottom = Math.Min(gridBottom, ptRB.Y);
        }

        var gridRect = Rectangle.FromLTRB(gridLeft, gridTop, gridRight, gridBottom);

        // Bug3：gridRect 与当前屏幕求交，防止异常大矩形；换屏后 screenBounds 已是正确屏幕
        gridRect = Rectangle.Intersect(gridRect, screenBounds);
        if (gridRect.IsEmpty)
            return;

        _bandForm?.BeginInvoke(
            (System.Action)(
                () =>
                {
                    _bandForm.SetExcelHwnd(excelHwnd);
                    _bandForm.ShowBands(cellRect, gridRect);
                }
            )
        );
    }

    public void ClearCross()
    {
        _lastAddress = null;
        _bandForm?.BeginInvoke((System.Action)(() => _bandForm?.HideBands()));
    }

    public void Dispose()
    {
        _bandForm?.BeginInvoke(
            (System.Action)(
                () =>
                {
                    _bandForm?.Close();
                    System.Windows.Forms.Application.ExitThread();
                }
            )
        );
        _bandForm = null;
    }

    public static void DisposeInstance()
    {
        _instance?.Dispose();
        _instance = null;
    }

    // ── Win32 helpers ────────────────────────────────────────────────────────

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr h, out RECT r);

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr h, out RECT r);

    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr h, ref POINT pt);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr h, EnumChildProc cb, IntPtr lp);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr h, System.Text.StringBuilder sb, int max);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);

    private const int SmCxvscroll = 2;
    private const int SmCyhscroll = 3;

    private delegate bool EnumChildProc(IntPtr h, IntPtr lp);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left,
            Top,
            Right,
            Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int X,
            Y;
    }

    /// <summary>
    /// 在 Excel 主窗口子树中找 EXCEL7 类窗口（网格区域），用于精确 clamp gridRect。
    /// </summary>
    private static IntPtr FindExcelGridHwnd(IntPtr excelHwnd)
    {
        IntPtr found = IntPtr.Zero;
        EnumChildWindows(
            excelHwnd,
            (h, _) =>
            {
                var sb = new System.Text.StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() == "EXCEL7")
                {
                    found = h;
                    return false; // 停止枚举
                }
                return true;
            },
            IntPtr.Zero
        );
        return found;
    }

    // ── RowColBandForm ────────────────────────────────────────────────────────

    private sealed class RowColBandForm : Form
    {
        private const int GwlExstyle = -20;
        private const int WsExLayered = 0x00080000;
        private const int WsExTransparent = 0x00000020;
        private const int WsExToolwindow = 0x00000080;
        private const int WsExNoactivate = 0x08000000;
        private const int SwHide = 0;

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr h, int i);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr h, int i, int v);

        [DllImport("user32.dll")]
        private static extern bool SetLayeredWindowAttributes(
            IntPtr h,
            uint key,
            byte alpha,
            uint flags
        );

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr h, int cmd);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(
            IntPtr h,
            IntPtr insertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint flags
        );

        private static readonly IntPtr HwndTopmost = new(-1);
        private const uint SwpNoactivate = 0x0010;
        private const uint SwpShowwindow = 0x0040;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr h, System.Text.StringBuilder sb, int max);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pidNull);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        private readonly BandStrip _rowBand;
        private readonly BandStrip _colBand;
        private readonly Timer _focusTimer;
        private IntPtr _excelHwnd;
        private Rectangle _lastCellRect;
        private Rectangle _lastGridRect;

        public RowColBandForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            Size = new Size(0, 0);
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-100, -100);
            Opacity = 0;
            Show();

            _rowBand = new BandStrip(RowColor, BandAlpha);
            _colBand = new BandStrip(ColColor, BandAlpha);

            _focusTimer = new Timer { Interval = 300 };
            _focusTimer.Tick += OnFocusCheck;
            _focusTimer.Start();
        }

        public void SetExcelHwnd(IntPtr hwnd) => _excelHwnd = hwnd;

        private void OnFocusCheck(object? sender, EventArgs e)
        {
            if (!_rowBand.IsVisible && !_colBand.IsVisible)
                return;
            if (_excelHwnd == IntPtr.Zero)
                return;
            try
            {
                // 切换到其他应用时隐藏
                var fg = GetForegroundWindow();
                GetWindowThreadProcessId(fg, out uint fgPid);
                GetWindowThreadProcessId(_excelHwnd, out uint excelPid);
                if (fgPid != excelPid)
                {
                    HideBands();
                    return;
                }

                // Backstage 检测：attach 到 Excel UI 线程，用 GetFocus 拿到真实焦点窗口
                // Backstage 打开时焦点落在非 EXCEL7 的子窗口上
                uint myTid = GetWindowThreadProcessId(Handle, IntPtr.Zero);
                uint excelTid = GetWindowThreadProcessId(_excelHwnd, IntPtr.Zero);
                AttachThreadInput(myTid, excelTid, true);
                IntPtr focusHwnd;
                try
                {
                    focusHwnd = GetFocus();
                }
                finally
                {
                    AttachThreadInput(myTid, excelTid, false);
                }

                if (focusHwnd != IntPtr.Zero)
                {
                    var sb = new System.Text.StringBuilder(64);
                    GetClassName(focusHwnd, sb, 64);
                    // EXCEL7 = 网格区域；NUIDialog / NetUIHWND = Backstage / Ribbon 弹出
                    if (sb.ToString() != "EXCEL7")
                        HideBands();
                }
            }
            catch { }
        }

        public void ShowBands(Rectangle cellRect, Rectangle gridRect)
        {
            // Bug-Scroll：cellRect 和 gridRect 均未变化时（滚轮未改变选中）跳过重绘，
            // 避免 overlay 在内容滚动时产生位移感。
            if (cellRect == _lastCellRect && gridRect == _lastGridRect)
                return;
            _lastCellRect = cellRect;
            _lastGridRect = gridRect;

            _rowBand.PlaceAndShow(gridRect.Left, cellRect.Top, gridRect.Width, cellRect.Height);
            _colBand.PlaceAndShow(cellRect.Left, gridRect.Top, cellRect.Width, gridRect.Height);
        }

        public void HideBands()
        {
            _lastCellRect = Rectangle.Empty;
            _lastGridRect = Rectangle.Empty;
            _rowBand.HideBand();
            _colBand.HideBand();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _focusTimer.Dispose();
                _rowBand.Dispose();
                _colBand.Dispose();
            }
            base.Dispose(disposing);
        }

        private sealed class BandStrip : IDisposable
        {
            private readonly Form _form;

            public BandStrip(Color color, byte alpha)
            {
                _form = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    ShowInTaskbar = false,
                    TopMost = true,
                    BackColor = color,
                    AutoScaleMode = AutoScaleMode.None,
                    StartPosition = FormStartPosition.Manual,
                };

                var ex = GetWindowLong(_form.Handle, GwlExstyle);
                SetWindowLong(
                    _form.Handle,
                    GwlExstyle,
                    ex | WsExLayered | WsExTransparent | WsExToolwindow | WsExNoactivate
                );
                SetLayeredWindowAttributes(
                    _form.Handle,
                    0,
                    alpha,
                    0x2 /* LWA_ALPHA */
                );
            }

            [DllImport("user32.dll")]
            private static extern bool IsWindowVisible(IntPtr h);

            public bool IsVisible => IsWindowVisible(_form.Handle);

            public void PlaceAndShow(int x, int y, int w, int h)
            {
                if (w <= 0 || h <= 0)
                    return;
                SetWindowPos(_form.Handle, HwndTopmost, x, y, w, h, SwpNoactivate | SwpShowwindow);
            }

            public void HideBand() => ShowWindow(_form.Handle, SwHide);

            public void Dispose() => _form.Dispose();
        }
    }
}

// ── CrosslightController ─────────────────────────────────────────────────────

internal static class CrosslightController
{
    private static Application? _app;
    private static bool _active;
    private static Range? _lastTarget;

    // 250 ms 轮询定时器：单元格屏幕坐标因滚动/窗口移动而变化时，
    // 以 forced=true 重算坐标；ShowBands 内部的矩形缓存在坐标未变时直接跳过，
    // 避免频繁 SetWindowPos。
    private static System.Timers.Timer? _refreshTimer;

    public static bool IsActive => _active;

    public static void Enable(Application app)
    {
        if (_active)
        {
            TriggerCurrent();
            return;
        }
        _active = true;
        _app = app;

        app.SheetSelectionChange += OnSelectionChange;
        app.WindowDeactivate += OnWindowDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
        app.WindowActivate += OnWindowActivate;
        app.SheetDeactivate += OnSheetDeactivate;
        app.WorkbookBeforeClose += OnWorkbookBeforeClose;

        _refreshTimer = new System.Timers.Timer(250) { AutoReset = true };
        _refreshTimer.Elapsed += (_, _) =>
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (_lastTarget is not null)
                    try
                    {
                        CrosslightOverlay.Instance.UpdateCross(_lastTarget, forced: true);
                    }
                    catch { }
            });
        _refreshTimer.Start();

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

        _refreshTimer?.Stop();
        _refreshTimer?.Dispose();
        _refreshTimer = null;
        _lastTarget = null;

        CrosslightOverlay.Instance.ClearCross();
        _app = null;
    }

    private static void TriggerCurrent()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                if (AppServices.App.Selection is Range sel)
                {
                    _lastTarget = sel;
                    CrosslightOverlay.Instance.UpdateCross(sel);
                }
            }
            catch { }
        });
    }

    private static void OnSelectionChange(object sh, Range target)
    {
        if (AppServices.App.CutCopyMode != 0)
            return;
        _lastTarget = target;
        ExcelAsyncUtil.QueueAsMacro(() => CrosslightOverlay.Instance.UpdateCross(target));
    }

    private static void OnWindowDeactivate(object wb, object wn) =>
        CrosslightOverlay.Instance.ClearCross();

    private static void OnWorkbookDeactivate(object wb) => CrosslightOverlay.Instance.ClearCross();

    // Bug3：WindowActivate 经 QueueAsMacro 延迟一个宏周期，等 Excel 完成渲染再触发
    private static void OnWindowActivate(object wb, object wn) =>
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                TriggerCurrent();
            }
            catch { }
        });

    private static void OnSheetDeactivate(object sh) => CrosslightOverlay.Instance.ClearCross();

    private static void OnWorkbookBeforeClose(Workbook wb, ref bool cancel) =>
        CrosslightOverlay.Instance.ClearCross();
}

// ── ExcelWindowWatcher ────────────────────────────────────────────────────────

/// <summary>
/// 通过 NativeWindow subclass 监听 Excel 主窗口的 WM_MOVE / WM_SIZE 消息，
/// 触发时通知调用方重新计算 Overlay 坐标，实现窗口拖动时色条实时跟随。
/// 仅当窗口屏幕 RECT 真正改变时才回调，过滤 Excel 内部滚动触发的虚假 WM_SIZE。
/// </summary>
internal sealed class ExcelWindowWatcher : NativeWindow
{
    private const int WmMove = 0x0003;
    private const int WmSize = 0x0005;
    private const int WmDestroy = 0x0002;

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr h, out RECT r);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left,
            Top,
            Right,
            Bottom;
    }

    private readonly System.Action _onMoved;
    private RECT _lastRect;

    public ExcelWindowWatcher(IntPtr hwnd, System.Action onMoved)
    {
        _onMoved = onMoved;
        AssignHandle(hwnd);
        GetWindowRect(hwnd, out _lastRect);
    }

    protected override void WndProc(ref Message m)
    {
        base.WndProc(ref m);
        switch (m.Msg)
        {
            case WmMove:
            case WmSize:
                if (
                    GetWindowRect(Handle, out var rect)
                    && (
                        rect.Left != _lastRect.Left
                        || rect.Top != _lastRect.Top
                        || rect.Right != _lastRect.Right
                        || rect.Bottom != _lastRect.Bottom
                    )
                )
                {
                    _lastRect = rect;
                    try
                    {
                        _onMoved();
                    }
                    catch { }
                }
                break;
            case WmDestroy:
                ReleaseHandle();
                break;
        }
    }
}
