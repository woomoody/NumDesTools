using System.Drawing;
using System.IO;
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
    private static readonly object _instanceLock = new();

    public static CrosslightOverlay Instance
    {
        get
        {
            if (_instance != null)
                return _instance;
            lock (_instanceLock)
                return _instance ??= new CrosslightOverlay();
        }
    }

    // ── 诊断日志 ─────────────────────────────────────────────────────────────
    private static readonly string _diagLog = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "tmp",
        "crosslight_diag.log"
    );
    private static readonly object _logLock = new();

    private static void DiagLog(string msg)
    {
#if DEBUG
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_diagLog)!);
            lock (_logLock)
                File.AppendAllText(_diagLog, $"{DateTime.Now:HH:mm:ss.fff} {msg}\n");
        }
        catch { }
#endif
    }

    private Thread? _staThread;
    private volatile RowColBandForm? _bandForm;
    private readonly object _staLock = new();

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
        // 宏体首行闸门：读任何 COM 之前作废，封死入队→执行之间的 TOCTOU 窗口。
        if (CrosslightController.ShouldSuppressMacro())
        {
            PluginLog.Verbose("[crosslight] UpdateCross suppressed");
            return;
        }

        // 图表激活（in-place 编辑/图表 Sheet）或非 Range 对象选中时清除并跳过，
        // 避免 PointsToScreenPixelsX/Y 等 COM 调用扰动图表 UI 导致闪烁。
        try
        {
            if (AppServices.App.ActiveChart != null || AppServices.App.Selection is not Range)
            {
                ClearCross();
                return;
            }
        }
        catch (Exception ex)
        {
            PluginLog.Write(
                $"[crosslight] UpdateCross chartCheck EX: {ex.GetType().Name} {ex.Message}"
            );
            return;
        }

        EnsureStaThread();

        var addr = "";
        try
        {
            addr = target.Address[false, false];
        }
        catch { }

        var cellRect = CellPositionProbe.GetCellScreenRect(target);
        if (cellRect == Rectangle.Empty)
        {
            DiagLog($"[{addr}] cellRect=Empty → return");
            return;
        }

        var excelHwnd = (IntPtr)AppServices.App.Hwnd;
        var currentScreen = Screen.FromHandle(excelHwnd);
        var screenBounds = currentScreen.Bounds;
        if (
            cellRect.Width > screenBounds.Width * 0.9
            || cellRect.Height > screenBounds.Height * 0.9
        )
        {
            DiagLog($"[{addr}] cellRect too large ({cellRect.Width}x{cellRect.Height}) → return");
            return;
        }

        var win = AppServices.App.ActiveWindow;

        // gridRect 策略：优先用 EXCEL7 子窗口客户区（最准确），
        // 再用 Panes[1] 首格坐标限定左上边界（排除行列标头区），
        // 不再从 VisibleRange 末行推算右下——末行可能是整表末行，高度异常大。
        int paneCount = win.Panes.Count;
        var win2 = AppServices.App.ActiveWindow;
        DiagLog(
            $"[{addr}] panes={paneCount} splitR={(int)win2.SplitRow} splitC={(int)win2.SplitColumn}"
                + $" cellRect=({cellRect.Left},{cellRect.Top},{cellRect.Width},{cellRect.Height})"
        );

        // 默认退路：用屏幕边界
        int gridLeft = screenBounds.Left;
        int gridTop = screenBounds.Top;
        int gridRight = screenBounds.Right;
        int gridBottom = screenBounds.Bottom;

        var gridHwnd = FindExcelGridHwnd(excelHwnd);
        if (gridHwnd != IntPtr.Zero)
        {
            // EXCEL7 客户区 → 屏幕坐标，减去滚动条
            GetClientRect(gridHwnd, out var cr);
            var ptLT = new POINT { X = 0, Y = 0 };
            var ptRB = new POINT
            {
                X = cr.Right - GetSystemMetrics(SmCxvscroll),
                Y = cr.Bottom - GetSystemMetrics(SmCyhscroll),
            };
            ClientToScreen(gridHwnd, ref ptLT);
            ClientToScreen(gridHwnd, ref ptRB);
            gridLeft = ptLT.X;
            gridTop = ptLT.Y;
            gridRight = ptRB.X;
            gridBottom = ptRB.Y;
            DiagLog($"[{addr}] EXCEL7: ({gridLeft},{gridTop}→{gridRight},{gridBottom})");
        }

        // 用 Panes[1] 第一个可见格的屏幕坐标收紧左/上边界（排除行列标头）
        dynamic pane1 = win.Panes[1];
        Range firstCell = ((Range)pane1.VisibleRange).Cells[1, 1];
        int firstX = (int)pane1.PointsToScreenPixelsX((double)firstCell.Left);
        int firstY = (int)pane1.PointsToScreenPixelsY((double)firstCell.Top);
        gridLeft = Math.Max(gridLeft, firstX);
        gridTop = Math.Max(gridTop, firstY);
        DiagLog($"[{addr}] pane1 first=({firstX},{firstY}) → gridLT=({gridLeft},{gridTop})");

        var gridRect = Rectangle.FromLTRB(gridLeft, gridTop, gridRight, gridBottom);
        gridRect = Rectangle.Intersect(gridRect, screenBounds);
        if (gridRect.IsEmpty)
        {
            DiagLog($"[{addr}] gridRect empty after intersect → return");
            return;
        }

        var cellCenter = new Point(
            cellRect.Left + cellRect.Width / 2,
            cellRect.Top + cellRect.Height / 2
        );
        bool inGrid = gridRect.Contains(cellCenter);
        DiagLog(
            $"[{addr}] gridRect=({gridRect.Left},{gridRect.Top},{gridRect.Width},{gridRect.Height})"
                + $" cellCenter=({cellCenter.X},{cellCenter.Y}) inGrid={inGrid}"
        );

        if (!inGrid)
        {
            _bandForm?.BeginInvoke((System.Action)(() => _bandForm?.HideBands()));
            return;
        }

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
        _bandForm?.BeginInvoke((System.Action)(() => _bandForm?.HideBands()));
    }

    /// <summary>
    /// 把 action 派发到 STA 消息线程执行（供 CrosslightController 安装 WinEventHook 用）。
    /// </summary>
    public void BeginInvokeOnSta(System.Action action) => _bandForm?.BeginInvoke(action);

    public void Dispose()
    {
        // 先取出引用再置 null，lambda 用局部变量，避免与 BeginInvoke 竞态
        var form = _bandForm;
        _bandForm = null;
        try
        {
            form?.BeginInvoke(
                (System.Action)(
                    () =>
                    {
                        try
                        {
                            form.Close();
                            System.Windows.Forms.Application.ExitThread();
                        }
                        catch { }
                    }
                )
            );
        }
        catch { }
    }

    public static void DisposeInstance()
    {
        CrosslightOverlay? old;
        lock (_instanceLock)
        {
            old = _instance;
            _instance = null;
        }
        old?.Dispose();
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

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public int cbSize;
            public uint flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll")]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        private readonly BandStrip _rowBand;
        private readonly BandStrip _colBand;
        private readonly Timer _focusTimer;
        private IntPtr _excelHwnd;

        // _focusTimer 每 300ms 刷新，避免每次 ShowBands 都做 AttachThreadInput。
        // 初始 false，SetExcelHwnd 设置 hwnd 后由第一次 OnFocusCheck 赋真实值。
        private bool _gridFocused;

        // Excel 内部编辑中（批注/单元格/公式栏），冻结 overlay 位置，
        // 不调 SetWindowPos 也不调 HideBands，避免打断编辑状态。
        private bool _editFreeze;

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

        public void SetExcelHwnd(IntPtr hwnd)
        {
            _excelHwnd = hwnd;
            // hwnd 刚设好时立即刷新一次，不等 _focusTimer 的 300ms
            OnFocusCheck(null, EventArgs.Empty);
        }

        private void OnFocusCheck(object? sender, EventArgs e)
        {
            if (_excelHwnd == IntPtr.Zero)
                return;
            try
            {
                // PID 检查：切换到其他应用
                var fg = GetForegroundWindow();
                GetWindowThreadProcessId(fg, out uint fgPid);
                GetWindowThreadProcessId(_excelHwnd, out uint excelPid);
                if (fgPid != excelPid)
                {
                    _gridFocused = false;
                    if (_rowBand.IsVisible || _colBand.IsVisible)
                        HideBands();
                    return;
                }

                // 焦点类名检查：CTP / Backstage / 对话框激活时 EXCEL7 失焦
                // 用 GetGUIThreadInfo 直查 Excel 线程的 focus，避免 AttachThreadInput
                // 干扰 Excel 输入队列（曾导致双击进入编辑状态偶发失灵）
                uint excelTid = GetWindowThreadProcessId(_excelHwnd, IntPtr.Zero);
                var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
                if (!GetGUIThreadInfo(excelTid, ref gti) || gti.hwndFocus == IntPtr.Zero)
                {
                    _gridFocused = false;
                    return;
                }
                var sb = new System.Text.StringBuilder(64);
                GetClassName(gti.hwndFocus, sb, 64);
                var focusClass = sb.ToString();
                var focusState = CrosslightController.ClassifyFocusWindow(focusClass);

                // 超时兜底：任何非 Editing 焦点状态下，若 _editing 卡 true 超过 600ms
                // 且无实际编辑迹象，认为 RICHEDIT60W HIDE 漏发，强制清零。
                // 原仅在 Grid 分支检查，XLDESK/Button 等 default 分支会让 _editing 永久卡死。
                if (
                    focusState != CrosslightController.FocusState.Editing
                    && CrosslightController.IsEditing()
                    && !CrosslightController.IsExcelEditFocusedPublic()
                    && !CrosslightController.IsExcelCaretActivePublic()
                    && CrosslightController.EditingStaleMs() > 600
                )
                {
                    PluginLog.Verbose("[crosslight] stale _editing cleared (HIDE event missed)"
                    );
                    CrosslightController.SetEditing(false);
                }

                switch (focusState)
                {
                    case CrosslightController.FocusState.Grid:
                        CrosslightController.SetNativeDialogActive(false);
                        _gridFocused = true;
                        if (!CrosslightController.IsEditing())
                            _editFreeze = false;
                        break;
                    case CrosslightController.FocusState.Editing:
                        CrosslightController.SetNativeDialogActive(false);
                        _editFreeze = true;
                        // WinEvent 是权威来源；此处作为 WinEvent 漏报时的防御兜底
                        CrosslightController.SetEditing(true);
                        if (_rowBand.IsVisible || _colBand.IsVisible)
                            HideBands();
                        break;
                    default:
                        // 原生对话框（Find/Replace/格式等）激活：屏蔽 refreshTimer 的 COM 调用，
                        // 防止与 Excel 内部操作并发导致堆损坏（0xc0000374）
                        CrosslightController.SetNativeDialogActive(true);
                        PluginLog.Verbose($"[crosslight] else-branch focusClass={focusClass}");
                        _gridFocused = false;
                        _editFreeze = false;
                        if (_rowBand.IsVisible || _colBand.IsVisible)
                            HideBands();
                        break;
                }
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[crosslight] OnFocusCheck EX: {ex.GetType().Name} {ex.Message}");
            }
        }

        public void ShowBands(Rectangle cellRect, Rectangle gridRect)
        {
            // 读缓存字段（_focusTimer 每 300ms 刷新），避免每次 ShowBands 做 AttachThreadInput
            if (_editFreeze)
                return; // Excel 内部编辑中，冻结位置，禁止调 SetWindowPos

            if (!_gridFocused)
            {
                if (_rowBand.IsVisible || _colBand.IsVisible)
                    HideBands();
                return;
            }

            _rowBand.PlaceAndShow(gridRect.Left, cellRect.Top, gridRect.Width, cellRect.Height);
            _colBand.PlaceAndShow(cellRect.Left, gridRect.Top, cellRect.Width, gridRect.Height);
        }

        public void HideBands()
        {
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
            private int _lastX = int.MinValue,
                _lastY = int.MinValue,
                _lastW,
                _lastH;
            private bool _lastVisible;

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
                // 位置和大小未变且已可见：跳过 SetWindowPos，避免 SWP_SHOWWINDOW 触发不必要的重绘闪烁
                if (_lastVisible && x == _lastX && y == _lastY && w == _lastW && h == _lastH)
                    return;
                _lastX = x;
                _lastY = y;
                _lastW = w;
                _lastH = h;
                _lastVisible = true;
                SetWindowPos(_form.Handle, HwndTopmost, x, y, w, h, SwpNoactivate | SwpShowwindow);
            }

            public void HideBand()
            {
                _lastVisible = false;
                ShowWindow(_form.Handle, SwHide);
            }

            public void Dispose() => _form.Dispose();
        }
    }
}

// ── CrosslightController ─────────────────────────────────────────────────────

internal static class CrosslightController
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr h, IntPtr pidNull);

    [DllImport("user32.dll")]
    private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(IntPtr h, System.Text.StringBuilder sb, int max);

    // ── WinEventHook P/Invokes ────────────────────────────────────────────────

    [DllImport("user32.dll")]
    private static extern IntPtr SetWinEventHook(
        uint eventMin,
        uint eventMax,
        IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc,
        uint idProcess,
        uint idThread,
        uint dwFlags
    );

    [DllImport("user32.dll")]
    private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public int cbSize;
        public uint flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public int rcCaretLeft,
            rcCaretTop,
            rcCaretRight,
            rcCaretBottom;
    }

    // GUITHREADINFO.flags bit：线程正在输入文字（单元格编辑 / 公式栏 / 批注文字编辑）
    private const uint GuiCaretBlinking = 0x00000001;

    private const uint EventObjectShow = 0x8002;
    private const uint EventObjectHide = 0x8003;
    private const uint WinEventOutofcontext = 0x0000;
    private const uint WinEventSkipownprocess = 0x0002;

    private delegate void WinEventDelegate(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint dwEventThread,
        uint dwmsEventTime
    );

    private static Application? _app;
    private static bool _active;
    private static Range? _lastTarget;
    private static IntPtr _excelHwnd = IntPtr.Zero;
    private static volatile bool _paused;

    // 诊断计数器（Release 可见）
    private static int _updateCrossTotal;
    private static int _suppressedTotal;
    private static int _setEditingCount;

    // 编辑闸门（权威状态由 WinEventHook 管理）：
    // - RICHEDIT60W SHOW  → SetEditing(true)  → 停 _refreshTimer
    // - RICHEDIT60W HIDE  → 400ms 后 SetEditing(false) → 重启 _refreshTimer
    // - OnFocusCheck Editing 分支可置 true 作为 WinEvent 漏报的兜底
    // - OnSelectionChange 在 inline 编辑结束后负责清零（WinEvent 不覆盖 inline 编辑）
    private static volatile bool _editing;

    // Excel 原生对话框（Find/Replace、格式等）激活时置 true，防止 refreshTimer COM 调用并发
    private static volatile bool _nativeDialogActive;

    // WinEventHook 相关字段
    private static IntPtr _winEventHook = IntPtr.Zero;
    private static WinEventDelegate? _winEventProc; // 必须持有强引用，防 GC 回收后回调崩溃

    // WinEvent HIDE 后的延迟释放计时器（与 _refreshTimer 是不同的两个对象）
    private static System.Timers.Timer? _resumeDelayTimer;

    // generation counter：每次 SHOW/HIDE 事件都递增，用于让 stale 的 HIDE 回调自我失效。
    // System.Timers.Timer.Stop() 无法取消已入队的 ThreadPool 回调，必须靠此字段区分。
    private static int _resumeGen;

    private static System.Timers.Timer? _refreshTimer;

    public static bool IsActive => _active;

    public static void Pause()
    {
        _paused = true;
        CrosslightOverlay.Instance.ClearCross();
    }

    public static void Resume()
    {
        _paused = false;
        if (_active)
            TriggerCurrent();
    }

    internal enum FocusState
    {
        Grid,
        Editing,
        Other,
    }

    /// <summary>
    /// 根据焦点窗口类名判断 overlay 应进入哪种状态，纯函数，可单元测试。
    /// </summary>
    internal static FocusState ClassifyFocusWindow(string className) =>
        className == "EXCEL7" ? FocusState.Grid
        : className == "EDTBX"
        || className == "NetUIHWND"
        || className == "RICHEDIT60W" // Excel 365 批注富文本编辑框
        || className.StartsWith("EXCEL", StringComparison.Ordinal)
            ? FocusState.Editing
        : FocusState.Other;

    private static bool IsExcelForeground(IntPtr excelHwnd)
    {
        var fg = GetForegroundWindow();
        GetWindowThreadProcessId(fg, out uint fgPid);
        GetWindowThreadProcessId(excelHwnd, out uint excelPid);
        return fgPid == excelPid;
    }

    // 检查 Excel 线程是否正在输入（caret 闪烁）—— 不调 COM，纯 Win32，可在 ThreadPool 线程调用
    private static bool IsExcelCaretActive()
    {
        if (_excelHwnd == IntPtr.Zero)
            return false;
        var excelTid = GetWindowThreadProcessId(_excelHwnd, IntPtr.Zero);
        var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
        return GetGUIThreadInfo(excelTid, ref gti) && (gti.flags & GuiCaretBlinking) != 0;
    }

    // 检查 Excel 线程焦点是否在"编辑类"窗口（批注/单元格编辑），可在 ThreadPool 线程调用
    private static bool IsExcelEditFocused()
    {
        if (_excelHwnd == IntPtr.Zero)
            return false;
        var excelTid = GetWindowThreadProcessId(_excelHwnd, IntPtr.Zero);
        var gti = new GUITHREADINFO { cbSize = Marshal.SizeOf<GUITHREADINFO>() };
        if (!GetGUIThreadInfo(excelTid, ref gti) || gti.hwndFocus == IntPtr.Zero)
            return false;
        var sb = new System.Text.StringBuilder(64);
        GetClassName(gti.hwndFocus, sb, 64);
        return ClassifyFocusWindow(sb.ToString()) == FocusState.Editing;
    }

    // _editing 置 true 时的时间戳（TickCount64 ms），用于超时兜底清零
    private static long _editingSetTick;

    internal static bool IsEditing() => _editing;

    // _editing 置 true 至今经过的毫秒数（供 OnFocusCheck 超时兜底用）
    internal static long EditingStaleMs() =>
        _editing ? Environment.TickCount64 - _editingSetTick : 0;

    // 暴露给 RowColBandForm（内部嵌套类）调用
    internal static bool IsExcelCaretActivePublic() => IsExcelCaretActive();

    internal static bool IsExcelEditFocusedPublic() => IsExcelEditFocused();

    // 统一宏抑制闸门：入队前调一次 + 宏体首行调一次，封死 TOCTOU 时间窗。
    // _editing 由 WinEvent 权威管理，保留两道 Win32 探测作为纯防御兜底。
    internal static bool ShouldSuppressMacro() =>
        _paused || _editing || _nativeDialogActive || IsExcelEditFocused() || IsExcelCaretActive();

    internal static void SetNativeDialogActive(bool active) => _nativeDialogActive = active;

    internal static void SetEditing(bool editing)
    {
        if (_editing == editing)
            return;
        _editing = editing;
        if (editing)
            _editingSetTick = Environment.TickCount64;
        var n = System.Threading.Interlocked.Increment(ref _setEditingCount);
        PluginLog.Write(
            $"[crosslight] SetEditing={editing} n={n}"
                + $" tid={Environment.CurrentManagedThreadId}"
                + $" updateTotal={_updateCrossTotal} suppressed={_suppressedTotal}"
        );
        // 编辑期停掉 250ms 主动刷新，消除唯一的主动 COM 打断源；退出编辑再启
        var timer = _refreshTimer; // 快照，避免与 Disable() 的 Dispose 竞态
        if (editing)
            timer?.Stop();
        else if (_active)
        {
            try
            {
                timer?.Start();
            }
            catch (Exception ex)
            {
                PluginLog.Write(
                    $"[crosslight] SetEditing Start EX: {ex.GetType().Name} {ex.Message}"
                );
            }
        }
    }

    // ── WinEventHook ──────────────────────────────────────────────────────────

    // 必须在有消息泵的线程（STA 线程）调用，才能让 OUTOFCONTEXT 回调正常投递。
    // 通过 CrosslightOverlay.Instance.BeginInvokeOnSta(InstallCommentHook) 派发。
    private static void InstallCommentHook()
    {
        if (_excelHwnd == IntPtr.Zero || _winEventHook != IntPtr.Zero)
            return;
        uint excelTid = GetWindowThreadProcessId(_excelHwnd, IntPtr.Zero);
        _winEventProc = OnWinEvent;
        _winEventHook = SetWinEventHook(
            EventObjectShow,
            EventObjectHide,
            IntPtr.Zero,
            _winEventProc,
            0,
            excelTid,
            WinEventOutofcontext | WinEventSkipownprocess
        );
        PluginLog.Write($"[crosslight] WinEventHook installed tid={excelTid} hook={_winEventHook}");
    }

    private static void UninstallCommentHook()
    {
        if (_winEventHook != IntPtr.Zero)
        {
            UnhookWinEvent(_winEventHook);
            _winEventHook = IntPtr.Zero;
        }
        // 不立即 null _winEventProc：WINEVENT_OUTOFCONTEXT 回调是 PostMessage 到 STA 队列的，
        // UnhookWinEvent 后队列里仍可能有 pending callback，立即 null 会让 GC 回收 delegate，
        // 下次回调执行时 function pointer 失效 → STA 线程崩溃。
        // 将 null 操作 BeginInvoke 到 STA 线程，保证在所有 pending callback 执行完后才释放引用。
        var keepAlive = _winEventProc;
        try
        {
            CrosslightOverlay.Instance.BeginInvokeOnSta(() =>
            {
                _ = keepAlive; // 让 GC 保留到此 lambda 执行完
                _winEventProc = null;
            });
        }
        catch
        {
            _winEventProc = null; // overlay 已销毁时退路
        }
        _resumeDelayTimer?.Stop();
        _resumeDelayTimer?.Dispose();
        _resumeDelayTimer = null;
        PluginLog.Write("[crosslight] UninstallCommentHook done");
    }

    private static void OnWinEvent(
        IntPtr hWinEventHook,
        uint eventType,
        IntPtr hwnd,
        int idObject,
        int idChild,
        uint dwEventThread,
        uint dwmsEventTime
    )
    {
        // OnWinEvent 在 STA 线程上执行，任何未处理异常会杀死 Application.Run() → Excel 卡死。
        try
        {
            if (hwnd == IntPtr.Zero)
                return;
            var sb = new System.Text.StringBuilder(64);
            GetClassName(hwnd, sb, 64);
            var cls = sb.ToString();

            // 只响应 RICHEDIT60W（Excel 365 批注富文本编辑框）。
            // NetUIHWND/EDTBX 在 Ribbon/公式栏也会频繁发 SHOW/HIDE，不能作为批注信号。
            // inline 单元格编辑由 IsExcelCaretActive/IsExcelEditFocused 兜底，无需 WinEvent。
            if (cls != "RICHEDIT60W")
                return;

            PluginLog.Write(
                $"[crosslight] WinEvent ev={eventType:X4} cls={cls} idObj={idObject}"
                    + $" tid={Environment.CurrentManagedThreadId}"
            );

            if (eventType == EventObjectShow)
            {
                _resumeDelayTimer?.Stop();
                _resumeDelayTimer?.Dispose(); // 防止 handle 泄漏
                _resumeDelayTimer = null;
                // generation 前进：让已入队的 HIDE 回调检测到 mismatch 后自我失效
                System.Threading.Interlocked.Increment(ref _resumeGen);
                SetEditing(true); // SHOW：立即关闸，同时 SetEditing 内部停 _refreshTimer
                CrosslightOverlay.Instance.ClearCross();
            }
            else // EventObjectHide
            {
                // 延迟 400ms 释放，留出批注真正销毁的余量（远 < 500ms 约束）
                _resumeDelayTimer?.Stop();
                _resumeDelayTimer?.Dispose(); // 防止 handle 泄漏
                // 捕获 generation：回调执行时若值已变（SHOW 到来），则为 stale 直接丢弃
                var expectedGen = System.Threading.Interlocked.Increment(ref _resumeGen);
                _resumeDelayTimer = new System.Timers.Timer(400) { AutoReset = false };
                _resumeDelayTimer.Elapsed += (_, _) =>
                {
                    try
                    {
                        if (_resumeGen != expectedGen)
                            return; // SHOW 已到来，此回调已过期
                        SetEditing(false); // 同时 SetEditing 内部重启 _refreshTimer
                        if (_active)
                            Resume();
                    }
                    catch (Exception ex)
                    {
                        PluginLog.Write(
                            $"[crosslight] resumeDelayTimer EX: {ex.GetType().Name} {ex.Message}"
                        );
                    }
                };
                _resumeDelayTimer.Start();
            }
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[crosslight] OnWinEvent EX: {ex.GetType().Name} {ex.Message}");
        }
    }

    // ── Enable / Disable ─────────────────────────────────────────────────────

    public static void Enable(Application app)
    {
        if (_active)
        {
            TriggerCurrent();
            return;
        }
        PluginLog.Write("[crosslight] Enable start");
        _active = true;
        _app = app;
        _excelHwnd = (IntPtr)app.Hwnd;
        _lastTarget = null; // 清除上一个会话的旧 Range，防止 Timer 在 QueueAsMacro 就绪前调用抛 NRE

        app.SheetSelectionChange += OnSelectionChange;
        app.SheetBeforeDoubleClick += OnSheetBeforeDoubleClick;
        app.WindowDeactivate += OnWindowDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
        app.WindowActivate += OnWindowActivate;
        app.SheetActivate += OnSheetActivate;
        app.SheetDeactivate += OnSheetDeactivate;
        app.WorkbookBeforeClose += OnWorkbookBeforeClose;

        _refreshTimer = new System.Timers.Timer(250) { AutoReset = true };
        _refreshTimer.Elapsed += (_, _) =>
        {
            try
            {
                if (_lastTarget is null || _app is null)
                    return;
                // 入队前闸门：编辑态下不投递宏
                if (ShouldSuppressMacro())
                {
                    System.Threading.Interlocked.Increment(ref _suppressedTotal);
                    return;
                }
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    if (_lastTarget is null || _app is null)
                        return;
                    try
                    {
                        System.Threading.Interlocked.Increment(ref _updateCrossTotal);
                        // 二次校验：封死入队→执行之间用户打开批注的 TOCTOU 窗口，
                        // 读任何 COM 之前再判断一次，为 true 则作废。
                        if (ShouldSuppressMacro() || !IsExcelForeground(_excelHwnd))
                            return;
                        CrosslightOverlay.Instance.UpdateCross(_lastTarget, forced: true);
                    }
                    catch (Exception ex)
                    {
                        PluginLog.Write(
                            $"[crosslight] refreshTimer QueueAsMacro EX: {ex.GetType().Name} {ex.Message}"
                        );
                    }
                });
            }
            catch (Exception ex)
            {
                PluginLog.Write(
                    $"[crosslight] refreshTimer Elapsed EX: {ex.GetType().Name} {ex.Message}"
                );
            }
        };
        _refreshTimer.Start();

        // 必须派发到 STA 消息线程，才能让 WINEVENT_OUTOFCONTEXT 回调正常投递
        CrosslightOverlay.Instance.BeginInvokeOnSta(InstallCommentHook);

        TriggerCurrent();
    }

    public static void Disable()
    {
        UninstallCommentHook();

        if (!_active || _app == null)
            return;
        _active = false;

        _app.SheetSelectionChange -= OnSelectionChange;
        _app.SheetBeforeDoubleClick -= OnSheetBeforeDoubleClick;
        _app.WindowDeactivate -= OnWindowDeactivate;
        _app.WorkbookDeactivate -= OnWorkbookDeactivate;
        _app.WindowActivate -= OnWindowActivate;
        _app.SheetActivate -= OnSheetActivate;
        _app.SheetDeactivate -= OnSheetDeactivate;
        _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;

        _refreshTimer?.Stop();
        _refreshTimer?.Dispose();
        _refreshTimer = null;
        _lastTarget = null;
        _editing = false;
        _excelHwnd = IntPtr.Zero;

        CrosslightOverlay.Instance.ClearCross();
        _app = null;
    }

    private static void TriggerCurrent()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (ShouldSuppressMacro())
                return;
            try
            {
                if (AppServices.App.Selection is Range sel)
                {
                    _lastTarget = sel;
                    CrosslightOverlay.Instance.UpdateCross(sel);
                }
            }
            catch (Exception ex)
            {
                PluginLog.Write(
                    $"[crosslight] TriggerCurrent EX: {ex.GetType().Name} {ex.Message}"
                );
            }
        });
    }

    private static void OnSelectionChange(object sh, Range target)
    {
        if (AppServices.App.CutCopyMode != 0)
            return;
        _lastTarget = target;
        // inline 单元格编辑（非批注）结束后，SelectionChange 触发且 caret 已消失。
        // 此处作为 WinEvent 未覆盖 inline 编辑的安全兜底，防止 _editing 永久为 true。
        if (_editing && !IsExcelCaretActive() && !IsExcelEditFocused())
            SetEditing(false);
        if (ShouldSuppressMacro())
            return;
        ExcelAsyncUtil.QueueAsMacro(() => CrosslightOverlay.Instance.UpdateCross(target));
    }

    // BeforeDoubleClick 不再直接置位 _editing：
    // WinEvent 会在 RICHEDIT60W SHOW 时精确置位，OnFocusCheck Editing 分支作为兜底。
    // 直接置位会导致非批注 double-click 时 _editing 永久为 true（WinEvent 不发 HIDE）。
    private static void OnSheetBeforeDoubleClick(object sh, Range target, ref bool cancel)
    {
        PluginLog.Verbose("[crosslight] BeforeDoubleClick");
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
            catch (Exception ex)
            {
                PluginLog.Write(
                    $"[crosslight] OnWindowActivate EX: {ex.GetType().Name} {ex.Message}"
                );
            }
        });

    private static void OnSheetActivate(object sh) =>
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                TriggerCurrent();
            }
            catch (Exception ex)
            {
                PluginLog.Write(
                    $"[crosslight] OnSheetActivate EX: {ex.GetType().Name} {ex.Message}"
                );
            }
        });

    private static void OnSheetDeactivate(object sh) => CrosslightOverlay.Instance.ClearCross();

    private static void OnWorkbookBeforeClose(Workbook wb, ref bool cancel) =>
        CrosslightOverlay.Instance.ClearCross();
}
