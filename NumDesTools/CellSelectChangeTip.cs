using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Font = System.Drawing.Font;
using Timer = System.Windows.Forms.Timer;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 跟随光标的单元格值气泡提示。
/// - 定位：Cursor.Position 直接作为 Form.Location，同为 WinForms 物理坐标，无需换算
/// - 不抢焦点：ShowWindow(SW_SHOWNOACTIVATE)
/// - Excel 失焦时自动隐藏
/// - 滚动检测：定时器轮询 ScrollRow/ScrollColumn，偏移时隐藏
/// </summary>
public sealed class CellSelectChangeTip : Form
{
    private string? _text;
    private static readonly Font _tipFont = new Font("微软雅黑", 11);
    private const int Pad = 8;

    private readonly Timer _scrollTimer;
    private readonly Timer _focusTimer;
    private int _lastScrollRow;
    private int _lastScrollCol;

    private static CellSelectChangeTip? _instance;
    public static CellSelectChangeTip Instance => _instance ??= new CellSelectChangeTip();

    private CellSelectChangeTip()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(40, 40, 40);
        ForeColor = Color.White;
        AutoScaleMode = AutoScaleMode.None;
        StartPosition = FormStartPosition.Manual;

        SetStyle(
            ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint,
            true
        );

        var ex = GetWindowLong(Handle, GWL_EXSTYLE);
        SetWindowLong(Handle, GWL_EXSTYLE, ex | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE);

        _scrollTimer = new Timer { Interval = 150 };
        _scrollTimer.Tick += OnScrollCheck;

        // 轮询前台窗口：Excel 失焦时隐藏气泡
        _focusTimer = new Timer { Interval = 300 };
        _focusTimer.Tick += OnFocusCheck;
        _focusTimer.Start();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        if (_text == null)
            return;
        using var brush = new SolidBrush(ForeColor);
        e.Graphics.DrawString(_text, _tipFont, brush, new PointF(Pad, Pad));
    }

    // 不拦截鼠标，透传给 Excel
    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_TRANSPARENT | WS_EX_NOACTIVATE;
            return cp;
        }
    }

    public void ShowBubble(string text)
    {
        _text = text;
        var sz = TextRenderer.MeasureText(text, _tipFont);
        int w = sz.Width + Pad * 2;
        int h = sz.Height + Pad * 2;

        var cursor = Cursor.Position;
        int x = cursor.X + 14;
        int y = cursor.Y + 14;

        var wa = Screen.FromPoint(cursor).WorkingArea;
        if (x + w > wa.Right)
            x = cursor.X - w - 2;
        if (y + h > wa.Bottom)
            y = cursor.Y - h - 2;
        if (x < wa.Left)
            x = wa.Left;
        if (y < wa.Top)
            y = wa.Top;

        ClientSize = new Size(w, h);
        Location = new Point(x, y);

        // SW_SHOWNOACTIVATE：显示但不抢 Excel 的键盘焦点
        ShowWindow(Handle, SW_SHOWNOACTIVATE);
        Invalidate();

        // 记录当前滚动位置，用于滚动检测
        try
        {
            var win = NumDesAddIn.App.ActiveWindow;
            _lastScrollRow = win.ScrollRow;
            _lastScrollCol = win.ScrollColumn;
        }
        catch
        { /* ignore */
        }
        _scrollTimer.Start();

        PluginLog.Verbose($"[CellTip] cursor=({cursor.X},{cursor.Y}) loc=({x},{y}) size=({w}x{h})");
    }

    public void ClearBubble()
    {
        _scrollTimer.Stop();
        _text = null;
        if (IsHandleCreated && !IsDisposed)
        {
            if (InvokeRequired)
                BeginInvoke((System.Action)Hide);
            else
                Hide();
        }
    }

    private void OnScrollCheck(object? sender, EventArgs e)
    {
        try
        {
            var win = NumDesAddIn.App.ActiveWindow;
            if (win.ScrollRow != _lastScrollRow || win.ScrollColumn != _lastScrollCol)
                ClearBubble();
        }
        catch
        {
            ClearBubble();
        }
    }

    private void OnFocusCheck(object? sender, EventArgs e)
    {
        if (!Visible)
            return;
        try
        {
            var fg = GetForegroundWindow();
            if (fg == Handle)
                return; // 气泡本身（不应发生，但防御）
            // 比较进程 ID：前台窗口属于 Excel 进程则保留
            GetWindowThreadProcessId(fg, out uint fgPid);
            GetWindowThreadProcessId((IntPtr)NumDesAddIn.App.Hwnd, out uint excelPid);
            if (fgPid != excelPid)
                ClearBubble();
        }
        catch
        { /* ignore */
        }
    }

    public static void DisposeInstance()
    {
        if (_instance is { IsDisposed: false })
        {
            _instance._scrollTimer.Dispose();
            _instance._focusTimer.Dispose();
            _instance.Close();
            _instance.Dispose();
        }
        _instance = null;
    }

    // ---- 事件控制器（由 ZoomInOut_Click 调用）----

    private static Application? _app;

    public static void Enable(Application app)
    {
        _app = app;
        app.SheetSelectionChange += OnSelectionChange;
        app.WindowDeactivate += OnWindowDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
    }

    public static void Disable()
    {
        if (_app == null)
            return;
        _app.SheetSelectionChange -= OnSelectionChange;
        _app.WindowDeactivate -= OnWindowDeactivate;
        _app.WorkbookDeactivate -= OnWorkbookDeactivate;
        _app = null;
        Instance.ClearBubble();
    }

    private static void OnWindowDeactivate(object wb, object wn) => Instance.ClearBubble();

    private static void OnWorkbookDeactivate(object wb) => Instance.ClearBubble();

    public static void OnSelectionChange(object sh, Range target)
    {
        ExcelAsyncUtil.QueueAsMacro(() => TryShow(target));
    }

    private static void TryShow(Range target)
    {
        try
        {
            if (target.Rows.Count >= 100 || target.Columns.Count >= 10)
            {
                Instance.ClearBubble();
                return;
            }

            object rawVal = target.Value;
            if (rawVal == null)
            {
                Instance.ClearBubble();
                return;
            }

            string text;
            if (rawVal is object[,] arr)
            {
                var sb = new System.Text.StringBuilder();
                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    for (int j = 1; j <= arr.GetLength(1); j++)
                    {
                        if (j > 1)
                            sb.Append("  ");
                        sb.Append(arr[i, j]?.ToString() ?? "");
                    }
                    sb.AppendLine();
                }
                text = sb.ToString().TrimEnd();
            }
            else
                text = rawVal.ToString() ?? "";

            if (string.IsNullOrEmpty(text))
            {
                Instance.ClearBubble();
                return;
            }

            Instance.ShowBubble(text);
        }
        catch (Exception ex)
        {
            PluginLog.Verbose($"[CellTip] {ex.GetType().Name}: {ex.Message}");
            Instance.ClearBubble();
        }
    }

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WS_EX_NOACTIVATE = 0x08000000;
    private const int SW_SHOWNOACTIVATE = 4;

    [DllImport("user32.dll")]
    static extern int GetWindowLong(IntPtr h, int i);

    [DllImport("user32.dll")]
    static extern int SetWindowLong(IntPtr h, int i, int v);

    [DllImport("user32.dll")]
    static extern bool ShowWindow(IntPtr h, int cmd);

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
