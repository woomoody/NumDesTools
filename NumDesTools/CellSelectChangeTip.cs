using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Font = System.Drawing.Font;
using IWin32Window = System.Windows.Forms.IWin32Window;

#pragma warning disable CA1416
#pragma warning disable CA1416


namespace NumDesTools;

public class CellSelectChangeTip : ClickThroughForm
{
    private string _displayText;
    private int _currentLeft;
    private int _currentTop;

    private Win32Window _owner;

    public CellSelectChangeTip()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        SuspendLayout();
        AutoScaleMode = AutoScaleMode.None;   // 禁止 WinForms 按字体DPI缩放 Location 坐标
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        BackColor = Color.Black;
        ClientSize = new Size(300, 200);
        ControlBox = false;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        StartPosition = FormStartPosition.Manual;  // 防止 Show() 重置 Location
        Name = "CellSelectChangeTip";
        Load += CellSelectChangeTip_Load;
        ResumeLayout(false);
    }

    private void TargetStrWrite(object sender, PaintEventArgs e)
    {
#pragma warning disable CA1416
        using var brush = new SolidBrush(Color.White);
#pragma warning restore CA1416
#pragma warning disable CA1416
        e.Graphics.DrawString(_displayText, new Font("微软雅黑", 13), brush, new PointF(10, 10));
#pragma warning restore CA1416
    }

    public void ShowToolTip(string text, Range target)
    {
        _displayText = text;
        if (_displayText == null)
        {
            HideToolTip();
            return;
        }

#pragma warning disable CA1416
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
#pragma warning restore CA1416
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;

        if (!CalcScreenPos(target, ref tipWidth, ref tipHeight, out _currentLeft, out _currentTop))
            return;

        Paint -= TargetStrWrite;
        Paint += TargetStrWrite;
        var excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
        _owner = new Win32Window(excelHandle);
        if (!Visible)
            Show(_owner);
        ClientSize = new Size(tipWidth, tipHeight);
        Location = new Point(_currentLeft, _currentTop);
        GetWindowRect(Handle, out var rc);
        var scr = Screen.FromHandle(Handle);
        System.Diagnostics.Debug.WriteLine(
            $"[CellTip] setLoc=({_currentLeft},{_currentTop}) → WinRect=({rc.Left},{rc.Top},{rc.Right},{rc.Bottom})" +
            $" | DeviceDpi={DeviceDpi} ScreenBounds={scr.Bounds} WorkingArea={scr.WorkingArea}");
        Invalidate();
    }

    public void HideToolTip()
    {
        Hide();
    }

    // COM Range.Left/Top + PointsToScreenPixels，再乘 DPI scale 转物理像素给 SetWindowPos
    private static bool CalcScreenPos(Range target, ref int tipWidth, ref int tipHeight,
        out int left, out int top)
    {
        left = top = 0;
        try
        {
            var win = NumDesAddIn.App.ActiveWindow;

            // PointsToScreenPixels 与 Form.Location / SetWindowPos 坐标系一致，直接使用
            int scrLeft   = win.PointsToScreenPixelsX((int)target.Left);
            int scrTop    = win.PointsToScreenPixelsY((int)target.Top);
            int scrRight  = win.PointsToScreenPixelsX((int)(target.Left + target.Width));
            int scrBottom = win.PointsToScreenPixelsY((int)(target.Top  + target.Height));

            left = scrRight;
            top  = scrBottom;

            var vs = SystemInformation.VirtualScreen;
            if (left + tipWidth  > vs.Right)  left = scrLeft - tipWidth;
            if (top  + tipHeight > vs.Bottom) top  = scrTop  - tipHeight;

            System.Diagnostics.Debug.WriteLine(
                $"[CellTip] addr={target.Address}" +
                $" | scr L={scrLeft} T={scrTop} R={scrRight} B={scrBottom}" +
                $" | finalPos ({left},{top}) size ({tipWidth}x{tipHeight})");
            return true;
        }
        catch
        {
            return false;
        }
    }

    private class Win32Window : IWin32Window
    {
        public IntPtr Handle { get; private set; }
#pragma warning disable IDE0290
        public Win32Window(IntPtr handle)
#pragma warning restore IDE0290
        {
            Handle = handle;
        }
    }

    public void GetCellValue(object sh, Range target)
    {
        HideToolTip();
        var rngRow = target.Rows.Count;
        var rngCol = target.Columns.Count;

        if (rngRow < 100 && rngCol < 10)
        {
            var cellStr = "";
            object rawVal = target.Value;
            if (rawVal is object[,] arr)
            {
                for (var i = 1; i <= arr.GetLength(0); i++)
                {
                    for (var j = 1; j <= arr.GetLength(1); j++)
                        cellStr += (arr[i, j]?.ToString() ?? string.Empty) + "#";
                    cellStr += "\r\n";
                }
            }
            else
                cellStr = (rawVal?.ToString() ?? string.Empty) + "\r\n";

            ShowToolTip(cellStr, target);
        }
        else
        {
            MessageBox.Show(@"选的格子太多了，重选" + @"\n" + @"最大99行，9列！");
            HideToolTip();
        }
    }

    private void CellSelectChangeTip_Load(object sender, EventArgs e)
    {
        var target = NumDesAddIn.App.ActiveCell;
#pragma warning disable CA1416
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
#pragma warning restore CA1416
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;

        if (!CalcScreenPos(target, ref tipWidth, ref tipHeight, out _currentLeft, out _currentTop))
            return;

        ClientSize = new Size(tipWidth, tipHeight);
        Location = new Point(_currentLeft, _currentTop);
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hWnd);
    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
        int x, int y, int cx, int cy, uint uFlags);
    private const uint SWP_NOZORDER   = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }
}

public class ClickThroughForm : Form
{
    private const int WmNchittest = 0x84;

    private const int Httransparent = -1;

    public ClickThroughForm()
    {
        SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WmNchittest)
        {
            m.Result = (IntPtr)Httransparent;
            return;
        }

        base.WndProc(ref m);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var createParams = base.CreateParams;
            createParams.ExStyle |= 0x00000020;
            return createParams;
        }
    }
}
