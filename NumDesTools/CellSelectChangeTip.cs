using Font = System.Drawing.Font;
using Range = Microsoft.Office.Interop.Excel.Range;
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

    //绘制窗口
    public CellSelectChangeTip()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        SuspendLayout();
        // 
        // CellSelectChangeTip
        // 
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        BackColor = Color.Black;
        ClientSize = new Size(300, 200);
        ControlBox = false;
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        Name = "CellSelectChangeTip";
        Load += CellSelectChangeTip_Load;
        ResumeLayout(false);
    }

    private void TargetStrWrite(object sender, PaintEventArgs e)
    {
        // 在窗体上绘制文本
#pragma warning disable CA1416
        using var brush = new SolidBrush(Color.White);
#pragma warning restore CA1416
#pragma warning disable CA1416
        e.Graphics.DrawString(_displayText, new Font("微软雅黑", 13), brush, new PointF(10, 10));
#pragma warning restore CA1416
    }

    public void ShowToolTip(string text, Range target)
    {
        // 更新文本内容
        _displayText = text;
        if (_displayText == null)
        {
            HideToolTip();
            return;
        }

        var workingArea = NumDesAddIn.App.ActiveWindow;
        var zoom = workingArea.Zoom / 100;
        var workingAreaLeft = workingArea.Left * 1.67;
        var workingAreaTop = workingArea.Top * 1.67;
        var workingAreaWidth = workingArea.Width * 1.67;
        var workingAreaHeight = workingArea.Height * 1.67;
        // 获取字体占的像素
#pragma warning disable CA1416
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
#pragma warning restore CA1416
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;
        // 获取单元格的工作区域坐标
        var targetLeftPixels = PubMetToExcel.ExcelRangePixelsX(target.Left * zoom);
        var targetWidthPixels = Convert.ToInt32(target.Width * 1.67 * zoom);
        var targetTopPixels = PubMetToExcel.ExcelRangePixelsY(target.Top * zoom);
        var targetHeightPixels = Convert.ToInt32(target.Height * 1.67 * zoom);
        _currentLeft = targetLeftPixels + targetWidthPixels;
        _currentTop = targetTopPixels + targetHeightPixels;
        if (_currentLeft + tipWidth > workingAreaLeft + workingAreaWidth) _currentLeft = targetLeftPixels - tipWidth;

        if (_currentTop + tipHeight > workingAreaTop + workingAreaHeight) _currentTop = targetTopPixels - tipHeight;
        //单位为像素
        Location = new System.Drawing.Point(_currentLeft, _currentTop);
        ClientSize = new Size(tipWidth, tipHeight);
        //写入文本
        Paint += TargetStrWrite;
        //获取工作区句柄,显示窗口，不使用次方法，窗口闪现
        var excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
        _owner = new Win32Window(excelHandle);
        Show(_owner);
    }

    public void HideToolTip()
    {
        Hide();
    }

    private class Win32Window : IWin32Window
    {
        public IntPtr Handle
        {
            get;
            // ReSharper disable once AutoPropertyCanBeMadeGetOnly.Local
            private set;
        }
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
            var arr = target.Value;
            var isArray = arr is object[,];
            if (isArray)
                for (var i = 1; i <= arr.GetLength(0); i++)
                {
                    for (var j = 1; j <= arr.GetLength(1); j++) cellStr += arr[i, j] + "#";
                    cellStr += "\r\n";
                }
            else
                cellStr = arr.ToString() + "\r\n";

            // 显示或更新提示窗口
            ShowToolTip(cellStr, target);
        }
        else
        {
            MessageBox.Show(@"选的格子太多了，重选" + @"\n" + @"最大99行，9列！");
            // 隐藏提示窗口
            HideToolTip();
        }
    }

    private void CellSelectChangeTip_Load(object sender, EventArgs e)
    {
        var target = NumDesAddIn.App.ActiveCell;
        // 获取字体占的像素
#pragma warning disable CA1416
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
#pragma warning restore CA1416
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;
        // 获取单元格的工作区域坐标
        var workingArea = NumDesAddIn.App.ActiveWindow;
        var zoom = workingArea.Zoom / 100;
        var workingAreaLeft = workingArea.Left * 1.67;
        var workingAreaTop = workingArea.Top * 1.67;
        var workingAreaWidth = workingArea.Width * 1.67;
        var workingAreaHeight = workingArea.Height * 1.67;
        var targetLeftPixels = PubMetToExcel.ExcelRangePixelsX(target.Left * zoom);
        var targetWidthPixels = Convert.ToInt32(target.Width * 1.67 * zoom);
        var targetTopPixels = PubMetToExcel.ExcelRangePixelsY(target.Top * zoom);
        var targetHeightPixels = Convert.ToInt32(target.Height * 1.67 * zoom);
        _currentLeft = targetLeftPixels + targetWidthPixels;
        _currentTop = targetTopPixels + targetHeightPixels;

        if (_currentLeft + tipWidth > workingAreaLeft + workingAreaWidth) _currentLeft = targetLeftPixels - tipWidth;

        if (_currentTop + tipHeight > workingAreaTop + workingAreaHeight) _currentTop = targetTopPixels - tipHeight;
        Location = new System.Drawing.Point(_currentLeft, _currentTop);
        ClientSize = new Size(tipWidth, tipHeight);
    }
}

//点击穿透界面
public class ClickThroughForm : Form
{
    // ReSharper disable once InconsistentNaming
    // ReSharper disable once IdentifierTypo
    private const int WM_NCHITTEST = 0x84;

    // ReSharper disable once IdentifierTypo
    // ReSharper disable once InconsistentNaming
    private const int HTTRANSPARENT = -1;

    public ClickThroughForm()
    {
        // 允许窗体透明度
        SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        // 设置窗体为透明
        // ReSharper disable once VirtualMemberCallInConstructor
        BackColor = Color.Transparent;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
            // ReSharper disable once CommentTypo
            // 如果鼠标在窗体区域，返回 HTTRANSPARENT，表示鼠标事件透传到下一层控件
            m.Result = (IntPtr)HTTRANSPARENT;
            return;
        }

        base.WndProc(ref m);
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var createParams = base.CreateParams;
            createParams.ExStyle |= 0x00000020; // WS_EX_TRANSPARENT
            return createParams;
        }
    }
}