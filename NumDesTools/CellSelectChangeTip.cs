using System;
using System.Drawing;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Range = Microsoft.Office.Interop.Excel.Range;
using IWin32Window = System.Windows.Forms.IWin32Window;


namespace NumDesTools;

public class CellSelectChangeTip : ClickThroughForm
{
    private string _displayText;
    int _currentLeft;
    int _currentTop;
    Win32Window _owner;
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
        //AutoSizeMode = AutoSizeMode.GrowAndShrink;
        BackColor = Color.Black;
        ClientSize = new Size(300, 200);
        ControlBox = false;
        //DoubleBuffered = true;
        //ForeColor = Color.DimGray;
        FormBorderStyle = FormBorderStyle.None;
        Name = "CellSelectChangeTip";
        ShowInTaskbar = false;
        Load += CellSelectChangeTip_Load;
        ResumeLayout(false);
    }

    private void TargetStrWrite(object sender, PaintEventArgs e)
    {
        // 在窗体上绘制文本
        using var brush = new SolidBrush(Color.White);
        e.Graphics.DrawString(_displayText, new Font("微软雅黑", 13), brush, new PointF(10, 10));
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
        var workingAreaLeft = workingArea.Left * 1.67;
        var workingAreaTop  = workingArea.Top * 1.67;
        var workingAreaWidth = workingArea.Width * 1.67;
        var workingAreaHeight = workingArea.Height * 1.67;
        // 获取字体占的像素
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;
        // 获取单元格的工作区域坐标
        var targetLeftPixels = PubMetToExcel.ExcelRangePixelsX(target.Left);
        var targetWidthPixels = Convert.ToInt32(target.Width * 1.67);
        var targetTopPixels = PubMetToExcel.ExcelRangePixelsY(target.Top);
        var targetHeightPixels = Convert.ToInt32(target.Height * 1.67);
        _currentLeft = targetLeftPixels + targetWidthPixels;
        _currentTop = targetTopPixels + targetHeightPixels;
        if (_currentLeft + tipWidth > workingAreaLeft + workingAreaWidth)
        {
            _currentLeft = targetLeftPixels - tipWidth;
        }

        if (_currentTop + tipHeight > workingAreaTop + workingAreaHeight)
        {
            _currentTop = targetTopPixels - tipHeight;
        }
        //单位为像素
        Location = new Point(_currentLeft, _currentTop);
        ClientSize = new Size(tipWidth, tipHeight);
        //写入文本
        Paint += TargetStrWrite;
        //获取工作区句柄,显示窗口，不使用次方法，窗口闪现
        IntPtr excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
        _owner = new Win32Window(excelHandle);
        Show(_owner);
    }

    public void HideToolTip()
    {
        Hide();
    }
    class Win32Window : IWin32Window
    {
        public IntPtr Handle
        {
            get;
            private set;
        }
        public Win32Window(IntPtr handle)
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
            var arr = target.get_Value();

            if (arr is object[,] arrayValue)
            {
                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    for (int j = 1; j <= arr.GetLength(1); j++)
                    {
                        cellStr += arr[i, j] + "#";
                    }
                    cellStr += "\r\n";
                }
            }
            else
            {
                cellStr = arr;
            }
          
      

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
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 13));
        var tipWidth = size.Width + 10;
        var tipHeight = size.Height + 10;
        // 获取单元格的工作区域坐标
        var workingArea = NumDesAddIn.App.ActiveWindow;
        var workingAreaLeft = workingArea.Left * 1.67;
        var workingAreaTop = workingArea.Top * 1.67;
        var workingAreaWidth = workingArea.Width * 1.67;
        var workingAreaHeight = workingArea.Height * 1.67;
        var targetLeftPixels = PubMetToExcel.ExcelRangePixelsX(target.Left);
        var targetWidthPixels = Convert.ToInt32(target.Width * 1.67);
        var targetTopPixels = PubMetToExcel.ExcelRangePixelsY(target.Top);
        var targetHeightPixels = Convert.ToInt32(target.Height * 1.67);
        _currentLeft = targetLeftPixels + targetWidthPixels;
        _currentTop = targetTopPixels + targetHeightPixels;

        if (_currentLeft + tipWidth > workingAreaLeft + workingAreaWidth)
        {
            _currentLeft = targetLeftPixels - tipWidth;
        }

        if (_currentTop + tipHeight > workingAreaTop + workingAreaHeight)
        {
            _currentTop = targetTopPixels - tipHeight;
        }
        Location = new Point(_currentLeft, _currentTop);
    }
}
//点击穿透界面
public class ClickThroughForm : Form
{
    private const int WM_NCHITTEST = 0x84;
    private const int HTTRANSPARENT = -1;

    public ClickThroughForm()
    {
        // 允许窗体透明度
        SetStyle(ControlStyles.SupportsTransparentBackColor, true);
        // 设置窗体为透明
        BackColor = Color.Transparent;
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_NCHITTEST)
        {
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
            CreateParams createParams = base.CreateParams;
            createParams.ExStyle |= 0x00000020; // WS_EX_TRANSPARENT
            return createParams;
        }
    }
}