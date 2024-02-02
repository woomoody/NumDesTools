using System;
using System.Drawing;
using System.Windows.Forms;
using Font = System.Drawing.Font;
using Range = Microsoft.Office.Interop.Excel.Range;
using IWin32Window = System.Windows.Forms.IWin32Window;

namespace NumDesTools;

public class CellSelectChangeTip : Form
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
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        ControlBox = false;
        Name = "CellSelectChange";
        DoubleBuffered = true;
        ForeColor = System.Drawing.Color.DimGray;
        ShowInTaskbar = false;
        // 设置窗体的属性，大小，位置等
        ClientSize = new Size(300, 200);
        FormBorderStyle = FormBorderStyle.None;
        BackColor = Color.Black;
        //TopMost = true;
        //Load += CellSelectChangeTip_Load;
        ResumeLayout(false);
    }
    private void CellSelectChangeTip_Paint(object sender, PaintEventArgs e)
    {
        // 在窗体上绘制文本
        using (var brush = new SolidBrush(Color.White))
        {
            e.Graphics.DrawString(_displayText, new Font("微软雅黑", 20), brush, new PointF(10, 10));
        }
    }
    public void ShowToolTip(string text, Range target)
    {
        // 更新文本内容
        _displayText = text;
        // 获取字体占的像素
        var size = TextRenderer.MeasureText(_displayText, new Font("微软雅黑", 20));
        var tipWidth = size.Width;
        var tipHeitht = size.Height;
        IntPtr excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
        var workingArea = NumDesAddIn.App.ActiveWindow;
        var baseLeft = (int)workingArea.Left + (int)workingArea.Width;
        var baseTop = (int)workingArea.Top + (int)workingArea.Height;
        _currentLeft = baseLeft - tipWidth - 25;
        _currentTop = baseTop - tipHeitht - 50;
        Location = new Point(_currentLeft, _currentTop);
        ClientSize = new Size(tipWidth, tipHeitht);
        //写入文本
        Paint += CellSelectChangeTip_Paint;
        Show(_owner);
    }

    public void HideToolTip()
    {
        Hide();
    }
    public IntPtr OwnerHandle
    {
        get
        {
            if (_owner == null)
                return IntPtr.Zero;
            return _owner.Handle;
        }
        set
        {
            if (_owner == null || _owner.Handle != value)
            {
                _owner = new Win32Window(value);
            }
        }
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
            var arr = target.Value2;

            foreach (var item in arr)
            {
                cellStr = cellStr+ item.ToString() +"\r\n";
            }
            //var gra = CreateGraphics();
            //var sF = gra.MeasureString(cellStr, new Font("微软雅黑", 20), 10000, StringFormat.GenericTypographic);
            // 显示或更新提示窗口
            ShowToolTip(cellStr, target);

            //// 释放资源
            //gra.Dispose();
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
    }
}