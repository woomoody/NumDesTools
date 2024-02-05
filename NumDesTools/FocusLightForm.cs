using System;
using System.Drawing;
using System.Windows.Forms;
using IWin32Window = System.Windows.Forms.IWin32Window;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace NumDesTools
{
    public  class FocusLightForm : Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }
        Win32Window _owner;
        private Color borderColor = Color.IndianRed; // 根据需要设置边框颜色
        private int borderWidth = 5; // 根据需要设置边框宽度
        private int bottomPadding = 5; // 根据需要设置底部间距
        private void InitializeComponent()
        {
            SuspendLayout();
            // 
            // FocusLightForm
            // 
            AutoScaleMode = AutoScaleMode.Inherit;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BackColor = Color.Magenta;
            TransparencyKey = BackColor;
            ClientSize = new Size(282, 153);
            ControlBox = false;
            FormBorderStyle = FormBorderStyle.None;
            // 增加底部间距
            Padding = new Padding(0, 0, 0, bottomPadding);
            Name = "FocusLightForm";
            ShowIcon = false;
            ShowInTaskbar = false;
            Paint += MainForm_Paint;
            Load += FocusLightForm_Load;
            ResumeLayout(false);
        }
        private void MainForm_Paint(object sender, PaintEventArgs e)
        {
            // 绘制窗体边框
            using (Pen borderPen = new Pen(borderColor, borderWidth))
            {
                e.Graphics.DrawRectangle(borderPen, new Rectangle(0, 0, this.Width - 1, this.Height - 1));
            }
        }
        public FocusLightForm()
        {
            InitializeComponent();
        }
        public void ShowToolTip(object sh, Range target)
        {
            var workingArea = NumDesAddIn.App.ActiveWindow;
            var workingAreaLeft = workingArea.Left * 1.67;
            var workingAreaTop = workingArea.Top * 1.67;
            var workingAreaWidth = workingArea.Width * 1.67;
            var workingAreaHeight = workingArea.Height * 1.67;

            // 获取单元格的工作区域坐标
            var targetLeftPixels = PubMetToExcel.ExcelRangePixelsX(target.Left);
            var targetWidthPixels = Convert.ToInt32(target.Width * 1.67);
            var targetTopPixels = PubMetToExcel.ExcelRangePixelsY(target.Top);
            var targetHeightPixels = Convert.ToInt32(target.Height * 1.67);

            int currentRowLeft;
            int currentRowTop;
            int currentRowWidth;
            int currentRowHeight;

            currentRowLeft = (int)workingAreaLeft;
            currentRowTop = targetTopPixels;
            currentRowWidth = (int)workingAreaWidth;
            currentRowHeight = targetHeightPixels;

            Location = new Point(currentRowLeft, currentRowTop);
            ClientSize = new Size(currentRowWidth, currentRowHeight);

            int currentColumnLeft;
            int currentColumnTop;
            int currentColumnWidth;
            int currentColumnHeight;

            currentColumnLeft = targetLeftPixels;
            currentColumnTop = (int)workingAreaTop;
            currentColumnWidth = targetWidthPixels;
            currentColumnHeight = (int)workingAreaHeight;



            //获取工作区句柄,显示窗口，不使用次方法，窗口闪现
            IntPtr excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
            _owner = new Win32Window(excelHandle);
            Show(_owner);
        }
        public void HideToolTip()
        {
            Hide();
        }
        public void ShowToolTip()
        {
            //获取工作区句柄,显示窗口，不使用次方法，窗口闪现
            IntPtr excelHandle = (IntPtr)NumDesAddIn.App.Hwnd;
            _owner = new Win32Window(excelHandle);
            Show(_owner);
        }

        private void FocusLightForm_Load(object sender, EventArgs e)
        {

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
    }
}
