using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Rectangle = System.Drawing.Rectangle;
using Timer = System.Windows.Forms.Timer;

namespace NumDesTools;

/// <summary>
/// 捕捉当前选中单元格的屏幕像素矩形，通过闪烁红框可视化验证坐标准确性。
/// </summary>
internal static class CellPositionProbe
{
    // ── 坐标获取 ──────────────────────────────────────────────────────────────

    /// <summary>
    /// 获取单元格的屏幕物理像素矩形。
    /// Range.Left/Top 单位是 Points（从 A1 起量），ActivePane.PointsToScreenPixelsX/Y
    /// 将其转换为绝对屏幕像素，内部已折合 DPI 和 Excel 视图缩放，无需手动乘系数。
    /// </summary>
    public static Rectangle GetCellScreenRect(Range cell)
    {
        try
        {
            var anchor = cell.Areas.Count > 1 ? (Range)cell.Areas[1] : cell;

            // ActivePane 的坐标原点与 Range.Left/Top 对齐（均从 A1 网格起量）
            // ActiveWindow.PointsToScreenPixelsX 原点含行列标头区，不可用
            dynamic pane = AppServices.App.ActiveWindow.ActivePane;

            var x1 = (int)pane.PointsToScreenPixelsX((double)anchor.Left);
            var y1 = (int)pane.PointsToScreenPixelsY((double)anchor.Top);
            var x2 = (int)pane.PointsToScreenPixelsX((double)(anchor.Left + anchor.Width));
            var y2 = (int)pane.PointsToScreenPixelsY((double)(anchor.Top + anchor.Height));

            return Rectangle.FromLTRB(x1, y1, x2, y2);
        }
        catch
        {
            return Rectangle.Empty;
        }
    }

    public static string GetCellScreenRectDebug(Range cell)
    {
        try
        {
            var anchor = cell.Areas.Count > 1 ? (Range)cell.Areas[1] : cell;
            var rect = GetCellScreenRect(cell);
            if (rect == Rectangle.Empty)
                return "ERR: Rectangle.Empty";

            dynamic pane = AppServices.App.ActiveWindow.ActivePane;
            int paneIndex = (int)pane.Index;

            return $"left={rect.Left} top={rect.Top} w={rect.Width} h={rect.Height} pane={paneIndex}"
                + $"  pts=({(double)anchor.Left:F1},{(double)anchor.Top:F1}"
                + $" {(double)anchor.Width:F1}x{(double)anchor.Height:F1})";
        }
        catch (Exception ex)
        {
            return $"ERR: {ex.Message}";
        }
    }

    // ── 视觉验证 ──────────────────────────────────────────────────────────────

    // FlashBorderForm 在独立 STA 线程上运行，不依赖 Excel-DNA 主线程的消息泵
    private static Thread? _staThread;
    private static FlashBorderForm? _flashForm;
    private static readonly object _flashLock = new();

    /// <summary>
    /// 在计算出的单元格屏幕位置画亮红色边框，1.5 秒后自动消失。
    /// 红框精确贴合单元格四条边则坐标正确，偏移或尺寸不对则一眼可见。
    /// </summary>
    public static void FlashCellBorder(Rectangle rect)
    {
        if (rect == Rectangle.Empty)
            return;

        lock (_flashLock)
        {
            // 若 STA 线程还没启动，创建一次
            if (_staThread is null || !_staThread.IsAlive)
            {
                _staThread = new Thread(() =>
                {
                    _flashForm = new FlashBorderForm();
                    // Application.Run 阻塞并处理消息，直到 _flashForm 调用 Application.ExitThread()
                    System.Windows.Forms.Application.Run();
                })
                {
                    IsBackground = true,
                    Name = "CellProbe-STA",
                };
                _staThread.SetApartmentState(ApartmentState.STA);
                _staThread.Start();
                // 等 _flashForm 初始化完成
                var sw = System.Diagnostics.Stopwatch.StartNew();
                while (_flashForm is null && sw.ElapsedMilliseconds < 2000)
                    Thread.Sleep(10);
            }
        }

        // 把 Flash 调用 marshal 到 STA 线程
        _flashForm?.BeginInvoke((System.Action)(() => _flashForm.Flash(rect)));
    }

    // ── 隐藏 UDF ──────────────────────────────────────────────────────────────

    [ExcelFunction(
        Name = "ExcelCellScreenRect",
        Description = "【诊断】返回当前活动单元格的屏幕像素矩形 \"x,y,w,h\"，同时触发闪烁红框视觉验证。",
        IsVolatile = true,
        IsMacroType = true,
        IsHidden = true
    )]
    public static string ExcelCellScreenRect()
    {
        try
        {
            var app = AppServices.App;
            if (app.ActiveCell is not Range cell)
                return "ERR: no active cell";

            var rect = GetCellScreenRect(cell);
            if (rect == Rectangle.Empty)
                return "ERR: Rectangle.Empty";

            // rect 已是纯值类型，可安全跨线程传递
            var rectCopy = rect;
            Task.Run(() => FlashCellBorder(rectCopy));

            var zoom = (int)(double)app.ActiveWindow.Zoom;
            return $"{rect.Left},{rect.Top},{rect.Width},{rect.Height} [{cell.Address[false, false]} zoom={zoom}%]";
        }
        catch (Exception ex)
        {
            return $"ERR: {ex.GetType().Name}: {ex.Message}";
        }
    }

    // ── 闪烁边框窗口 ──────────────────────────────────────────────────────────

    private sealed class FlashBorderForm : Form
    {
        private const int BorderThickness = 3;
        private const int DisplayMs = 1500;
        private const int GwlExstyle = -20;
        private const int WsExLayered = 0x00080000;
        private const int WsExTransparent = 0x00000020;
        private const int WsExToolwindow = 0x00000080;
        private const int WsExNoactivate = 0x08000000;
        private const int LwaColorkey = 0x1;
        private const int LwaAlpha = 0x2;
        private const int SwShownoactivate = 4;

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

        private readonly Timer _hideTimer;

        public FlashBorderForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Magenta;
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.Manual;

            var ex = GetWindowLong(Handle, GwlExstyle);
            SetWindowLong(
                Handle,
                GwlExstyle,
                ex | WsExLayered | WsExTransparent | WsExToolwindow | WsExNoactivate
            );
            SetLayeredWindowAttributes(
                Handle,
                (uint)ColorTranslator.ToWin32(Color.Magenta),
                220,
                LwaColorkey | LwaAlpha
            );

            _hideTimer = new Timer { Interval = DisplayMs };
            _hideTimer.Tick += (_, _) =>
            {
                _hideTimer.Stop();
                Hide();
            };
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WsExLayered | WsExTransparent | WsExToolwindow | WsExNoactivate;
                return cp;
            }
        }

        public void Flash(Rectangle cellRect)
        {
            var outer = Rectangle.Inflate(cellRect, BorderThickness, BorderThickness);
            Location = new Point(outer.Left, outer.Top);
            ClientSize = new Size(outer.Width, outer.Height);

            _hideTimer.Stop();
            Invalidate();
            ShowWindow(Handle, SwShownoactivate);
            _hideTimer.Start();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.Clear(Color.Magenta);
            using var pen = new Pen(Color.Red, BorderThickness);
            var borderRect = new Rectangle(
                BorderThickness / 2,
                BorderThickness / 2,
                ClientSize.Width - BorderThickness,
                ClientSize.Height - BorderThickness
            );
            e.Graphics.DrawRectangle(pen, borderRect);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _hideTimer.Dispose();
            base.Dispose(disposing);
        }
    }
}
