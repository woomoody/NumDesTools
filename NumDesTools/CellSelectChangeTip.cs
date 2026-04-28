using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Font = System.Drawing.Font;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 跟随光标的单元格值气泡提示，小窗口方案。
/// 定位：Cursor.Position 直接作为 Form.Location，两者都是 WinForms 物理坐标，无需换算。
/// </summary>
public sealed class CellSelectChangeTip : Form
{
    private string? _text;
    private static readonly Font TipFont = new Font("微软雅黑", 11);
    private const int Pad = 8;

    private static CellSelectChangeTip? _instance;
    public  static CellSelectChangeTip  Instance => _instance ??= new CellSelectChangeTip();

    private CellSelectChangeTip()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar   = false;
        TopMost         = true;
        BackColor       = Color.FromArgb(40, 40, 40);
        ForeColor       = Color.White;
        AutoScaleMode   = AutoScaleMode.None;
        StartPosition   = FormStartPosition.Manual;

        SetStyle(ControlStyles.OptimizedDoubleBuffer
               | ControlStyles.AllPaintingInWmPaint
               | ControlStyles.UserPaint, true);

        // 穿透鼠标点击
        var ex = GetWindowLong(Handle, GWL_EXSTYLE);
        SetWindowLong(Handle, GWL_EXSTYLE, ex | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        if (_text == null) return;
        using var brush = new SolidBrush(ForeColor);
        e.Graphics.DrawString(_text, TipFont, brush, new PointF(Pad, Pad));
    }

    public void ShowBubble(string text)
    {
        _text = text;
        var sz = TextRenderer.MeasureText(text, TipFont);
        int w  = sz.Width  + Pad * 2;
        int h  = sz.Height + Pad * 2;

        // Cursor.Position 与 Form.Location 同为 WinForms 物理坐标，直接相加偏移 14px
        var cursor = Cursor.Position;
        int x = cursor.X + 14;
        int y = cursor.Y + 14;

        // 边界检测：用当前光标所在屏幕的工作区
        var wa = Screen.FromPoint(cursor).WorkingArea;
        if (x + w > wa.Right)  x = cursor.X - w - 2;
        if (y + h > wa.Bottom) y = cursor.Y - h - 2;
        if (x < wa.Left) x = wa.Left;
        if (y < wa.Top)  y = wa.Top;

        ClientSize = new Size(w, h);
        Location   = new Point(x, y);

        GetWindowRect(Handle, out var rc);
        System.Diagnostics.Debug.WriteLine(
            $"[CellTip] cursor=({cursor.X},{cursor.Y}) → loc=({x},{y}) WinRect=({rc.Left},{rc.Top})");

        if (!Visible) Show();
        Invalidate();
    }

    public void ClearBubble()
    {
        if (Visible) Hide();
    }

    public static void DisposeInstance()
    {
        if (_instance is { IsDisposed: false })
        {
            _instance.Close();
            _instance.Dispose();
        }
        _instance = null;
    }

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
            if (rawVal == null) { Instance.ClearBubble(); return; }

            string text;
            if (rawVal is object[,] arr)
            {
                var sb = new System.Text.StringBuilder();
                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    for (int j = 1; j <= arr.GetLength(1); j++)
                    {
                        if (j > 1) sb.Append("  ");
                        sb.Append(arr[i, j]?.ToString() ?? "");
                    }
                    sb.AppendLine();
                }
                text = sb.ToString().TrimEnd();
            }
            else
                text = rawVal.ToString() ?? "";

            if (string.IsNullOrEmpty(text)) { Instance.ClearBubble(); return; }

            Instance.ShowBubble(text);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[CellTip] {ex.GetType().Name}: {ex.Message}");
            Instance.ClearBubble();
        }
    }

    private const int GWL_EXSTYLE       = -20;
    private const int WS_EX_TRANSPARENT = 0x00000020;
    private const int WS_EX_NOACTIVATE  = 0x08000000;

    [DllImport("user32.dll")] static extern int  GetWindowLong(IntPtr h, int i);
    [DllImport("user32.dll")] static extern int  SetWindowLong(IntPtr h, int i, int v);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [StructLayout(LayoutKind.Sequential)]
    private struct RECT { public int Left, Top, Right, Bottom; }
}
