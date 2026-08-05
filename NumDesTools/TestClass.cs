namespace NumDesTools;

using ExcelDna.Integration.CustomUI;
using NumDesTools.UI;

/// <summary>主题调试命令（临时）：通过 Application.Run 后台触发，无需点击 Ribbon。</summary>
public static class ThemeDebugCommands
{
    [ExcelCommand]
    public static void DebugOpenSheetCTP()
    {
        var ctpName = "表格目录";
        NumDesCTP.DeleteCTP(true, ctpName);
        NumDesCTP.ShowCTP(
            400,
            ctpName,
            true,
            ctpName,
            new SheetListControl(),
            MsoCTPDockPosition.msoCTPDockPositionLeft
        );
    }

    [ExcelCommand]
    public static void DebugThemeDark() => ThemeService.SetMode(ThemeService.ThemeMode.Dark);

    [ExcelCommand]
    public static void DebugThemeLight() => ThemeService.SetMode(ThemeService.ThemeMode.Light);

    [ExcelCommand]
    public static void DebugOpenChatCTP()
    {
        var name = "AI对话-Excel";
        NumDesCTP.DeleteCTP(true, name);
        NumDesCTP.ShowCTP(
            1500,
            name,
            true,
            name,
            new AiChatTaskPanel(),
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }

    [ExcelCommand]
    public static void DebugOpenAgentCTP()
    {
        var name = "AI Agent-Excel";
        NumDesCTP.DeleteCTP(true, name);
        NumDesCTP.ShowCTP(
            1500,
            name,
            true,
            name,
            new AIAgentPanel(),
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }

    [ExcelCommand]
    public static void DebugOpenSearchCTP()
    {
        var name = "搜索结果";
        NumDesCTP.DeleteCTP(true, name);
        NumDesCTP.ShowCTP(
            400,
            name,
            true,
            name,
            new SheetSeachResult([]),
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }

    [ExcelCommand]
    public static void DebugOpenImageCTP()
    {
        var name = "图片预览";
        NumDesCTP.DeleteCTP(true, name);
        NumDesCTP.ShowCTP(
            400,
            name,
            true,
            name,
            new ImagePreviewControl(new Dictionary<string, List<string>>()),
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }

    [ExcelCommand]
    public static void DebugOpenBatchReplaceCTP()
    {
        var name = "批量替换";
        NumDesCTP.DeleteCTP(true, name);
        NumDesCTP.ShowCTP(
            400,
            name,
            true,
            name,
            new BatchReplacePanel(),
            MsoCTPDockPosition.msoCTPDockPositionRight
        );
    }
}

public class ScreenCoordinateFix
{
    // Windows API
    [DllImport("user32.dll")]
    private static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

    [StructLayout(LayoutKind.Sequential)]
    public struct Point
    {
        public int X;
        public int Y;
    }

    [ExcelCommand]
    public static void GetCorrectScreenCoordinates()
    {
        var excelApp = AppServices.App;
        Range range = excelApp.Selection as Range;
        if (range == null)
            return;

        IntPtr hwnd = (IntPtr)excelApp.Hwnd;
        Window window = excelApp.ActiveWindow;

        // Step 1: 获取客户区原点屏幕坐标
        Point clientOrigin = new Point { X = 0, Y = 0 };
        ClientToScreen(hwnd, ref clientOrigin);

        // Step 2: 获取DPI（精确到当前显示器）
        float dpiX,
            dpiY;
        using (Graphics g = Graphics.FromHwnd(hwnd))
        {
            dpiX = g.DpiX;
            dpiY = g.DpiY;
        }

        // Step 3: 计算滚动偏移（改进方法）
        Range visibleStart = window.VisibleRange.Cells[1, 1];
        double scrollOffsetX = visibleStart.Left;
        double scrollOffsetY = visibleStart.Top;

        // Step 4: 处理缩放比例
        double zoomFactor = window.Zoom / 100.0;

        // Step 5: 计算最终坐标
        double adjustedLeft = (range.Left - scrollOffsetX) * zoomFactor;
        double adjustedTop = (range.Top - scrollOffsetY) * zoomFactor;

        int screenX = clientOrigin.X + (int)(adjustedLeft * dpiX / 72.0);
        int screenY = clientOrigin.Y + (int)(adjustedTop * dpiY / 72.0);

        // Step 6: 处理行列标题偏移
        screenX += GetColumnHeaderWidth(window); // 添加列标题宽度
        screenY += GetRowHeaderHeight(window); // 添加行标题高度

        // 在消息框中显示中间值
        MessageBox.Show(
            $"ClientOrigin: ({clientOrigin.X},{clientOrigin.Y})\n"
                + $"DPI: {dpiX}x{dpiY}\n"
                + $"ScrollOffset: {scrollOffsetX},{scrollOffsetY}\n"
                + $"Zoom: {window.Zoom}%\n"
                + $"修正后坐标: ({screenX}, {screenY})"
        );
    }

    private static int GetColumnHeaderWidth(Window window)
    {
        // 列标题区域宽度（默认37像素，可根据实际调整）
        return window.DisplayHeadings ? 37 : 0;
    }

    private static int GetRowHeaderHeight(Window window)
    {
        // 行标题区域高度（默认20像素，可根据实际调整）
        return window.DisplayHeadings ? 20 : 0;
    }
}
