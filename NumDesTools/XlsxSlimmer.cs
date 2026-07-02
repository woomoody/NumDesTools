using OfficeOpenXml;

namespace NumDesTools;

// 真正的 xlsx 体积膨胀源不是"样式重复"，而是"超出真实数据边界、仍带格式的空白行列尾部"
// （常见于误 Ctrl+A 设置过格式）。收缩这部分风险低、效果显著。
// 命名区域可能被公式/宏引用，本工具只诊断展示，不自动删除。
internal static class XlsxSlimmer
{
    internal sealed record SheetDiag(
        string Name,
        int UsedRows,
        int UsedCols,
        int TrueMaxRow,
        int TrueMaxCol,
        int CommentCount,
        int ConditionalFormatCount
    );

    internal sealed record DiagnoseResult(
        string FilePath,
        long OriginalBytes,
        List<SheetDiag> Sheets,
        int NamedRangeCount
    );

    internal static void Run()
    {
        var wb = AppServices.App.ActiveWorkbook;
        if (wb is null || string.IsNullOrEmpty(wb.Path))
        {
            MessageBox.Show("请先打开并保存一个 xlsx 文件。", "格式瘦身诊断");
            return;
        }

        new UI.XlsxSlimmerWindow(wb, Diagnose(wb.FullName)).Show();
    }

    private static DiagnoseResult Diagnose(string path)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var sheets = new List<SheetDiag>();
        int namedRangeCount;
        using (var pkg = new ExcelPackage(new FileInfo(path)))
        {
            namedRangeCount = pkg.Workbook.Names.Count;
            foreach (var ws in pkg.Workbook.Worksheets)
            {
                if (ws.Dimension is null)
                    continue;

                int trueMaxRow = 0,
                    trueMaxCol = 0;
                // ws.Cells 只枚举内部实际存在的 cell（稀疏存储），不会实例化整个矩形范围
                foreach (var cell in ws.Cells)
                {
                    if (cell.Value is null)
                        continue;
                    if (cell.Start.Row > trueMaxRow)
                        trueMaxRow = cell.Start.Row;
                    if (cell.Start.Column > trueMaxCol)
                        trueMaxCol = cell.Start.Column;
                }

                sheets.Add(
                    new SheetDiag(
                        ws.Name,
                        ws.Dimension.End.Row,
                        ws.Dimension.End.Column,
                        trueMaxRow,
                        trueMaxCol,
                        ws.Comments.Count,
                        ws.ConditionalFormatting.Count
                    )
                );
            }
        }
        return new DiagnoseResult(path, new FileInfo(path).Length, sheets, namedRangeCount);
    }

    // 全程只用 Excel COM 自己的 Save，文件句柄从头到尾都在 Excel 手里，不给别的进程留抢锁窗口
    // （之前 close→EPPlus 重开→Save 的方案在 close 和 EPPlus 重开之间有个空档，SmartGit/杀软等偶尔会抢到）。
    internal static void Slim(Workbook wb, DiagnoseResult diag)
    {
        foreach (var sd in diag.Sheets)
        {
            var ws = (Worksheet)wb.Worksheets[sd.Name];
            var used = ws.UsedRange;
            var endRow = used.Row + used.Rows.Count - 1;
            var endCol = used.Column + used.Columns.Count - 1;

            if (ws.Cells.FormatConditions.Count > 0)
                ws.Cells.FormatConditions.Delete();

            if (sd.TrueMaxRow > 0 && endRow > sd.TrueMaxRow)
                ((Range)ws.Rows[$"{sd.TrueMaxRow + 1}:{endRow}"]).Delete();

            if (sd.TrueMaxCol > 0 && endCol > sd.TrueMaxCol)
                ws.Range[ws.Cells[1, sd.TrueMaxCol + 1], ws.Cells[1, endCol]].EntireColumn.Delete();
        }

        try
        {
            wb.Save();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"保存失败：{ex.Message}", "格式瘦身");
            return;
        }

        var newSize = new FileInfo(diag.FilePath).Length;
        MessageBox.Show(
            $"瘦身完成（原地保存，git 可回溯）：\n原体积 {diag.OriginalBytes / 1024.0 / 1024.0:F1} MB\n新体积 {newSize / 1024.0 / 1024.0:F1} MB",
            "格式瘦身完成"
        );
    }
}
