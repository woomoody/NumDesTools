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

        ShowDiagnosis(wb, Diagnose(wb.FullName));
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

    private static void Slim(Workbook wb, DiagnoseResult diag)
    {
        // Saved 可能因 Excel 打开时重算公式/易变函数而变 false，未必是用户真的改了内容；
        // 直接保存再关闭比拦住用户手动 Ctrl+S 更顺畅，且同样安全（没改动=空操作）。
        if (!wb.Saved)
            wb.Save();

        var path = diag.FilePath;
        wb.Close(false);

        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using (var pkg = new ExcelPackage(new FileInfo(path)))
        {
            foreach (var sd in diag.Sheets)
            {
                var ws = pkg.Workbook.Worksheets[sd.Name];
                if (ws.Dimension is null)
                    continue;

                if (ws.ConditionalFormatting.Count > 0)
                    ws.ConditionalFormatting.RemoveAll();

                if (sd.TrueMaxRow > 0 && ws.Dimension.End.Row > sd.TrueMaxRow)
                    ws.DeleteRow(sd.TrueMaxRow + 1, ws.Dimension.End.Row - sd.TrueMaxRow);

                if (sd.TrueMaxCol > 0 && ws.Dimension.End.Column > sd.TrueMaxCol)
                    ws.DeleteColumn(sd.TrueMaxCol + 1, ws.Dimension.End.Column - sd.TrueMaxCol);
            }
            pkg.Save();
        }

        AppServices.App.Workbooks.Open(path);

        var newSize = new FileInfo(path).Length;
        MessageBox.Show(
            $"瘦身完成（原地保存，git 可回溯）：\n原体积 {diag.OriginalBytes / 1024.0 / 1024.0:F1} MB\n新体积 {newSize / 1024.0 / 1024.0:F1} MB",
            "格式瘦身完成"
        );
    }

    private static void ShowDiagnosis(Workbook wb, DiagnoseResult diag)
    {
        var form = new Form
        {
            Text = $"格式瘦身诊断 — {Path.GetFileName(diag.FilePath)}",
            Width = 720,
            Height = 520,
            MinimumSize = new System.Drawing.Size(640, 420),
            StartPosition = FormStartPosition.CenterScreen,
            KeyPreview = true,
            Padding = new Padding(12),
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F),
        };
        form.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                form.Close();
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 3,
            ColumnCount = 1,
        };
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 10),
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            BackgroundColor = System.Drawing.SystemColors.Window,
            BorderStyle = BorderStyle.Fixed3D,
            RowHeadersVisible = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            ColumnHeadersDefaultCellStyle =
            {
                Font = new System.Drawing.Font(
                    "Microsoft YaHei UI",
                    9F,
                    System.Drawing.FontStyle.Bold
                ),
            },
        };
        grid.Columns.Add("Sheet", "Sheet");
        grid.Columns.Add("Used", "当前范围(行x列)");
        grid.Columns.Add("True", "真实数据范围(行x列)");
        grid.Columns.Add("Waste", "冗余行/列");
        grid.Columns.Add("Comment", "批注数");
        grid.Columns.Add("CF", "条件格式");

        var hasWaste = false;
        foreach (var sd in diag.Sheets)
        {
            var wasteRows = sd.UsedRows - sd.TrueMaxRow;
            var wasteCols = sd.UsedCols - sd.TrueMaxCol;
            if (wasteRows > 0 || wasteCols > 0 || sd.ConditionalFormatCount > 0)
                hasWaste = true;
            grid.Rows.Add(
                sd.Name,
                $"{sd.UsedRows}x{sd.UsedCols}",
                $"{sd.TrueMaxRow}x{sd.TrueMaxCol}",
                $"{wasteRows}行/{wasteCols}列",
                sd.CommentCount,
                sd.ConditionalFormatCount
            );
        }

        var infoLabel = new System.Windows.Forms.Label
        {
            Text =
                $"原文件 {diag.OriginalBytes / 1024.0 / 1024.0:F1} MB ｜ 命名区域 {diag.NamedRangeCount} 个（不自动清理，需手动核查）\n"
                + "批注保留不清理；瘦身只收缩超出真实数据的冗余行列 + 清空条件格式，原地覆写（git 可回溯，不再另存副本）。",
            AutoSize = false,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 10),
            Padding = new Padding(2),
            ForeColor = System.Drawing.SystemColors.GrayText,
        };

        var slimButton = new System.Windows.Forms.Button
        {
            Text = "执行瘦身（原地保存）",
            Dock = DockStyle.Fill,
            Height = 34,
            Margin = new Padding(0),
            FlatStyle = FlatStyle.System,
            Enabled = hasWaste,
        };
        slimButton.Click += (_, _) =>
        {
            form.Close();
            Slim(wb, diag);
        };

        layout.Controls.Add(grid, 0, 0);
        layout.Controls.Add(infoLabel, 0, 1);
        layout.Controls.Add(slimButton, 0, 2);

        form.Controls.Add(layout);
        form.Show();
    }
}
