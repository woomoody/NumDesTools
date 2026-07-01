using OfficeOpenXml;

namespace NumDesTools;

internal static class StaleActivityScanner
{
    internal static void Scan()
    {
        var basePath = AppServices.Config.Paths.BasePath;
        var expiredIds = LoadExpiredLteIds(basePath);

        if (expiredIds.Count == 0)
        {
            MessageBox.Show("未找到过期 LTE 活动数据，请检查 BasePath 配置。", "扫描陈旧活动数据");
            return;
        }

        if (AppServices.App.ActiveSheet is not Worksheet ws)
        {
            MessageBox.Show("没有活动的工作表。", "扫描陈旧活动数据");
            return;
        }

        var results = ScanSheet(ws, expiredIds);
        ShowResults(results, ws.Name, expiredIds.Count);
    }

    // Returns expired LTE activityIDs (closeTime < now-1yr) from the config files
    private static HashSet<string> LoadExpiredLteIds(string basePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var clientPath = Path.Combine(basePath, "ActivityClientData.xlsx");
        if (!File.Exists(clientPath))
            return new HashSet<string>();

        // G=col7=type, H=col8=activityID; data from row 5
        var lteIds = new HashSet<string>();
        using (var pkg = new ExcelPackage(new FileInfo(clientPath)))
        {
            var sh = pkg.Workbook.Worksheets[0];
            if (sh?.Dimension is null)
                return lteIds;
            for (int r = 5; r <= sh.Dimension.End.Row; r++)
            {
                if (sh.Cells[r, 7].Text.Trim() != "2")
                    continue;
                var id = sh.Cells[r, 8].Text.Trim();
                if (id.Length > 0)
                    lteIds.Add(id);
            }
        }

        if (lteIds.Count == 0)
            return lteIds;

        // B=col2=id, O=col15=closeTime (Unix sec); data from row 5
        var serverFiles = Directory.GetFiles(basePath, "*ActivityServerData.xlsm");
        if (serverFiles.Length == 0)
            return new HashSet<string>();

        var serverPath =
            Array.Find(serverFiles, f => !Path.GetFileName(f).Contains("旧")) ?? serverFiles[0];

        var cutoff = DateTimeOffset.UtcNow.AddYears(-1).ToUnixTimeSeconds();
        var expired = new HashSet<string>();
        using (var pkg = new ExcelPackage(new FileInfo(serverPath)))
        {
            var sh = pkg.Workbook.Worksheets[0];
            if (sh?.Dimension is null)
                return expired;
            for (int r = 5; r <= sh.Dimension.End.Row; r++)
            {
                var id = sh.Cells[r, 2].Text.Trim();
                if (!lteIds.Contains(id))
                    continue;
                if (long.TryParse(sh.Cells[r, 15].Text.Trim(), out long t) && t > 0 && t < cutoff)
                    expired.Add(id);
            }
        }
        return expired;
    }

    private static List<(int Row, string ColName, string Value)> ScanSheet(
        Worksheet ws,
        HashSet<string> expiredIds
    )
    {
        var used = ws.UsedRange;
        if (used is null)
            return new List<(int, string, string)>();

        // Read all values in one COM call for performance
        var values = used.Value2 as object[,];
        if (values is null)
            return new List<(int, string, string)>();

        int r1 = used.Row;
        int c1 = used.Column;
        int rowCount = values.GetLength(0);
        int colCount = values.GetLength(1);

        // Row 2 (sheet-absolute) → array index
        int hdrIdx = 2 - r1 + 1;
        if (hdrIdx < 1 || hdrIdx > rowCount)
            return new List<(int, string, string)>();

        // Columns whose row-2 name looks like an activity ID column
        var actCols = new List<(int ArrCol, string Name)>();
        for (int c = 1; c <= colCount; c++)
        {
            var name = values[hdrIdx, c]?.ToString()?.TrimStart('#') ?? "";
            if (IsActivityIdCol(name))
                actCols.Add((c, name));
        }
        // Fallback: scan every column
        if (actCols.Count == 0)
            for (int c = 1; c <= colCount; c++)
                actCols.Add((c, values[hdrIdx, c]?.ToString() ?? $"Col{c1 + c - 1}"));

        int dataStartIdx = Math.Max(5 - r1 + 1, 1);
        var results = new List<(int Row, string ColName, string Value)>();
        for (int ri = dataStartIdx; ri <= rowCount; ri++)
        {
            foreach (var (ac, colName) in actCols)
            {
                var val = values[ri, ac]?.ToString()?.Trim() ?? "";
                if (!expiredIds.Contains(val))
                    continue;
                results.Add((r1 + ri - 1, colName, val));
                break;
            }
        }
        return results;
    }

    private static bool IsActivityIdCol(string name)
    {
        if (string.IsNullOrEmpty(name))
            return false;
        var lc = name.ToLowerInvariant();
        return lc.Contains("activityid") || lc == "_sub_table_id" || lc == "activity_id";
    }

    private static void ShowResults(
        List<(int Row, string ColName, string Value)> results,
        string sheetName,
        int expiredCount
    )
    {
        var form = new Form
        {
            Text =
                $"陈旧活动数据 — {sheetName}（命中 {results.Count} 行 / {expiredCount} 个过期ID）",
            Width = 580,
            Height = 460,
            StartPosition = FormStartPosition.CenterScreen,
            KeyPreview = true,
        };
        form.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                form.Close();
        };

        if (results.Count == 0)
        {
            form.Controls.Add(
                new System.Windows.Forms.Label
                {
                    Text = "✓ 未发现陈旧数据",
                    Dock = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Font = new System.Drawing.Font("Segoe UI", 14),
                }
            );
        }
        else
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = System.Drawing.SystemColors.Window,
            };
            grid.Columns.Add("Row", "行号");
            grid.Columns.Add("Col", "列名");
            grid.Columns.Add("Value", "活动ID");
            grid.Columns[0].FillWeight = 15;
            grid.Columns[1].FillWeight = 35;
            grid.Columns[2].FillWeight = 50;

            foreach (var (row, col, val) in results)
                grid.Rows.Add(row, col, val);

            form.Controls.Add(grid);
        }

        form.Show();
    }
}
