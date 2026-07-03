using OfficeOpenXml;

namespace NumDesTools;

// 瘦身：把 xlsx 里被误存成 sharedStrings 的纯数字改回原生数字类型（判定见 CellValueNormalizer），
// 省体积、去重零收益的冗余。只处理数据行（第5行起），不动行1-4的字段结构元数据。
// 两种模式机制不同：当前工作簿走 COM（整块 Value2 读写，避免逐格调用），全量扫描走 EPPlus（离线文件，无锁竞态）。
internal static class XlsxSlimmer
{
    private const int DataStartRow = 5;

    internal record SheetSlimResult(string SheetName, int Scanned, int Converted);

    internal record FileSlimResult(
        string FilePath,
        long SizeBefore,
        long SizeAfter,
        int Converted,
        string? Error
    );

    // ── 当前工作簿：纯 COM，全程不 Close/不用 EPPlus 重开，避免文件锁竞态 ──────

    internal static List<SheetSlimResult> SlimCurrentWorkbook(Workbook wb, bool preview)
    {
        var results = new List<SheetSlimResult>();
        foreach (Worksheet sheet in wb.Worksheets)
        {
            if (sheet.Name.StartsWith('#'))
                continue;
            var used = sheet.UsedRange;
            var lastRow = used.Row + used.Rows.Count - 1;
            var lastCol = used.Column + used.Columns.Count - 1;
            var dataStart = Math.Max(DataStartRow, used.Row);
            if (dataStart > lastRow)
                continue;

            var range = (Range)
                sheet.Range[sheet.Cells[dataStart, used.Column], sheet.Cells[lastRow, lastCol]];
            var data = (object[,])range.Value2; // 一次性整块读，避免逐格 COM 往返
            int rows = data.GetLength(0),
                cols = data.GetLength(1);

            int scanned = 0,
                converted = 0;
            var colFormats = new Dictionary<int, HashSet<string>>(); // 列 -> 该列出现过的格式集合

            for (int r = 1; r <= rows; r++)
            for (int c = 1; c <= cols; c++)
            {
                if (data[r, c] is not string s || s.Length == 0)
                    continue;
                scanned++;
                var normalized = CellValueNormalizer.Normalize(s);
                if (normalized is null)
                    continue;
                data[r, c] = normalized;
                converted++;
                var fmt = normalized is long ? "0" : "0.##############";
                if (!colFormats.TryGetValue(c, out var fmts))
                    colFormats[c] = fmts = [];
                fmts.Add(fmt);
            }

            if (scanned > 0)
                results.Add(new SheetSlimResult(sheet.Name, scanned, converted));

            if (preview || converted == 0)
                continue;

            range.Value2 = data; // 一次性写回

            // 按列批量锁格式，一列一次 COM 调用：实测踩过逐格调用的坑——56 万格逐格设置
            // NumberFormat 在 Item.xlsx(12.8 万行)上直接卡死看起来像死机。前提：同一列的
            // 转换结果类型一致（实测三张大表全是 long，没混过 double），这个前提当前成立；
            // 真出现混合类型的列，退化成对该列逐格设置（数量通常很小，不会再卡）。
            foreach (var (col, fmts) in colFormats)
            {
                var colRange = (Range)
                    sheet.Range[
                        sheet.Cells[dataStart, used.Column + col - 1],
                        sheet.Cells[lastRow, used.Column + col - 1]
                    ];
                if (fmts.Count == 1)
                {
                    colRange.NumberFormat = fmts.First();
                    continue;
                }
                for (int r = 1; r <= rows; r++)
                {
                    if (data[r, col] is not (long or double))
                        continue;
                    ((Range)sheet.Cells[dataStart + r - 1, used.Column + col - 1]).NumberFormat =
                        data[r, col] is long ? "0" : "0.##############";
                }
            }
        }

        if (!preview)
            wb.Save();
        return results;
    }

    // ── 全量扫描：EPPlus 逐文件处理，离线文件无锁竞态 ─────────────────────────

    internal static List<string> FindSlimmableFiles(string rootDir) =>
        Directory
            .EnumerateFiles(rootDir, "*.xlsx", SearchOption.AllDirectories)
            .Where(f =>
            {
                var name = Path.GetFileName(f);
                return !name.Contains('#') && !name.Contains('~');
            })
            .ToList();

    internal static FileSlimResult SlimFile(string path, bool preview)
    {
        var sizeBefore = new FileInfo(path).Length;
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
            using var pkg = new ExcelPackage(new FileInfo(path));
            // P1：EPPlus 8.x 的压缩级别不在 Save() 参数上，是 ExcelPackage.Compression 属性
            // （反射确认过，文档示例没提到）。锦上添花，风险低，跟 P0 数字归一化一起顺手做。
            pkg.Compression = CompressionLevel.BestCompression;
            var converted = 0;
            foreach (var sheet in pkg.Workbook.Worksheets)
            {
                if (sheet.Name.StartsWith('#') || sheet.Dimension is null)
                    continue;
                // ws.Cells 只枚举实际存在的稀疏 cell，比双重 for + 索引器快
                foreach (var cell in sheet.Cells)
                {
                    if (cell.Start.Row < DataStartRow)
                        continue;
                    if (cell.Value is not string s || s.Length == 0)
                        continue;
                    var normalized = CellValueNormalizer.Normalize(s);
                    if (normalized is null)
                        continue;
                    converted++;
                    if (preview)
                        continue;
                    cell.Value = normalized;
                    cell.Style.Numberformat.Format = normalized is long ? "0" : "0.##############";
                }
            }

            if (preview || converted == 0)
                return new FileSlimResult(path, sizeBefore, sizeBefore, converted, null);

            pkg.Save();
            return new FileSlimResult(path, sizeBefore, new FileInfo(path).Length, converted, null);
        }
        catch (Exception ex)
        {
            return new FileSlimResult(path, sizeBefore, sizeBefore, 0, ex.Message);
        }
    }
}
