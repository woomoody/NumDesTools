using OfficeOpenXml;

namespace NumDesTools;

// 瘦身：全量扫描指定根目录下的 xlsx，EPPlus 离线处理，无文件锁竞态。两类瘦身：
// 1) 把被误存成 sharedStrings 的纯数字改回原生数字类型（判定见 CellValueNormalizer）；
// 2) 清掉空白行列尾部残留格式（老版本"格式瘦身诊断"功能的诊断逻辑，复用同一次遍历顺便算）。
// 只处理数据行（第5行起），不动行1-4的字段结构元数据。
// "当前工作簿"(COM)模式已删除：Excel 自己的 Save() 不会像 EPPlus 全量重写那样重新压缩 zip，
// 体积瘦不下去，COM 逐格/整块操作还多一堆性能和锁竞态坑，收益不值得维护这条路径。
internal static class XlsxSlimmer
{
    private const int DataStartRow = 5;

    internal record FileSlimResult(
        string FilePath,
        long SizeBefore,
        long SizeAfter,
        int Converted,
        int TrimmedRows,
        int TrimmedCols,
        string? Error
    );

    internal static List<string> FindSlimmableFiles(string rootDir, double minSizeMb = 0) =>
        Directory
            .EnumerateFiles(rootDir, "*.xlsx", SearchOption.AllDirectories)
            .Where(f =>
            {
                var name = Path.GetFileName(f);
                if (name.Contains('#') || name.Contains('~'))
                    return false;
                // 小文件本来就没多少 sharedStrings 冗余，转换收益微乎其微，还多一次踩坑
                // 机会（EPPlus 存盘偏偏对某些文件有独立于本工具的内部 bug），不值得处理。
                return minSizeMb <= 0 || new FileInfo(f).Length >= minSizeMb * 1024 * 1024;
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
            var trimmedRows = 0;
            var trimmedCols = 0;
            foreach (var sheet in pkg.Workbook.Worksheets)
            {
                if (sheet.Name.StartsWith('#') || sheet.Dimension is null)
                    continue;

                // 真实数据边界：跟数字归一化复用同一次 sheet.Cells 遍历顺便算，不用再多扫一遍。
                // ws.Cells 只枚举实际存在的稀疏 cell，比双重 for + 索引器快。
                int trueMaxRow = 0,
                    trueMaxCol = 0;
                foreach (var cell in sheet.Cells)
                {
                    if (cell.Value is not null)
                    {
                        if (cell.Start.Row > trueMaxRow)
                            trueMaxRow = cell.Start.Row;
                        if (cell.Start.Column > trueMaxCol)
                            trueMaxCol = cell.Start.Column;
                    }

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

                var (rowsTrimmedHere, colsTrimmedHere) = TrimTrailingBlank(
                    sheet,
                    trueMaxRow,
                    trueMaxCol,
                    apply: !preview
                );
                trimmedRows += rowsTrimmedHere;
                trimmedCols += colsTrimmedHere;
            }

            if (preview || (converted == 0 && trimmedRows == 0 && trimmedCols == 0))
                return new FileSlimResult(
                    path,
                    sizeBefore,
                    sizeBefore,
                    converted,
                    trimmedRows,
                    trimmedCols,
                    null
                );

            pkg.Save();
            return new FileSlimResult(
                path,
                sizeBefore,
                new FileInfo(path).Length,
                converted,
                trimmedRows,
                trimmedCols,
                null
            );
        }
        catch (Exception ex)
        {
            return new FileSlimResult(path, sizeBefore, sizeBefore, 0, 0, 0, ex.Message);
        }
    }

    // 空白行列尾部残留格式：常见于误 Ctrl+A 设置过格式，本身不带数据但仍占 cell entry，
    // Dimension 会把这些也算进去，看起来像是表里"莫名多出一大截空行"。
    // apply=false 只统计不改动（预览用）；apply=true 真的执行 DeleteRow/DeleteColumn。
    internal static (int TrimmedRows, int TrimmedCols) TrimTrailingBlank(
        ExcelWorksheet sheet,
        int trueMaxRow,
        int trueMaxCol,
        bool apply
    )
    {
        if (sheet.Dimension is null)
            return (0, 0);
        var dimEndRow = sheet.Dimension.End.Row;
        var dimEndCol = sheet.Dimension.End.Column;
        var trimmedRows = trueMaxRow > 0 && dimEndRow > trueMaxRow ? dimEndRow - trueMaxRow : 0;
        var trimmedCols = trueMaxCol > 0 && dimEndCol > trueMaxCol ? dimEndCol - trueMaxCol : 0;
        if (apply)
        {
            if (trimmedRows > 0)
                sheet.DeleteRow(trueMaxRow + 1, trimmedRows);
            if (trimmedCols > 0)
                sheet.DeleteColumn(trueMaxCol + 1, trimmedCols);
        }
        return (trimmedRows, trimmedCols);
    }

    // ActivityDataBackupTool 删除活动数据区间后顺手调这个清一次尾部残留格式：没有别的既有扫描可以
    // 顺路捎带真实边界，所以自己单独扫一遍（比起省掉整个文件重新 load/save 一轮，这一次扫描便宜得多）。
    internal static (int TrimmedRows, int TrimmedCols) TrimTrailingBlank(ExcelWorksheet sheet)
    {
        if (sheet.Dimension is null)
            return (0, 0);
        int trueMaxRow = 0,
            trueMaxCol = 0;
        foreach (var cell in sheet.Cells)
        {
            if (cell.Value is null)
                continue;
            if (cell.Start.Row > trueMaxRow)
                trueMaxRow = cell.Start.Row;
            if (cell.Start.Column > trueMaxCol)
                trueMaxCol = cell.Start.Column;
        }
        return TrimTrailingBlank(sheet, trueMaxRow, trueMaxCol, apply: true);
    }
}
