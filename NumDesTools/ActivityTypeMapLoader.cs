using OfficeOpenXml;
using System.Text;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 从 public\Excels\Tables\TableTools\ActivityTypeMap.xlsx 读取活动类型→关联表映射。
/// 替代 ActivityTableRules.json 中的 typeMultiSubTableRules，支持人工在 xlsx 中追加行。
///
/// xlsx 结构：
///   Sheet "TypeDesc"  : #说明 | type | 枚举名 | 中文说明 | LogicBase类  （行3起，行1标题,行2表头）
///   Sheet "TypeTables": #说明 | type | 枚举名 | excelFile | lookupField | 来源 | 备注
///       excelFile 格式：普通表 "ActivityXxx.xlsx"，多 Sheet 表 "$活动xxx.xlsx#SheetName"
/// </summary>
public static class ActivityTypeMapLoader
{
    // ── 数据模型 ──────────────────────────────────────────────────────────────

    public record TypeDescEntry(int Type, string EnumName, string Desc, string LogicBase);

    public record TableEntry(
        int    Type,
        string EnumName,
        string ExcelFile,   // 含 #SheetName 后缀时为多 Sheet 表
        string LookupField, // 关联字段，默认 activityID
        string Source,      // 来源标注（typeTableMap / multiRules / 人工）
        string Remark       // 备注
    );

    // ── xlsx 路径（相对活动数据目录解析） ─────────────────────────────────────

    /// <summary>
    /// 根据活动数据目录（如 C:\M1Work\Public\Excels\Tables）推算 xlsx 路径。
    /// 返回 null 表示文件不存在。
    /// </summary>
    public static string? ResolveXlsxPath(string activityExcelDir)
    {
        // activityExcelDir 通常是 Public\Excels\Tables，xlsx 挪到了同级 TablesTools 子目录
        var path = Path.Combine(activityExcelDir, "TablesTools", "#ActivityTypeMap.xlsx");
        if (File.Exists(path)) return path;

        // 向上找两级，兼容从子目录传入的情况
        var parent = Path.GetDirectoryName(activityExcelDir);
        if (parent != null)
        {
            path = Path.Combine(parent, "Excels", "TablesTools", "#ActivityTypeMap.xlsx");
            if (File.Exists(path)) return path;
        }
        return null;
    }

    // ── 加载 ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// 读取 xlsx，返回所有 TypeTables 行按 type 分组的字典。
    /// 找不到 xlsx 返回 null（调用方可回退到 JSON）。
    /// </summary>
    public static Dictionary<int, List<TableEntry>>? LoadTypeTables(string xlsxPath)
    {
        if (!File.Exists(xlsxPath)) return null;
        try
        {
            using var pkg = new ExcelPackage(new FileInfo(xlsxPath));
            var sheet = pkg.Workbook.Worksheets["TypeTables"];
            if (sheet?.Dimension == null) return null;

            var result = new Dictionary<int, List<TableEntry>>();
            var end    = sheet.Dimension.End.Row;

            for (int row = 3; row <= end; row++)
            {
                var typeStr = sheet.Cells[row, 2].Value?.ToString();
                if (!int.TryParse(typeStr, out var type)) continue;

                var excelFile   = sheet.Cells[row, 4].Value?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(excelFile)) continue;

                var lookupField = sheet.Cells[row, 5].Value?.ToString() ?? "activityID";
                if (string.IsNullOrWhiteSpace(lookupField)) lookupField = "activityID";

                var entry = new TableEntry(
                    type,
                    sheet.Cells[row, 3].Value?.ToString() ?? string.Empty,
                    excelFile,
                    lookupField,
                    sheet.Cells[row, 6].Value?.ToString() ?? string.Empty,
                    sheet.Cells[row, 7].Value?.ToString() ?? string.Empty
                );

                if (!result.TryGetValue(type, out var list))
                    result[type] = list = [];
                list.Add(entry);
            }
            return result;
        }
        catch { return null; }
    }

    /// <summary>
    /// 读取 TypeDesc sheet，返回 type → TypeDescEntry 字典。
    /// </summary>
    public static Dictionary<int, TypeDescEntry> LoadTypeDesc(string xlsxPath)
    {
        var result = new Dictionary<int, TypeDescEntry>();
        if (!File.Exists(xlsxPath)) return result;
        try
        {
            using var pkg = new ExcelPackage(new FileInfo(xlsxPath));
            var sheet = pkg.Workbook.Worksheets["TypeDesc"];
            if (sheet?.Dimension == null) return result;

            for (int row = 3; row <= sheet.Dimension.End.Row; row++)
            {
                var typeStr = sheet.Cells[row, 2].Value?.ToString();
                if (!int.TryParse(typeStr, out var type)) continue;

                result[type] = new TypeDescEntry(
                    type,
                    sheet.Cells[row, 3].Value?.ToString() ?? string.Empty,
                    sheet.Cells[row, 4].Value?.ToString() ?? string.Empty,
                    sheet.Cells[row, 5].Value?.ToString() ?? string.Empty
                );
            }
        }
        catch { /* 读取失败返回空字典 */ }
        return result;
    }

    // ── 全目录扫描：找未配置的关联表 ─────────────────────────────────────────

    /// <summary>
    /// 扫描 <paramref name="excelDir"/> 下所有 .xlsx/.xlsm，
    /// 找出可能与 <paramref name="activityDataId"/> 相关但未在 <paramref name="configuredFiles"/> 中的表。
    ///
    /// 返回：每张疑似表包含文件名、匹配列名、若干样本行（id+备注）。
    /// </summary>
    public static List<MissingTableHint> ScanForMissingTables(
        string excelDir,
        string activityDataId,
        HashSet<string> configuredFiles,
        StringBuilder? report = null)
    {
        var hints = new List<MissingTableHint>();
        if (!Directory.Exists(excelDir)) return hints;

        var allFiles = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Concat(Directory.GetFiles(excelDir, "*.xlsm", SearchOption.TopDirectoryOnly))
            .OrderBy(p => p)
            .ToList();

        int scanned = 0, skipped = 0;

        foreach (var filePath in allFiles)
        {
            var fileName = Path.GetFileName(filePath);

            // 跳过已配置的表、模板文件（#开头）、临时文件
            if (configuredFiles.Contains(fileName)) continue;
            if (fileName.StartsWith('#') || fileName.StartsWith('~')) continue;

            try
            {
                using var pkg  = new ExcelPackage(new FileInfo(filePath));
                var sheet = pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0];
                if (sheet?.Dimension == null) { skipped++; continue; }

                var hint = ProbeSheet(sheet, fileName, activityDataId);
                if (hint != null)
                {
                    hints.Add(hint);
                    scanned++;
                }
                else skipped++;
            }
            catch { skipped++; }
        }

        report?.AppendLine($"全目录扫描：发现 {scanned} 张未配置的疑似关联表，跳过 {skipped} 张");
        return hints;
    }

    // ── 内部：探测一张 sheet 是否含 activityId 关联数据 ─────────────────────

    private static readonly HashSet<string> _skipProbeFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "sub_table_id", "subtableid", "sub_id"
    };

    private static MissingTableHint? ProbeSheet(
        ExcelWorksheet sheet, string fileName, string activityDataId)
    {
        if (sheet.Dimension == null) return null;

        var endRow = sheet.Dimension.End.Row;
        var endCol = sheet.Dimension.End.Column;
        if (endRow < 3) return null;

        // 找备注列（#备注 或 备注）
        int remarkCol = -1;
        for (int c = 1; c <= endCol; c++)
        {
            var h = sheet.Cells[2, c].Value?.ToString() ?? string.Empty;
            if (h == "#备注" || h == "备注") { remarkCol = c; break; }
        }

        // 从第 2 列开始扫，找前缀匹配
        for (int col = 2; col <= endCol; col++)
        {
            var header = sheet.Cells[2, col].Value?.ToString() ?? string.Empty;
            if (_skipProbeFields.Contains(header)) continue;

            var samples = new List<SampleRow>();
            for (int row = 3; row <= endRow && row <= 3 + 200; row++) // 最多扫 200 行
            {
                var cellStr = sheet.Cells[row, col].Value?.ToString();
                if (cellStr == null) continue;
                if (cellStr != activityDataId && !cellStr.StartsWith(activityDataId)) continue;

                // 取 id 列（第2列）和备注列
                var id     = sheet.Cells[row, 2].Value?.ToString() ?? string.Empty;
                var remark = remarkCol > 0
                    ? sheet.Cells[row, remarkCol].Value?.ToString() ?? string.Empty
                    : string.Empty;
                samples.Add(new SampleRow(id, remark));
                if (samples.Count >= 5) break;
            }

            if (samples.Count > 0)
                return new MissingTableHint(fileName, header, samples);
        }
        return null;
    }

    // ── 结果数据结构 ─────────────────────────────────────────────────────────

    /// <summary>扫描发现的疑似未配置关联表</summary>
    public record MissingTableHint(
        string FileName,        // 文件名（含扩展名）
        string MatchedField,    // 疑似关联字段名
        List<SampleRow> Samples // 前几行样本（id + 备注）
    );

    public record SampleRow(string Id, string Remark);
}
