using OfficeOpenXml;

namespace NumDesTools.Scanner;

/// <summary>
/// 读取配置 Excel，约定表头格式：
///   Row1 = 表名注释 (#开头)
///   Row2 = 字段名   (#开头为注释列，跳过)
///   Row3 = 类型声明
///   Row4 = 中文说明
///   Row5+ = 数据行
/// 跳过 A 列（#注释列）和字段名以 # 开头的列。
/// </summary>
public static class ExcelReader
{
    public const int HeaderRow = 2;
    public const int DataStartRow = 5;

    /// <summary>
    /// 读取工作表，返回 (字段名列表, 数据行列表)。
    /// 每行数据是 fieldName→cellText 的字典。
    /// </summary>
    public static (List<string> Fields, List<(int Row, Dictionary<string, string> Data)> Rows)
        Read(string excelPath, string? sheetName = null)
    {
        using var pkg = new ExcelPackage(new FileInfo(excelPath));
        var ws = sheetName != null
            ? pkg.Workbook.Worksheets[sheetName] ?? pkg.Workbook.Worksheets[0]
            : pkg.Workbook.Worksheets[0];

        if (ws == null) return ([], []);

        int colCount = ws.Dimension?.Columns ?? 0;
        int rowCount = ws.Dimension?.Rows    ?? 0;
        if (colCount == 0 || rowCount < DataStartRow) return ([], []);

        // 收集字段名（Row2），跳过 # 开头
        var fields     = new List<string>();
        var colIndexes = new List<int>(); // 有效列索引（1-based）
        for (int c = 1; c <= colCount; c++)
        {
            var name = ws.Cells[HeaderRow, c].Text?.Trim() ?? "";
            if (string.IsNullOrEmpty(name) || name.StartsWith('#')) continue;
            fields.Add(name);
            colIndexes.Add(c);
        }

        // 读数据行
        var rows = new List<(int Row, Dictionary<string, string> Data)>();
        for (int r = DataStartRow; r <= rowCount; r++)
        {
            var dict = new Dictionary<string, string>(fields.Count);
            bool anyValue = false;
            for (int i = 0; i < fields.Count; i++)
            {
                var text = ws.Cells[r, colIndexes[i]].Text?.Trim() ?? "";
                dict[fields[i]] = text;
                if (!string.IsNullOrEmpty(text)) anyValue = true;
            }
            if (!anyValue) continue; // 跳过全空行
            rows.Add((r, dict));
        }

        return (fields, rows);
    }

    /// <summary>
    /// 仅读取某列所有非空值，返回值集合（用于 L2 构建 ID 索引）。
    /// </summary>
    public static HashSet<string> ReadKeySet(string excelPath, string keyField, string? sheetName = null)
    {
        var (_, rows) = Read(excelPath, sheetName);
        var set = new HashSet<string>();
        foreach (var (_, data) in rows)
            if (data.TryGetValue(keyField, out var v) && !string.IsNullOrEmpty(v))
                set.Add(v);
        return set;
    }
}
