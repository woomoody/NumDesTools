using OfficeOpenXml;

namespace NumDesTools;

/// <summary>
/// 从 xlsx worksheet 解析单元格历史查询所需的行列数据。
/// 无 git 依赖，纯内存操作，供 CellGitHistoryService 和单元测试共用。
/// </summary>
public static class CellHistoryXlsxReader
{
    /// <summary>
    /// 找 key 列：row 2 中第一个不以 # 开头的列（1-based）。
    /// </summary>
    public static int FindKeyColIdx(ExcelWorksheet ws)
    {
        if (ws.Dimension == null)
            return 1;
        for (int c = 1; c <= Math.Min(ws.Dimension.End.Column, 30); c++)
        {
            var h = ws.Cells[2, c].Value?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !h.StartsWith('#'))
                return c;
        }
        return 1;
    }

    /// <summary>
    /// 解析 worksheet 数据，返回 rowKey → colName → value 映射。
    /// 从 row 3 开始扫描（兼容标准 config 表 row5 起和 type 表 row3 起）。
    /// </summary>
    public static Dictionary<string, Dictionary<string, string>> ParseSheetData(
        ExcelWorksheet ws
    )
    {
        var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
        if (ws.Dimension == null)
            return data;

        var keyColIdx = FindKeyColIdx(ws);

        // 建列名 → 列号映射（row 2）
        var colNameToIdx = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int c = 1; c <= ws.Dimension.End.Column; c++)
        {
            var h = ws.Cells[2, c].Value?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !colNameToIdx.ContainsKey(h))
                colNameToIdx[h] = c;
        }

        // 从 row 3 开始扫描，兼容 type 表（无 type/label 行）和标准表（row5+）
        for (int r = 3; r <= ws.Dimension.End.Row; r++)
        {
            var key = ws.Cells[r, keyColIdx].Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(key))
                continue;

            var row = new Dictionary<string, string>(
                colNameToIdx.Count,
                StringComparer.Ordinal
            );
            foreach (var (col, idx) in colNameToIdx)
                row[col] = ws.Cells[r, idx].Value?.ToString() ?? "";

            data[key] = row;
        }
        return data;
    }
}
