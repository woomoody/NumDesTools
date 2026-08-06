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
    /// 优化：批量读取整个 range 为 2D 数组，避免逐格创建 ExcelCell 对象。
    /// 注意：必须从 row 1 开始读 range，因为 values 索引是 0-based 相对 range 偏移，
    /// 如果 startRow != 1 则 values[1,c] 不是 row 2 导致列名映射全错。
    /// </summary>
    public static Dictionary<string, Dictionary<string, string>> ParseSheetData(
        ExcelWorksheet ws
    )
    {
        var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
        if (ws.Dimension == null)
            return data;

        var dim = ws.Dimension;
        int endRow = dim.End.Row;
        int endCol = dim.End.Column;

        // 始终从 row 1, col 1 开始读，保证 values[0,0] = ws.Cells[1,1]
        // 这样 values[1, c] 一定是 row 2（列名行），values[2, *] 是 row 3（数据开始）
        int rangeStartRow = 1;
        int rangeStartCol = 1;
        var range = ws.Cells[rangeStartRow, rangeStartCol, endRow, endCol];
        var values = range.Value as object[,];
        if (values == null)
            return data;

        int rows = values.GetLength(0); // 0-based: values[0..rows-1, *]
        int cols = values.GetLength(1); // 0-based: values[*, 0..cols-1]

        // 0-based col → 实际 Excel 列号（1-based）
        int ColAt(int relCol) => relCol + rangeStartCol;

        // 列名行在 row 2（values[1, *]），找 key 列
        int keyColIdx = -1;
        for (int c = 0; c < cols && c < 30; c++)
        {
            var h = values[1, c]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !h.StartsWith('#'))
            {
                keyColIdx = ColAt(c);
                break;
            }
        }
        if (keyColIdx < 0)
            keyColIdx = rangeStartCol;

        // 建列名 → 实际列号映射
        var colNameToCol = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int c = 0; c < cols; c++)
        {
            var h = values[1, c]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !colNameToCol.ContainsKey(h))
                colNameToCol[h] = ColAt(c);
        }

        // 从 row 3（values[2, *]）开始扫描
        int keyRelCol = keyColIdx - rangeStartCol;
        for (int r = 2; r < rows; r++)
        {
            var key = keyRelCol >= 0 && keyRelCol < cols ? values[r, keyRelCol]?.ToString() ?? "" : "";
            if (string.IsNullOrEmpty(key))
                continue;

            var row = new Dictionary<string, string>(colNameToCol.Count, StringComparer.Ordinal);
            foreach (var (colName, actualCol) in colNameToCol)
            {
                int relCol = actualCol - rangeStartCol;
                row[colName] = relCol >= 0 && relCol < cols ? values[r, relCol]?.ToString() ?? "" : "";
            }

            data[key] = row;
        }
        return data;
    }
}
