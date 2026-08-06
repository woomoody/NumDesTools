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
    /// </summary>
    public static Dictionary<string, Dictionary<string, string>> ParseSheetData(
        ExcelWorksheet ws
    )
    {
        var data = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
        if (ws.Dimension == null)
            return data;

        var dim = ws.Dimension;
        int startRow = dim.Start.Row;
        int endRow = dim.End.Row;
        int startCol = dim.Start.Column;
        int endCol = dim.End.Column;

        // 批量读取整张表为 2D object 数组（0-based: values[0,0] = ws.Cells[startRow,startCol]）
        var range = ws.Cells[startRow, startCol, endRow, endCol];
        var values = range.Value as object[,];
        if (values == null)
            return data;

        int rows = values.GetLength(0); // 0-based 行数
        int cols = values.GetLength(1); // 0-based 列数

        // 0-based 索引 → 实际列号（1-based Excel）
        int ColAt(int relCol) => relCol + startCol;

        // 找 key 列：row 2 在 values 中索引 = 1（0-based）
        int keyColIdx = -1;
        for (int c = 0; c < cols && c < 30; c++)
        {
            var h = values[1, c]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !h.StartsWith('#'))
            {
                keyColIdx = ColAt(c); // 实际列号
                break;
            }
        }
        if (keyColIdx < 0)
            keyColIdx = startCol;

        // 建列名 → 实际列号映射（0-based row=1 即 Excel row 2）
        var colNameToCol = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int c = 0; c < cols; c++)
        {
            var h = values[1, c]?.ToString() ?? "";
            if (!string.IsNullOrEmpty(h) && !colNameToCol.ContainsKey(h))
                colNameToCol[h] = ColAt(c);
        }

        // 从 row 3（0-based row=2）开始扫描
        for (int r = 2; r < rows; r++)
        {
            // keyColIdx 是实际列号，找到它在 values 中的 0-based 列索引
            int keyRelCol = keyColIdx - startCol;
            var key = keyRelCol >= 0 && keyRelCol < cols ? values[r, keyRelCol]?.ToString() ?? "" : "";
            if (string.IsNullOrEmpty(key))
                continue;

            var row = new Dictionary<string, string>(colNameToCol.Count, StringComparer.Ordinal);
            foreach (var (colName, actualCol) in colNameToCol)
            {
                int relCol = actualCol - startCol;
                row[colName] = relCol >= 0 && relCol < cols ? values[r, relCol]?.ToString() ?? "" : "";
            }

            data[key] = row;
        }
        return data;
    }
}
