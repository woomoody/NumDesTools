using Sylvan.Data.Excel;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 用 Sylvan.Data.Excel 流式读取一个 xlsx 的指定 sheet，一次性构建填好数据的 <see cref="ColumnStore"/>。
/// 只读、单趟前向扫描：不重开 zip、不再逐行 Dispatcher.Invoke（对比旧 <c>OoxmlLazyReader</c> 每次重开
/// zip + 重解析 sharedStrings 的开销）。列名沿用 Excel 原生列名 A/B/.../CF（与 MainWindow 一致，1-based）。
/// </summary>
public static class ColumnStoreExcelLoader
{
    /// <summary>
    /// 加载 <paramref name="xlsxPath"/> 的首个工作表（或 <paramref name="sheetName"/> 指定的表），
    /// 全部单元格按文本存入 ColumnStore。空单元格存为 <c>null</c>。
    /// 逐行列数不齐（Sylvan 的 jagged 行）时按需扩展列，最终 <see cref="ColumnStore.ColumnCount"/>
    /// 为所有行出现过的最大列数。
    /// </summary>
    public static ColumnStore Load(string xlsxPath, string? sheetName = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(xlsxPath);

        var options = new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders };
        using var reader = ExcelDataReader.Create(xlsxPath, options);

        if (sheetName is not null)
        {
            while (
                !string.Equals(reader.WorksheetName, sheetName, StringComparison.Ordinal)
                && reader.NextResult()
            )
            {
                // 前进到目标工作表
            }

            if (!string.Equals(reader.WorksheetName, sheetName, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    $"Worksheet '{sheetName}' not found in {xlsxPath}",
                    nameof(sheetName)
                );
            }
        }

        return ReadSheet(reader);
    }

    private static ColumnStore ReadSheet(ExcelDataReader reader)
    {
        var initialColumns = Math.Max(reader.FieldCount, 1);
        var names = new string[initialColumns];
        for (var col = 0; col < initialColumns; col++)
        {
            names[col] = ExcelColumnName(col + 1);
        }

        var store = ColumnStore.Create(names);

        while (reader.Read())
        {
            var fieldCount = reader.RowFieldCount;
            if (fieldCount > store.ColumnCount)
            {
                store.EnsureColumnCount(fieldCount, col => ExcelColumnName(col + 1));
            }

            var row = store.AppendRow();
            for (var col = 0; col < fieldCount; col++)
            {
                if (!reader.IsDBNull(col))
                {
                    store.SetCellQuiet(row, col, reader.GetString(col));
                }
            }
        }

        // 刚加载完是"干净"基线：清脏 + 重置 StructureChanged（加载期 EnsureColumnCount 处理 jagged
        // 行会置 StructureChanged，但那不是用户编辑，保存路径不应据此误判为"整表重写"）。
        store.ClearDirty();
        return store;
    }

    /// <summary>1-based 列序号转 Excel 列名（1→A、26→Z、27→AA）。</summary>
    private static string ExcelColumnName(int col)
    {
        var name = string.Empty;
        while (col > 0)
        {
            var remainder = (col - 1) % 26;
            name = (char)('A' + remainder) + name;
            col = (col - 1) / 26;
        }

        return name;
    }
}
