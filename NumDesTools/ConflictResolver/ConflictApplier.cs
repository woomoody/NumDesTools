using LibGit2Sharp;
using OfficeOpenXml;

namespace NumDesTools.ConflictResolver;

/// <summary>
/// 将用户的选择写回目标文件，保留原文件所有格式。
/// 写回策略：以 OURS 文件为基础，按冲突选择定向修改单元格 / 追加行 / 标记删除行。
/// 不重建文件——直接用 EPPlus 在原文件上操作，样式/公式/列宽完整保留。
/// </summary>
public static class ConflictApplier
{
    public static void Apply(FileDiff diff, string outPath, bool gitAdd = true)
    {
        // 以 OURS 文件为基础复制到目标路径（若路径不同）
        if (!outPath.Equals(diff.OursPath, StringComparison.OrdinalIgnoreCase))
            File.Copy(diff.OursPath, outPath, overwrite: true);

        using var pkg = new ExcelPackage(new FileInfo(outPath));

        foreach (var sheetDiff in diff.Sheets)
        {
            if (!sheetDiff.HasConflict) continue;

            var sheet = pkg.Workbook.Worksheets[sheetDiff.SheetName];
            if (sheet?.Dimension == null) continue;

            // 建立 key → 行号 映射（表头在第2行，数据从第3行起）
            var keyColIdx = FindKeyColIndex(sheet);
            var keyToRow  = BuildKeyRowMap(sheet, keyColIdx);

            var allCols = sheetDiff.AllColumns;

            // 先处理 Modified（只改单元格值，不移动行，不影响行号）
            foreach (var rc in sheetDiff.Rows)
            {
                if (rc.DiffType == RowDiffType.Modified)
                    ApplyModifiedRow(sheet, rc, keyToRow, allCols);
            }

            // 再追加 OnlyTheirs+Theirs（不影响已有行的行号）
            foreach (var rc in sheetDiff.Rows)
            {
                if (rc.DiffType == RowDiffType.OnlyTheirs && rc.RowChoice == ConflictChoice.Theirs)
                    AppendTheirsRow(sheet, rc, allCols);
            }

            // 确保新增列的 row3(type) / row4(label) 从 THEIRS 文件元数据补全
            EnsureNewColsMeta(sheet, sheetDiff);

            // 最后处理删除行（从大行号到小行号，避免移位影响前面的行号）
            var deleteRows = sheetDiff.Rows
                .Where(rc => rc.DiffType == RowDiffType.OnlyOurs && rc.RowChoice == ConflictChoice.Theirs)
                .Select(rc => keyToRow.TryGetValue(rc.RowKey, out var r) ? r : -1)
                .Where(r => r > 0)
                .OrderByDescending(r => r)
                .ToList();
            foreach (var delRow in deleteRows)
                sheet.DeleteRow(delRow);
        }

        pkg.Save();

        if (gitAdd) GitAdd(outPath);
    }

    // ── 定向修改 Modified 行 ────────────────────────────────────────────────

    private static void ApplyModifiedRow(
        ExcelWorksheet sheet,
        RowConflict rc,
        Dictionary<string, int> keyToRow,
        List<string> allColumns)
    {
        if (!keyToRow.TryGetValue(rc.RowKey, out var rowIdx)) return;

        foreach (var cell in rc.Cells)
        {
            if (cell.Choice != ConflictChoice.Theirs) continue;

            var colIdx = FindOrCreateHeaderCol(sheet, cell.ColName, allColumns);
            sheet.Cells[rowIdx, colIdx].Value = cell.TheirsValue;
        }
    }

    // ── 追加 OnlyTheirs 行 ───────────────────────────────────────────────────

    private static void AppendTheirsRow(ExcelWorksheet sheet, RowConflict rc, List<string> allColumns)
    {
        if (rc.TheirsFullRow == null) return;

        var lastRow = sheet.Dimension.End.Row + 1;

        foreach (var (header, val) in rc.TheirsFullRow)
        {
            if (string.IsNullOrEmpty(header)) continue;
            var colIdx = FindOrCreateHeaderCol(sheet, header, allColumns);
            sheet.Cells[lastRow, colIdx].Value = val;
        }

        // 复制上一行的样式
        var srcRow   = lastRow - 1;
        var colCount = sheet.Dimension.End.Column;
        if (srcRow >= 3)
            for (int col = 1; col <= colCount; col++)
                sheet.Cells[lastRow, col].StyleID = sheet.Cells[srcRow, col].StyleID;
    }

    // 补全新增列的 row3(type) / row4(label)
    private static void EnsureNewColsMeta(ExcelWorksheet sheet, SheetDiff sheetDiff)
    {
        var end = sheet.Dimension.End.Column;
        for (int col = 1; col <= end; col++)
        {
            var colName = sheet.Cells[2, col].Value?.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(colName)) continue;

            if (sheet.Cells[3, col].Value == null && sheetDiff.TypeRow.TryGetValue(colName, out var t) && !string.IsNullOrEmpty(t))
                sheet.Cells[3, col].Value = t;
            if (sheet.Cells[4, col].Value == null && sheetDiff.LabelRow.TryGetValue(colName, out var l) && !string.IsNullOrEmpty(l))
                sheet.Cells[4, col].Value = l;
        }
    }

    // 在 row2 查找列名，找不到则按 SheetDiff.AllColumns 顺序插入到正确位置
    private static int FindOrCreateHeaderCol(ExcelWorksheet sheet, string colName,
        List<string>? allColumns = null)
    {
        var end = sheet.Dimension.End.Column;
        for (int col = 1; col <= end; col++)
            if (sheet.Cells[2, col].Value?.ToString() == colName)
                return col;

        // 找插入位置：在 allColumns 中找到此列左侧最近的已存在列
        int insertAfterCol = end; // 默认追加到末尾
        if (allColumns != null)
        {
            var idx = allColumns.IndexOf(colName);
            if (idx > 0)
            {
                // 向左找最近已存在于 sheet 的列
                for (int i = idx - 1; i >= 0; i--)
                {
                    var anchor = allColumns[i];
                    for (int c = 1; c <= end; c++)
                    {
                        if (sheet.Cells[2, c].Value?.ToString() == anchor)
                        {
                            insertAfterCol = c;
                            goto found;
                        }
                    }
                }
                found:;
            }
        }

        // 在 insertAfterCol 之后插入一列
        var newCol = insertAfterCol + 1;
        if (newCol <= sheet.Dimension.End.Column)
            sheet.InsertColumn(newCol, 1);

        sheet.Cells[2, newCol].Value = colName;

        // 复制左侧相邻列的列宽和所有行样式，使新列格式与其他列一致
        var srcCol = insertAfterCol;
        if (srcCol >= 1)
        {
            sheet.Column(newCol).Width = sheet.Column(srcCol).Width;
            var rowEnd = sheet.Dimension.End.Row;
            for (int r = 1; r <= rowEnd; r++)
                sheet.Cells[r, newCol].StyleID = sheet.Cells[r, srcCol].StyleID;
        }

        return newCol;
    }

    // ── 辅助：找 key 列（第2列，index=1）────────────────────────────────────

    private static int FindKeyColIndex(ExcelWorksheet sheet)
    {
        // 与 Differ 对齐：key 列为表头行第2列
        if (sheet.Dimension.End.Column >= 2) return 2;
        return 1;
    }

    private static Dictionary<string, int> BuildKeyRowMap(ExcelWorksheet sheet, int keyCol)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
        {
            var key = sheet.Cells[row, keyCol].Value?.ToString();
            if (!string.IsNullOrEmpty(key) && !map.ContainsKey(key))
                map[key] = row;
        }
        return map;
    }

    // ── Git ─────────────────────────────────────────────────────────────────

    private static void GitAdd(string filePath)
    {
        var repoRoot = SvnGitTools.FindGitRoot(filePath);
        if (repoRoot == null) return;
        using var repo = new LibGit2Sharp.Repository(repoRoot);
        var relativePath = Path.GetRelativePath(repoRoot, filePath).Replace('\\', '/');
        repo.Index.Add(relativePath);
        repo.Index.Write();
    }
}
