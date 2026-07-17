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
            if (!sheetDiff.HasConflict)
                continue;

            var sheet = pkg.Workbook.Worksheets[sheetDiff.SheetName];
            if (sheet?.Dimension == null)
                continue;

            // 建立 key → 行号 映射（表头在第2行，数据从第3行起）
            var keyColIdx = FindKeyColIndex(sheet);
            var keyToRow = BuildKeyRowMap(sheet, keyColIdx);

            var allCols = sheetDiff.AllColumns;

            // 1. 先处理 Modified（只改单元格值，不影响行号）
            foreach (var rc in sheetDiff.Rows)
            {
                if (rc.DiffType == RowDiffType.Modified)
                    ApplyModifiedRow(sheet, rc, keyToRow, allCols);
            }

            // 2. 按 THEIRS 顺序插入 OnlyTheirs 行（保留行间相对位置）
            InsertTheirsRowsInOrder(sheet, sheetDiff, keyToRow, allCols);

            // 3. 确保新增列的 row3(type) / row4(label) 从 THEIRS 文件元数据补全
            EnsureNewColsMeta(sheet, sheetDiff);

            // 4. 最后处理删除行（从大行号到小行号，避免移位影响前面的行号）
            var deleteRows = sheetDiff
                .Rows.Where(rc =>
                    rc.DiffType == RowDiffType.OnlyOurs && rc.RowChoice == ConflictChoice.Theirs
                )
                .Select(rc => keyToRow.TryGetValue(rc.RowKey, out var r) ? r : -1)
                .Where(r => r > 0)
                .OrderByDescending(r => r)
                .ToList();
            foreach (var delRow in deleteRows)
                sheet.DeleteRow(delRow);
        }

        pkg.Save();

        if (gitAdd)
            GitAdd(outPath);
    }

    // ── 定向修改 Modified 行 ────────────────────────────────────────────────

    private static void ApplyModifiedRow(
        ExcelWorksheet sheet,
        RowConflict rc,
        Dictionary<string, int> keyToRow,
        List<string> allColumns
    )
    {
        if (!keyToRow.TryGetValue(rc.RowKey, out var rowIdx))
            return;

        foreach (var cell in rc.Cells)
        {
            if (cell.Choice != ConflictChoice.Theirs)
                continue;

            var colIdx = FindOrCreateHeaderCol(sheet, cell.ColName, allColumns);
            // 冲突 diff 引擎内部按 string 比较，数字也会被读成字符串；写回时归一化一次，
            // 避免把本该是数字的值(如 id)固化成 sharedStrings 里的文本。
            CellValueNormalizer.ApplyTo(sheet.Cells[rowIdx, colIdx], cell.TheirsValue?.ToString());
        }
    }

    // ── 按 THEIRS 顺序在正确位置插入 OnlyTheirs 行 ───────────────────────────

    /// <summary>
    /// 按 THEIRS 原始顺序将 OnlyTheirs 行插入到正确位置，而不是追加到末尾。
    /// 使用 EPPlus InsertRow：物理插入新行，下方行整体下移，现有格式完整保留。
    /// 算法：沿 THEIRS 顺序维护"前驱 OURS 行号"游标；每次 InsertRow 后更新 keyToRow。
    /// </summary>
    private static void InsertTheirsRowsInOrder(
        ExcelWorksheet sheet,
        SheetDiff sheetDiff,
        Dictionary<string, int> keyToRow,
        List<string> allCols
    )
    {
        // 按 THEIRS 原始行索引排序（TheirsRowIndex >= 0 的行）
        var theirsOrdered = sheetDiff
            .Rows.Where(r => r.TheirsRowIndex >= 0)
            .OrderBy(r => r.TheirsRowIndex)
            .ToList();

        if (theirsOrdered.Count == 0)
            return;

        // 如果所有 OnlyTheirs 行都没有 TheirsRowIndex（旧数据兼容）则退回追加末尾
        if (theirsOrdered.All(r => r.DiffType == RowDiffType.OnlyTheirs && r.TheirsRowIndex < 0))
        {
            foreach (
                var rc in sheetDiff.Rows.Where(r =>
                    r.DiffType == RowDiffType.OnlyTheirs && r.RowChoice == ConflictChoice.Theirs
                )
            )
                LegacyAppendTheirsRow(sheet, rc, allCols);
            return;
        }

        // lastSharedOursRow：最近一个"共有行（Same/Modified）"在当前 sheet 中的行号
        // 初始值 4 = 最后一个表头行（数据从第 5 行开始），表示"所有共有行之前"
        int lastSharedOursRow = 4;

        foreach (var row in theirsOrdered)
        {
            if (row.DiffType != RowDiffType.OnlyTheirs)
            {
                // 共有行：更新前驱游标（keyToRow 已反映之前所有 InsertRow 的偏移）
                if (keyToRow.TryGetValue(row.RowKey, out var currentOursRow))
                    lastSharedOursRow = currentOursRow;
            }
            else if (row.RowChoice == ConflictChoice.Theirs && row.TheirsFullRow != null)
            {
                // OnlyTheirs 要保留：在前驱行紧后方插入
                var insertAt = lastSharedOursRow + 1;

                sheet.InsertRow(insertAt, 1);

                // 样式从上一行复制（同包内 StyleID 安全）
                // 不跨 Package 复制 THEIRS 样式：EPPlus StyleID 是包内索引，跨包赋值无效
                if (insertAt > 1)
                {
                    var colCount = sheet.Dimension?.End.Column ?? 1;
                    for (int col = 1; col <= colCount; col++)
                        sheet.Cells[insertAt, col].StyleID = sheet.Cells[insertAt - 1, col].StyleID;
                }

                // 写入 THEIRS 行数据
                foreach (var (header, val) in row.TheirsFullRow)
                {
                    if (string.IsNullOrEmpty(header))
                        continue;
                    var colIdx = FindOrCreateHeaderCol(sheet, header, allCols);
                    CellValueNormalizer.ApplyTo(sheet.Cells[insertAt, colIdx], val?.ToString());
                }

                // InsertRow 使 insertAt 及之后的所有行号 +1，同步更新 keyToRow
                foreach (var key in keyToRow.Keys.ToList())
                    if (keyToRow[key] >= insertAt)
                        keyToRow[key]++;

                // 下一个连续 OnlyTheirs 行插在刚插入行的正下方
                lastSharedOursRow = insertAt;
            }
        }
    }

    // 旧版追加到末尾（仅用于无 TheirsRowIndex 数据的旧 diff 兼容）
    private static void LegacyAppendTheirsRow(
        ExcelWorksheet sheet,
        RowConflict rc,
        List<string> allColumns
    )
    {
        if (rc.TheirsFullRow == null)
            return;

        var lastRow = sheet.Dimension.End.Row + 1;

        foreach (var (header, val) in rc.TheirsFullRow)
        {
            if (string.IsNullOrEmpty(header))
                continue;
            var colIdx = FindOrCreateHeaderCol(sheet, header, allColumns);
            CellValueNormalizer.ApplyTo(sheet.Cells[lastRow, colIdx], val?.ToString());
        }

        var srcRow = lastRow - 1;
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
            if (string.IsNullOrEmpty(colName))
                continue;

            if (
                sheet.Cells[3, col].Value == null
                && sheetDiff.TypeRow.TryGetValue(colName, out var t)
                && !string.IsNullOrEmpty(t)
            )
                sheet.Cells[3, col].Value = t;
            if (
                sheet.Cells[4, col].Value == null
                && sheetDiff.LabelRow.TryGetValue(colName, out var l)
                && !string.IsNullOrEmpty(l)
            )
                sheet.Cells[4, col].Value = l;
        }
    }

    // 在 row2 查找列名，找不到则按 SheetDiff.AllColumns 顺序插入到正确位置
    private static int FindOrCreateHeaderCol(
        ExcelWorksheet sheet,
        string colName,
        List<string>? allColumns = null
    )
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
                found:
                ;
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
        if (sheet.Dimension.End.Column >= 2)
            return 2;
        return 1;
    }

    private static Dictionary<string, int> BuildKeyRowMap(ExcelWorksheet sheet, int keyCol)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        // 与 ExcelConflictDiffer.dataStartRow 保持一致：行2=header，行3=type，行4=label，行5起=数据
        for (int row = 5; row <= sheet.Dimension.End.Row; row++)
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
        if (repoRoot == null)
            return;
        using var repo = new LibGit2Sharp.Repository(repoRoot);
        var relativePath = Path.GetRelativePath(repoRoot, filePath).Replace('\\', '/');
        repo.Index.Add(relativePath);
        repo.Index.Write();
        AppendMergeMsg(repoRoot, Path.GetFileName(filePath));
    }

    // 把解决的文件名写入当前 git 操作对应的消息文件，确保所有冲突解决场景都有日志
    public static void AppendMergeMsgPublic(string repoRoot, string fileName) =>
        AppendMergeMsg(repoRoot, fileName);

    private static void AppendMergeMsg(string repoRoot, string fileName)
    {
        try
        {
            var line = $"解决冲突（NumDesTools）: {fileName}";
            var git = Path.Combine(repoRoot, ".git");

            // 按 git 操作类型找对应的消息文件：
            //   merge / cherry-pick → MERGE_MSG
            //   rebase (merge mode) → rebase-merge/message
            //   rebase (apply mode) → rebase-apply/msg
            //   其他（兜底）        → 创建 MERGE_MSG
            var candidates = new[]
            {
                Path.Combine(git, "MERGE_MSG"),
                Path.Combine(git, "rebase-merge", "message"),
                Path.Combine(git, "rebase-apply", "msg"),
            };
            var target = candidates.FirstOrDefault(File.Exists) ?? candidates[0]; // 兜底：建 MERGE_MSG

            var existing = File.Exists(target) ? File.ReadAllText(target) : "";
            if (existing.Contains(line))
                return;

            string updated;
            if (!File.Exists(target))
            {
                updated = line;
            }
            else
            {
                // 插到第一个 "# " 注释行之前（# Conflicts: 块），让解决记录出现在输入框可见区
                var insertAt = existing.IndexOf("\n# ", StringComparison.Ordinal);
                updated =
                    insertAt >= 0 ? existing.Insert(insertAt, $"\n{line}") : existing + $"\n{line}";
            }
            File.WriteAllText(target, updated);

            PluginLog.Verbose($"[ConflictApplier] 冲突日志 → {Path.GetFileName(target)}: {line}");
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ConflictApplier] AppendMergeMsg 失败（非致命）: {ex.Message}");
        }
    }
}
