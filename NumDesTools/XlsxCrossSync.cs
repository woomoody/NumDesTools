using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace NumDesTools;

// a.xlsx ↔ b.xlsx 按 Key 列（约定列名 "id"）跨表同步，全自动：给定根目录 A/B，
// 按当前打开文件在哪侧根目录下推出对侧同名文件，逐个同名 Sheet 比较表头，
// 交集列（去掉 id）就同步，两侧都没有的列或没有 id 列的 Sheet 直接跳过。
// 插入位置复用 LTEData.cs 的分组算法：按 Key 前缀找同组最后一行插入其后。
internal static class XlsxCrossSync
{
    private const int HeaderRow = 2;
    private const int DataStartRow = 5;
    private const string KeyColumnName = "id";
    private const int GroupPrefixLen = 6;
    private const string RootAKey = "XlsxSyncRootA";
    private const string RootBKey = "XlsxSyncRootB";

    internal static (string RootA, string RootB) LoadRoots() =>
        (
            AppServices.GlobalValue.Value.GetValueOrDefault(RootAKey, ""),
            AppServices.GlobalValue.Value.GetValueOrDefault(RootBKey, "")
        );

    internal static void SaveRoots(string rootA, string rootB)
    {
        AppServices.GlobalValue.SaveValue(RootAKey, rootA);
        AppServices.GlobalValue.SaveValue(RootBKey, rootB);
    }

    internal static void OpenSettings() => new UI.XlsxSyncSettingsWindow().Show();

    internal static void RunForward() => RunSync(reverse: false);

    internal static void RunReverse() => RunSync(reverse: true);

    private static void RunSync(bool reverse)
    {
        var (rootA, rootB) = LoadRoots();
        if (string.IsNullOrWhiteSpace(rootA) || string.IsNullOrWhiteSpace(rootB))
        {
            MessageBox.Show("还没有配置根目录 A/B，请先点「同步设置」。", "跨表同步");
            return;
        }

        var wb = AppServices.App.ActiveWorkbook;
        if (wb is null || string.IsNullOrEmpty(wb.Path))
        {
            MessageBox.Show("请先打开一个 xlsx 文件。", "跨表同步");
            return;
        }

        var (sourceRoot, targetRoot) = reverse ? (rootB, rootA) : (rootA, rootB);
        var activePath = wb.FullName;
        if (!IsUnderRoot(activePath, sourceRoot))
        {
            MessageBox.Show(
                $"当前文件不在{(reverse ? "根目录 B" : "根目录 A")}下，请确认打开的是{(reverse ? "b" : "a")}侧文件。",
                "跨表同步"
            );
            return;
        }

        var relativeName = Path.GetRelativePath(sourceRoot, activePath);
        var targetPath = Path.Combine(targetRoot, relativeName);
        if (!File.Exists(targetPath))
        {
            MessageBox.Show($"对侧文件不存在：{targetPath}", "跨表同步");
            return;
        }

        try
        {
            ExecuteAutoSync(activePath, targetPath, reverse);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"同步失败：{ex.Message}", "跨表同步");
        }
    }

    private static bool IsUnderRoot(string path, string root)
    {
        var fullRoot =
            Path.GetFullPath(root)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return Path.GetFullPath(path).StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase);
    }

    private static HashSet<string> ReadHeaderColumns(ExcelWorksheet sheet)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (sheet.Dimension is null)
            return names;
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            var name = sheet.Cells[HeaderRow, col].Text?.Trim();
            if (!string.IsNullOrEmpty(name) && !name.StartsWith('#'))
                names.Add(name);
        }
        return names;
    }

    private static void ExecuteAutoSync(string fromPath, string toPath, bool reverse)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        using var fromPkg = new ExcelPackage(new FileInfo(fromPath));
        using var toPkg = new ExcelPackage(new FileInfo(toPath));

        var plans = new List<(ExcelWorksheet From, ExcelWorksheet To, List<string> Cols)>();
        foreach (var fromSheet in fromPkg.Workbook.Worksheets)
        {
            if (fromSheet.Name.StartsWith('#'))
                continue;
            var toSheet = toPkg.Workbook.Worksheets[fromSheet.Name];
            if (toSheet is null)
                continue;

            var fromCols = ReadHeaderColumns(fromSheet);
            var toCols = ReadHeaderColumns(toSheet);
            if (!fromCols.Contains(KeyColumnName) || !toCols.Contains(KeyColumnName))
                continue;

            var syncCols = fromCols.Intersect(toCols).Where(c => c != KeyColumnName).ToList();
            if (syncCols.Count == 0)
                continue;

            plans.Add((fromSheet, toSheet, syncCols));
        }

        if (plans.Count == 0)
        {
            MessageBox.Show(
                "没有找到可同步的同名 Sheet（需要两侧都有 id 列且至少一列同名）。",
                "跨表同步"
            );
            return;
        }

        var previews = new List<(string SheetName, int Updates, int Inserts, int Deletes)>();
        foreach (var (fromSheet, toSheet, cols) in plans)
        {
            var (u, i, d) = ExecuteSync(
                fromSheet,
                toSheet,
                KeyColumnName,
                GroupPrefixLen,
                cols,
                preview: true
            );
            previews.Add((fromSheet.Name, u, i, d));
        }

        var totalUpdates = previews.Sum(p => p.Updates);
        var totalInserts = previews.Sum(p => p.Inserts);
        var totalDeletes = previews.Sum(p => p.Deletes);
        if (totalUpdates == 0 && totalInserts == 0 && totalDeletes == 0)
        {
            MessageBox.Show("没有需要同步的差异。", "跨表同步");
            return;
        }

        var detail = string.Join(
            "\n",
            previews
                .Where(p => p.Updates + p.Inserts + p.Deletes > 0)
                .Select(p => $"[{p.SheetName}] 更新{p.Updates} / 新增{p.Inserts} / 删除{p.Deletes}")
        );
        var confirm = MessageBox.Show(
            $"{(reverse ? "反向 b→a" : "正向 a→b")}\n{detail}\n\n"
                + "新增行只会写入 id 列 + 同步列，其余列留空需自行补全。\n"
                + $"确认后原地覆写 {Path.GetFileName(toPath)}（git 可回溯）。是否继续？",
            "跨表同步 - 预览确认",
            MessageBoxButtons.OKCancel
        );
        if (confirm != DialogResult.OK)
            return;

        foreach (var (fromSheet, toSheet, cols) in plans)
            ExecuteSync(fromSheet, toSheet, KeyColumnName, GroupPrefixLen, cols, preview: false);

        if (!SaveWithFriendlyError(toPkg, toPath))
            return;

        MessageBox.Show(
            $"同步完成：更新 {totalUpdates} / 新增 {totalInserts} / 删除 {totalDeletes}",
            "跨表同步"
        );
    }

    // EPPlus 遇到 DeleteFile 失败时不是直接抛 IOException，是包成 InvalidOperationException（内部异常才是 IOException）。
    private static bool SaveWithFriendlyError(ExcelPackage pkg, string path)
    {
        try
        {
            pkg.Save();
            return true;
        }
        catch (InvalidOperationException ex) when (ex.InnerException is IOException)
        {
            MessageBox.Show(
                $"{Path.GetFileName(path)} 当前被其他程序占用（可能在 Excel 中打开），请关闭后重试。",
                "跨表同步"
            );
            return false;
        }
    }

    private static (int Updates, int Inserts, int Deletes) ExecuteSync(
        ExcelWorksheet source,
        ExcelWorksheet target,
        string keyColumnName,
        int groupPrefixLen,
        List<string> syncCols,
        bool preview
    )
    {
        var sourceKeyCol = PubMetToExcel.FindSourceCol(source, HeaderRow, keyColumnName);
        var targetKeyCol = PubMetToExcel.FindSourceCol(target, HeaderRow, keyColumnName);
        if (sourceKeyCol == -1 || targetKeyCol == -1)
            throw new InvalidOperationException(
                $"找不到 Key 列「{keyColumnName}」，请检查同步设置。"
            );

        var sourceColIdx = new Dictionary<string, int>();
        var targetColIdx = new Dictionary<string, int>();
        foreach (var col in syncCols)
        {
            var sc = PubMetToExcel.FindSourceCol(source, HeaderRow, col);
            var tc = PubMetToExcel.FindSourceCol(target, HeaderRow, col);
            if (sc == -1 || tc == -1)
                throw new InvalidOperationException($"找不到同步列「{col}」，请检查同步设置。");
            sourceColIdx[col] = sc;
            targetColIdx[col] = tc;
        }

        var targetKeyRow = new Dictionary<string, int>(StringComparer.Ordinal);
        if (target.Dimension is not null)
        {
            for (int r = DataStartRow; r <= target.Dimension.End.Row; r++)
            {
                var k = target.Cells[r, targetKeyCol].Text?.Trim();
                if (!string.IsNullOrEmpty(k))
                    targetKeyRow[k] = r;
            }
        }

        var targetEmpty = targetKeyRow.Count == 0;
        var sourceKeys = new HashSet<string>(StringComparer.Ordinal);
        var updateOps = new List<(int TargetRow, int SourceRow)>();
        var insertOps = new List<(string Key, int SourceRow)>();

        if (source.Dimension is not null)
        {
            for (int r = DataStartRow; r <= source.Dimension.End.Row; r++)
            {
                var k = source.Cells[r, sourceKeyCol].Text?.Trim();
                if (string.IsNullOrEmpty(k))
                    continue;
                sourceKeys.Add(k);
                if (targetKeyRow.TryGetValue(k, out var existingRow))
                    updateOps.Add((existingRow, r));
                else
                    insertOps.Add((k, r));
            }
        }

        var deleteRows = targetKeyRow
            .Where(kv => !sourceKeys.Contains(kv.Key))
            .Select(kv => kv.Value)
            .ToList();

        if (preview)
            return (updateOps.Count, insertOps.Count, deleteRows.Count);

        // 2a. 删除（倒序，修正后续行号）
        deleteRows.Sort((a, b) => b.CompareTo(a));
        foreach (var rowToDel in deleteRows)
        {
            target.DeleteRow(rowToDel);
            for (int i = 0; i < updateOps.Count; i++)
                if (updateOps[i].TargetRow > rowToDel)
                    updateOps[i] = (updateOps[i].TargetRow - 1, updateOps[i].SourceRow);
        }

        // 2b. 原地更新
        foreach (var (targetRow, sourceRow) in updateOps)
        foreach (var col in syncCols)
            target.Cells[targetRow, targetColIdx[col]].Value = source
                .Cells[sourceRow, sourceColIdx[col]]
                .Value;

        // 2c. 分组插入：按 Key 前缀找同组最后一行插入其后；找不到同组或目标为空则追加末尾
        var groupedInserts = new List<(int BaseRow, string Key, int SourceRow)>();
        var tailInserts = new List<(string Key, int SourceRow)>();
        foreach (var op in insertOps)
        {
            if (targetEmpty || op.Key.Length < groupPrefixLen)
            {
                tailInserts.Add(op);
                continue;
            }
            var regex = new Regex($"^{Regex.Escape(op.Key[..groupPrefixLen])}");
            var baseRow = PubMetToExcel.FindSourceRowBlur(target, targetKeyCol, regex);
            if (baseRow == -1)
                tailInserts.Add(op);
            else
                groupedInserts.Add((baseRow, op.Key, op.SourceRow));
        }
        groupedInserts.Sort((a, b) => a.BaseRow.CompareTo(b.BaseRow));

        int rowOffset = 0;
        foreach (var (baseRow, key, sourceRow) in groupedInserts)
        {
            int writeRow = baseRow + 1 + rowOffset;
            target.InsertRow(writeRow, 1);
            rowOffset++;
            target.Cells[writeRow, targetKeyCol].Value = key;
            foreach (var col in syncCols)
                target.Cells[writeRow, targetColIdx[col]].Value = source
                    .Cells[sourceRow, sourceColIdx[col]]
                    .Value;
        }

        foreach (var (key, sourceRow) in tailInserts)
        {
            int writeRow = (target.Dimension?.End.Row ?? DataStartRow - 1) + 1;
            target.Cells[writeRow, targetKeyCol].Value = key;
            foreach (var col in syncCols)
                target.Cells[writeRow, targetColIdx[col]].Value = source
                    .Cells[sourceRow, sourceColIdx[col]]
                    .Value;
        }

        return (updateOps.Count, groupedInserts.Count + tailInserts.Count, deleteRows.Count);
    }
}
