using OfficeOpenXml;

namespace NumDesTools.Sync;

// a.xlsx ↔ b.xlsx 按 Key 列（约定列名 "id"）跨表同步，全自动：给定根目录 A/B，
// 按当前打开文件算出对侧文件名（可选后缀），在对侧根目录下递归搜同名文件（不假设子文件夹结构一致），
// 逐个同名 Sheet 比较表头，没有 id 列的 Sheet 直接跳过。
// 两个方向机制不同：
//   a→b：交集列（去掉 id）整行同步——已存在 id 更新、新 id 插入整行、标记「删除」的行整行删掉。
//        插入位置复用 LTEData.cs 的分组算法：按 Key 前缀找同组最后一行插入其后。
//        若发现 b 有 a 没有的合法列，只弹提示，不处理（引导去跑 b→a）。
//   b→a：只把 b 有而 a 没有的列（字段结构，行1-4）搬到 a；该列数据只回填 a 中已存在的 id 行，
//        不新增行、不管 a 独有列——这是列结构补齐，不是行同步。
internal static class XlsxCrossSync
{
    internal const int HeaderRow = 2;
    internal const int DataStartRow = 5;
    internal const string KeyColumnName = "id";
    private const string DeleteMarker = "删除";
    internal const int GroupPrefixLen = 6;
    private const string RootAKey = "XlsxSyncRootA";
    private const string RootBKey = "XlsxSyncRootB";
    private const string SuffixAKey = "XlsxSyncSuffixA";

    // a 侧文件名可能比 b 侧多一个后缀（比如 ActivityClientData_Update.xlsx ↔ ActivityClientData.xlsx），
    // 后缀留空则要求两侧文件名完全一致。
    internal static (string RootA, string RootB, string SuffixA) LoadRoots() =>
        (
            AppServices.GlobalValue.Value.GetValueOrDefault(RootAKey, ""),
            AppServices.GlobalValue.Value.GetValueOrDefault(RootBKey, ""),
            AppServices.GlobalValue.Value.GetValueOrDefault(SuffixAKey, "")
        );

    internal static void SaveRoots(string rootA, string rootB, string suffixA)
    {
        AppServices.GlobalValue.SaveValue(RootAKey, rootA);
        AppServices.GlobalValue.SaveValue(RootBKey, rootB);
        AppServices.GlobalValue.SaveValue(SuffixAKey, suffixA);
    }

    internal static void OpenSettings() => new UI.XlsxSyncSettingsWindow().Show();

    internal static void RunForward() => RunSync(reverse: false);

    internal static void RunReverse() => RunSync(reverse: true);

    // SheetName 非空时仅处理该 Sheet；Row/Col 区间非空时仅同步落在区间内的 id 行 / 同步列。
    // 取代旧的「合并表格Row/Col」右键功能——选中整行/整列后再跑同步即可缩小范围。
    private readonly record struct SyncScope(
        string? SheetName,
        int? RowStart,
        int? RowEnd,
        int? ColStart,
        int? ColEnd
    );

    private static SyncScope BuildSelectionScope(Workbook wb)
    {
        if (
            wb.ActiveSheet is not Worksheet activeSheet
            || AppServices.App.Selection is not Range selection
        )
            return default;

        if (selection.Address == selection.EntireRow.Address)
            return new SyncScope(
                activeSheet.Name,
                selection.Row,
                selection.Row + selection.Rows.Count - 1,
                null,
                null
            );

        if (selection.Address == selection.EntireColumn.Address)
            return new SyncScope(
                activeSheet.Name,
                null,
                null,
                selection.Column,
                selection.Column + selection.Columns.Count - 1
            );

        return default;
    }

    private static void RunSync(bool reverse)
    {
        var (rootA, rootB, suffixA) = LoadRoots();
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

        var fileName = Path.GetFileName(activePath);
        var targetFileName = reverse
            ? AddSuffix(fileName, suffixA)
            : RemoveSuffix(fileName, suffixA);
        // 根目录下子文件夹名不一定对得上（比如 Table_Update ↔ Tables 不是简单去后缀能推出来的），
        // 按文件名在根目录下递归搜，不假设子路径结构一致。
        var targetPath = FindFileUnder(targetRoot, targetFileName);
        if (targetPath is null)
        {
            MessageBox.Show($"在「{targetRoot}」下没找到对侧文件：{targetFileName}", "跨表同步");
            return;
        }

        try
        {
            ExecuteAutoSync(activePath, targetPath, reverse, BuildSelectionScope(wb));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"同步失败：{ex.Message}", "跨表同步");
        }
    }

    private static string? FindFileUnder(string root, string fileName) =>
        Directory.Exists(root)
            ? Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories).FirstOrDefault()
            : null;

    private static bool IsUnderRoot(string path, string root)
    {
        var fullRoot =
            Path.GetFullPath(root)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return Path.GetFullPath(path).StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase);
    }

    private static string RemoveSuffix(string fileName, string suffix)
    {
        if (string.IsNullOrEmpty(suffix))
            return fileName;
        var name = Path.GetFileNameWithoutExtension(fileName);
        return name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
            ? name[..^suffix.Length] + Path.GetExtension(fileName)
            : fileName;
    }

    private static string AddSuffix(string fileName, string suffix) =>
        string.IsNullOrEmpty(suffix)
            ? fileName
            : Path.GetFileNameWithoutExtension(fileName) + suffix + Path.GetExtension(fileName);

    internal static HashSet<string> ReadHeaderColumns(ExcelWorksheet sheet)
    {
        var names = new HashSet<string>(StringComparer.Ordinal);
        if (sheet.Dimension is null)
            return names;
        for (var col = 2; col <= sheet.Dimension.End.Column; col++)
        {
            var name = sheet.Cells[HeaderRow, col].Text?.Trim();
            if (!string.IsNullOrEmpty(name))
                names.Add(name);
        }
        return names;
    }

    private static bool IsColumnInRange(
        ExcelWorksheet sheet,
        string colName,
        int colStart,
        int colEnd
    )
    {
        var idx = PubMetToExcel.FindSourceCol(sheet, HeaderRow, colName);
        return idx >= colStart && idx <= colEnd;
    }

    private static void ExecuteAutoSync(
        string fromPath,
        string toPath,
        bool reverse,
        SyncScope scope
    )
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        using var fromPkg = new ExcelPackage(new FileInfo(fromPath));
        using var toPkg = new ExcelPackage(new FileInfo(toPath));

        if (reverse)
        {
            ExecuteColumnSync(fromPkg, toPkg, toPath, scope);
            return;
        }

        var plans = new List<(ExcelWorksheet From, ExcelWorksheet To, List<string> Cols)>();
        var bOnlyBySheet = new List<(string SheetName, List<string> Cols)>();
        foreach (var fromSheet in fromPkg.Workbook.Worksheets)
        {
            if (fromSheet.Name.StartsWith('#'))
                continue;
            if (scope.SheetName is not null && fromSheet.Name != scope.SheetName)
                continue;
            var toSheet = toPkg.Workbook.Worksheets[fromSheet.Name];
            if (toSheet is null)
                continue;

            var fromCols = ReadHeaderColumns(fromSheet);
            var toCols = ReadHeaderColumns(toSheet);
            if (!fromCols.Contains(KeyColumnName) || !toCols.Contains(KeyColumnName))
                continue;

            var bOnly = toCols.Except(fromCols).Where(c => c != KeyColumnName).ToList();
            if (bOnly.Count > 0)
                bOnlyBySheet.Add((fromSheet.Name, bOnly));

            var syncCols = fromCols.Intersect(toCols).Where(c => c != KeyColumnName).ToList();
            if (scope.ColStart is not null)
                syncCols = syncCols
                    .Where(c =>
                        IsColumnInRange(fromSheet, c, scope.ColStart.Value, scope.ColEnd!.Value)
                    )
                    .ToList();
            if (syncCols.Count == 0)
                continue;

            plans.Add((fromSheet, toSheet, syncCols));
        }

        if (bOnlyBySheet.Count > 0)
            MessageBox.Show(
                "检测到 b 有 a 没有的合法列，本次 a→b 不会处理：\n"
                    + string.Join(
                        "\n",
                        bOnlyBySheet.Select(s => $"[{s.SheetName}] {string.Join("、", s.Cols)}")
                    )
                    + "\n\n建议先打开 b 侧文件执行「b→a」把这些列同步过去，再继续 a→b。",
                "跨表同步 - 发现 b 独有列"
            );

        if (plans.Count == 0)
        {
            MessageBox.Show(
                "没有找到可同步的同名 Sheet（需要两侧都有 id 列且至少一列同名）。",
                "跨表同步"
            );
            return;
        }

        var previews = new List<(string SheetName, int Updates, int Inserts, int Deleted)>();
        foreach (var (fromSheet, toSheet, cols) in plans)
        {
            var (u, i, d) = ExecuteSync(
                fromSheet,
                toSheet,
                KeyColumnName,
                GroupPrefixLen,
                cols,
                preview: true,
                scope.RowStart,
                scope.RowEnd
            );
            previews.Add((fromSheet.Name, u, i, d));
        }

        var totalUpdates = previews.Sum(p => p.Updates);
        var totalInserts = previews.Sum(p => p.Inserts);
        var totalDeleted = previews.Sum(p => p.Deleted);
        if (totalUpdates == 0 && totalInserts == 0 && totalDeleted == 0)
        {
            MessageBox.Show("没有需要同步的差异。", "跨表同步");
            return;
        }

        var detail = string.Join(
            "\n",
            previews
                .Where(p => p.Updates + p.Inserts + p.Deleted > 0)
                .Select(p => $"[{p.SheetName}] 更新{p.Updates} / 新增{p.Inserts} / 删除{p.Deleted}")
        );
        // 增量同步：源侧新 key 插入、已存在的 key 更新；不按 key 集合清理对侧独有行。
        // 唯一删除方式：源行第1列文本严格等于"删除"，且对侧存在同 key 时，删掉对侧那一整行。
        var scopeNote = scope.SheetName is null ? "" : "\n（当前选中了整行/整列，仅同步选中范围）";
        var confirm = MessageBox.Show(
            $"正向 a→b{scopeNote}\n{detail}\n\n"
                + "新增行只会写入 id 列 + 同步列，其余列留空需自行补全。\n"
                + "源第1列标注「删除」的行，若对侧存在同 key 会被整行删除。\n"
                + $"确认后原地覆写 {Path.GetFileName(toPath)}（git 可回溯）。是否继续？",
            "跨表同步 - 预览确认",
            MessageBoxButtons.OKCancel
        );
        if (confirm != DialogResult.OK)
            return;

        foreach (var (fromSheet, toSheet, cols) in plans)
            ExecuteSync(
                fromSheet,
                toSheet,
                KeyColumnName,
                GroupPrefixLen,
                cols,
                preview: false,
                scope.RowStart,
                scope.RowEnd
            );

        if (!SaveWithFriendlyError(toPkg, toPath))
            return;

        MessageBox.Show(
            $"同步完成：更新 {totalUpdates} / 新增 {totalInserts} / 删除 {totalDeleted}",
            "跨表同步"
        );
    }

    // b→a：只把 b 有而 a 没有的列（字段）搬到 a，数据只回填 a 中已存在的 id 行，不新增行、不管 a 独有列。
    private static void ExecuteColumnSync(
        ExcelPackage fromPkg,
        ExcelPackage toPkg,
        string toPath,
        SyncScope scope
    )
    {
        var plans =
            new List<(
                ExcelWorksheet From,
                ExcelWorksheet To,
                List<int> FromColIdx,
                List<string> Names
            )>();
        foreach (var fromSheet in fromPkg.Workbook.Worksheets)
        {
            if (fromSheet.Name.StartsWith('#'))
                continue;
            if (scope.SheetName is not null && fromSheet.Name != scope.SheetName)
                continue;
            var toSheet = toPkg.Workbook.Worksheets[fromSheet.Name];
            if (toSheet is null)
                continue;

            var fromCols = ReadHeaderColumns(fromSheet);
            var toCols = ReadHeaderColumns(toSheet);
            if (!fromCols.Contains(KeyColumnName) || !toCols.Contains(KeyColumnName))
                continue;

            var newColNames = fromCols.Except(toCols).Where(c => c != KeyColumnName).ToList();
            if (scope.ColStart is not null)
                newColNames = newColNames
                    .Where(c =>
                        IsColumnInRange(fromSheet, c, scope.ColStart.Value, scope.ColEnd!.Value)
                    )
                    .ToList();
            if (newColNames.Count == 0)
                continue;

            var fromColIdx = newColNames
                .Select(name => PubMetToExcel.FindSourceCol(fromSheet, HeaderRow, name))
                .ToList();
            plans.Add((fromSheet, toSheet, fromColIdx, newColNames));
        }

        if (plans.Count == 0)
        {
            MessageBox.Show("没有发现 b 独有的新列（a 已具备 b 的全部字段）。", "跨表同步 - b→a");
            return;
        }

        var sourceKeyCols = plans.ToDictionary(
            p => p.From.Name,
            p => PubMetToExcel.FindSourceCol(p.From, HeaderRow, KeyColumnName)
        );
        var targetKeyCols = plans.ToDictionary(
            p => p.From.Name,
            p => PubMetToExcel.FindSourceCol(p.To, HeaderRow, KeyColumnName)
        );
        var previews = plans
            .Select(p =>
                (
                    p.From.Name,
                    p.Names,
                    Matched: CountMatchingRows(
                        p.From,
                        p.To,
                        sourceKeyCols[p.From.Name],
                        targetKeyCols[p.From.Name]
                    )
                )
            )
            .ToList();

        var detail = string.Join(
            "\n",
            previews.Select(p =>
                $"[{p.Name}] 新增列: {string.Join("、", p.Names)}，可回填 {p.Matched} 行"
            )
        );
        var scopeNote = scope.SheetName is null ? "" : "\n（当前选中了整列，仅同步选中范围）";
        var confirm = MessageBox.Show(
            $"反向 b→a 列结构同步{scopeNote}\n{detail}\n\n"
                + "只新增列头（字段名/类型/标签），数据只回填 a 中已存在的 id 行，不新增行，a 独有列不受影响。\n"
                + $"确认后原地覆写 {Path.GetFileName(toPath)}（git 可回溯）。是否继续？",
            "跨表同步 - 预览确认",
            MessageBoxButtons.OKCancel
        );
        if (confirm != DialogResult.OK)
            return;

        foreach (var (fromSheet, toSheet, fromColIdx, names) in plans)
            ApplyColumnSync(fromSheet, toSheet, fromColIdx, names);

        if (!SaveWithFriendlyError(toPkg, toPath))
            return;

        MessageBox.Show(
            $"列同步完成：新增 {previews.Sum(p => p.Names.Count)} 列，共 {previews.Sum(p => p.Matched)} 行被回填",
            "跨表同步"
        );
    }

    // 统计 target 中有多少行的 id 在 source 里也存在（即会被回填的行数，仅预览用，不修改任何数据）。
    private static int CountMatchingRows(
        ExcelWorksheet source,
        ExcelWorksheet target,
        int sourceKeyCol,
        int targetKeyCol
    )
    {
        var sourceKeys = new HashSet<string>(StringComparer.Ordinal);
        if (source.Dimension is not null)
            for (int r = DataStartRow; r <= source.Dimension.End.Row; r++)
            {
                var k = source.Cells[r, sourceKeyCol].Text?.Trim();
                if (!string.IsNullOrEmpty(k))
                    sourceKeys.Add(k);
            }

        if (target.Dimension is null)
            return 0;
        var count = 0;
        for (int r = DataStartRow; r <= target.Dimension.End.Row; r++)
        {
            var k = target.Cells[r, targetKeyCol].Text?.Trim();
            if (!string.IsNullOrEmpty(k) && sourceKeys.Contains(k))
                count++;
        }
        return count;
    }

    // 在 target 末尾追加 source 独有列（行1-4：分组/字段名/类型/标签），并对 target 中已存在的 id 行回填数据。
    private static void ApplyColumnSync(
        ExcelWorksheet source,
        ExcelWorksheet target,
        List<int> fromColIndices,
        List<string> colNames
    )
    {
        var sourceKeyCol = PubMetToExcel.FindSourceCol(source, HeaderRow, KeyColumnName);
        var targetKeyCol = PubMetToExcel.FindSourceCol(target, HeaderRow, KeyColumnName);

        var sourceKeyRow = new Dictionary<string, int>(StringComparer.Ordinal);
        if (source.Dimension is not null)
            for (int r = DataStartRow; r <= source.Dimension.End.Row; r++)
            {
                var k = source.Cells[r, sourceKeyCol].Text?.Trim();
                if (!string.IsNullOrEmpty(k))
                    sourceKeyRow[k] = r;
            }

        var appendAt = target.Dimension?.End.Column ?? 1;
        var targetCols = new List<int>();
        foreach (var idx in fromColIndices)
        {
            appendAt++;
            for (var headerRow = 1; headerRow <= DataStartRow - 1; headerRow++)
                target.Cells[headerRow, appendAt].Value = source.Cells[headerRow, idx].Value;
            targetCols.Add(appendAt);
        }

        if (target.Dimension is null)
            return;
        for (int r = DataStartRow; r <= target.Dimension.End.Row; r++)
        {
            var k = target.Cells[r, targetKeyCol].Text?.Trim();
            if (string.IsNullOrEmpty(k) || !sourceKeyRow.TryGetValue(k, out var sr))
                continue;
            for (var i = 0; i < fromColIndices.Count; i++)
                target.Cells[r, targetCols[i]].Value = source.Cells[sr, fromColIndices[i]].Value;
        }
    }

    // EPPlus 遇到 DeleteFile 失败时不是直接抛 IOException，是包成 InvalidOperationException（内部异常才是 IOException）。
    internal static bool SaveWithFriendlyError(
        ExcelPackage pkg,
        string path,
        string caption = "跨表同步"
    )
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
                caption
            );
            return false;
        }
    }

    // 增量同步：源侧新 key 插入，已存在的 key 更新。不按 key 集合清理对侧独有行（对侧可能有不来自源的合法行）。
    // 唯一删除方式：源行第1列文本严格等于 DeleteMarker（"删除"），且对侧存在同 key 时，删掉对侧那一整行。
    // internal：ActivityDataBackupTool 复用同一套按 id 更新/分组插入逻辑做"还原"。
    internal static (int Updates, int Inserts, int Deleted) ExecuteSync(
        ExcelWorksheet source,
        ExcelWorksheet target,
        string keyColumnName,
        int groupPrefixLen,
        List<string> syncCols,
        bool preview,
        int? rowStart = null,
        int? rowEnd = null,
        HashSet<string>? keyFilter = null
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
        var updateOps = new List<(int TargetRow, int SourceRow)>();
        var insertOps = new List<(string Key, int SourceRow)>();
        var deleteRows = new List<int>();

        if (source.Dimension is not null)
        {
            for (int r = DataStartRow; r <= source.Dimension.End.Row; r++)
            {
                if (rowStart is not null && (r < rowStart || r > rowEnd))
                    continue;
                var k = source.Cells[r, sourceKeyCol].Text?.Trim();
                if (string.IsNullOrEmpty(k))
                    continue;
                if (keyFilter is not null && !keyFilter.Contains(k))
                    continue;
                var markedDeleted = source.Cells[r, 1].Text?.Trim() == DeleteMarker;
                if (targetKeyRow.TryGetValue(k, out var existingRow))
                {
                    if (markedDeleted)
                        deleteRows.Add(existingRow);
                    else
                        updateOps.Add((existingRow, r));
                }
                else if (!markedDeleted)
                {
                    insertOps.Add((k, r));
                }
            }
        }

        if (preview)
            return (updateOps.Count, insertOps.Count, deleteRows.Count);

        // 2a. 删除标记行（倒序，修正后续行号）
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
        // 锚点行一次性从 targetKeyRow 按前缀分组算出来（O(rows)），而不是每插入一行就整表反向扫一次
        // （O(inserts*rows)——插入几千行时是主要耗时来源，见「大文件备份」还原大表的性能问题）。
        // targetKeyRow 在插入循环开始前就已建好，跟旧的逐行 FindSourceRowBlur 扫的是同一份"插入前"
        // 的表状态，语义等价。
        var prefixLastRow = targetKeyRow
            .Where(kv => kv.Key.Length >= groupPrefixLen)
            .GroupBy(kv => kv.Key[..groupPrefixLen])
            .ToDictionary(g => g.Key, g => g.Max(kv => kv.Value));

        var groupedInserts = new List<(int BaseRow, string Key, int SourceRow)>();
        var tailInserts = new List<(string Key, int SourceRow)>();
        foreach (var op in insertOps)
        {
            if (
                targetEmpty
                || op.Key.Length < groupPrefixLen
                || !prefixLastRow.TryGetValue(op.Key[..groupPrefixLen], out var baseRow)
            )
            {
                tailInserts.Add(op);
                continue;
            }
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
