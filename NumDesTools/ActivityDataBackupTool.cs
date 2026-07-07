using OfficeOpenXml;

namespace NumDesTools;

// 「大文件备份」表（#【A自动填表】创新活动【数值模板】.xlsm 的同名 Sheet）记录了每个创新活动在
// Type.xlsx / Icon.xlsx / Item.xlsx 三个大表里各自的起始/终止 id。右键该 Sheet 的某一行/整表：
//   删除：只删活号（正式表，不带 backup+时间戳），按 id 区间把整段行删掉。
//   还原：把 backup+时间戳表里同一 id 区间的数据同步回正式表——复用 XlsxCrossSync.ExecuteSync
//        （已存在 id 原地更新，缺失 id 按分组前缀插回原位置），只是把同步范围限定在这段 id 区间内。
internal static class ActivityDataBackupTool
{
    private const string BackupRootKey = "ActivityBackupRoot";
    private const string LiveRootKey = "ActivityLiveRoot";
    private const string DefaultLiveRoot = @"C:\M1Work\Public\Excels\Tables";
    private const string TemplateSheetName = "大文件备份";
    private const int TableNameRow = 1;
    private const int ColumnLabelRow = 2;
    private const int DataStartRow = 3;
    private static readonly string[] TrackedTables = { "Type.xlsx", "Icon.xlsx", "Item.xlsx" };

    internal static (string BackupRoot, string LiveRoot) LoadRoots()
    {
        var liveRoot = AppServices.GlobalValue.Value.GetValueOrDefault(
            LiveRootKey,
            DefaultLiveRoot
        );
        var backupRoot = AppServices.GlobalValue.Value.GetValueOrDefault(
            BackupRootKey,
            DeriveDefaultBackupRoot(liveRoot)
        );
        return (backupRoot, liveRoot);
    }

    // 没手动配过备份根目录时，跟正式表根目录联动推算（Excels\Tables → Excels_Backup\TablesBackup
    // 是同级目录，换机器/换环境时正式表根目录一变，备份根目录默认值跟着变，不用写死绝对路径）。
    private static string DeriveDefaultBackupRoot(string liveRoot)
    {
        const string liveMarker = @"Excels\Tables";
        const string backupMarker = @"Excels_Backup\TablesBackup";
        return liveRoot.Contains(liveMarker, StringComparison.OrdinalIgnoreCase)
            ? liveRoot.Replace(liveMarker, backupMarker, StringComparison.OrdinalIgnoreCase)
            : Path.Combine(Path.GetDirectoryName(liveRoot) ?? liveRoot, "TablesBackup");
    }

    internal static void SaveRoots(string backupRoot, string liveRoot)
    {
        AppServices.GlobalValue.SaveValue(BackupRootKey, backupRoot);
        AppServices.GlobalValue.SaveValue(LiveRootKey, liveRoot);
    }

    internal static void OpenSettings() => new UI.ActivityBackupSettingsWindow().Show();

    private const string StatusDeleted = "已删除";

    // internal：便于 NumDesTools.Tests 直接构造样例数据测试 ApplyDelete/ApplyRestore。
    // IsIgnored 对应「数据状态」左边那一列（Ignore），只要填了内容（不管填的是"半年"还是别的字），
    // 这个活动就跳过，不做删除/还原。
    // Status 是「数据状态」列当前的值：删除不限制状态；还原要求正好是"已删除"，防止拿备份盖掉正式表已改内容。
    internal sealed record Activity(
        int Row,
        string Id,
        string Name,
        Dictionary<string, (string Start, string End)> RangeByTable,
        bool IsIgnored = false,
        string Status = ""
    );

    // 删除不看状态：空的默认当作"还没删，可以删"；已经是"已删除"也不拦，走到 ApplyDelete 自然会因为
    // 找不到区间而报"没有可删除的东西"，不是错误。还原要求状态正好是"已删除"——不然会拿备份数据盖掉
    // 可能已经手动改过的正式表内容，这个必须拦。
    private static bool StatusAllows(Activity activity, bool delete) =>
        delete || activity.Status == StatusDeleted;

    public static void DeleteSelected_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true;
        RunOnSelected(delete: true);
    }

    public static void RestoreSelected_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true;
        RunOnSelected(delete: false);
    }

    public static void DeleteAll_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true;
        RunOnAll(delete: true);
    }

    public static void RestoreAll_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true;
        RunOnAll(delete: false);
    }

    public static void OpenSettings_Click(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true;
        OpenSettings();
    }

    private static void RunOnSelected(bool delete)
    {
        var activeRow = AppServices.App.ActiveCell.Row;
        var activities = LoadTemplateActivities();
        if (activities is null)
            return;

        var activity = activities.Find(a => a.Row == activeRow);
        if (activity is null)
        {
            MessageBox.Show(
                "当前选中行不是一条活动数据，请把光标定位到活动所在行再操作。",
                "大文件备份"
            );
            return;
        }
        if (activity.IsIgnored)
        {
            MessageBox.Show($"活动 {activity.Id} 的 Ignore 列有填值，忽略，不处理。", "大文件备份");
            return;
        }
        if (activity.RangeByTable.Count == 0)
        {
            MessageBox.Show(
                $"活动 {activity.Id} 没有填起始/终止id，无需处理，忽略。",
                "大文件备份"
            );
            return;
        }
        if (!StatusAllows(activity, delete))
        {
            var reason = string.IsNullOrEmpty(activity.Status)
                ? "数据状态是空的"
                : $"数据状态是「{activity.Status}」";
            MessageBox.Show(
                $"活动 {activity.Id} {reason}，还原要求数据状态正好是「已删除」，跳过（防止拿备份盖掉正式表里可能已经改过的内容）。",
                "大文件备份"
            );
            return;
        }

        Run(delete, new List<Activity> { activity });
    }

    private static void RunOnAll(bool delete)
    {
        var activities = LoadTemplateActivities();
        if (activities is null)
            return;
        // Ignore 列填了值的、没填起始/终止id的、数据状态不满足删除/还原前提的行都忽略，不当成错误。
        var toProcess = activities
            .Where(a => !a.IsIgnored && a.RangeByTable.Count > 0 && StatusAllows(a, delete))
            .ToList();
        if (toProcess.Count == 0)
        {
            MessageBox.Show("「大文件备份」表里没有需要处理的活动数据。", "大文件备份");
            return;
        }
        Run(delete, toProcess);
    }

    private static List<Activity>? LoadTemplateActivities()
    {
        var wb = AppServices.App.ActiveWorkbook;
        var templatePath = System.IO.Path.Combine(wb.Path, wb.Name);
        // 数据状态列是本工具自己用 COM 直接改活的 workbook（不强制存盘），
        // 从磁盘用 EPPlus 重新读这一列可能还是上一次存盘前的旧值，所以单独从 COM 读，跟别的字段区分开。
        var liveStatusSheet = wb.Worksheets[TemplateSheetName] as Worksheet;
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var pkg = new ExcelPackage(new FileInfo(templatePath));
        var sheet = pkg.Workbook.Worksheets[TemplateSheetName];
        if (sheet is null)
        {
            MessageBox.Show($"没找到「{TemplateSheetName}」Sheet。", "大文件备份");
            return null;
        }

        var idCol = FindColByText(sheet, ColumnLabelRow, "活动ID");
        var nameCol = FindColByText(sheet, ColumnLabelRow, "活动名称");
        if (idCol == -1 || nameCol == -1)
        {
            MessageBox.Show(
                $"「{TemplateSheetName}」第{ColumnLabelRow}行没找到\"活动ID\"/\"活动名称\"列，模板结构可能变了。",
                "大文件备份"
            );
            return null;
        }
        // Ignore 列没有就当没这个字段，不影响其它逻辑（不是所有版本的模板都一定有这一列）。
        var ignoreCol = FindColByText(sheet, ColumnLabelRow, "Ignore");

        var tableStartCols = new Dictionary<string, int>();
        foreach (var table in TrackedTables)
        {
            var col = FindColByText(sheet, TableNameRow, table);
            if (col != -1)
                tableStartCols[table] = col;
        }
        if (tableStartCols.Count == 0)
        {
            MessageBox.Show(
                $"「{TemplateSheetName}」第{TableNameRow}行没找到 {string.Join("/", TrackedTables)} 任何一个表名，模板结构可能变了。",
                "大文件备份"
            );
            return null;
        }

        var lastRow = sheet.Dimension?.End.Row ?? DataStartRow - 1;
        var activities = new List<Activity>();
        for (var r = DataStartRow; r <= lastRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim();
            if (string.IsNullOrEmpty(id))
                continue;
            var name = sheet.Cells[r, nameCol].Text?.Trim() ?? "";
            var isIgnored =
                ignoreCol != -1 && !string.IsNullOrEmpty(sheet.Cells[r, ignoreCol].Text?.Trim());
            var status = liveStatusSheet is not null
                ? (liveStatusSheet.Cells[r, 2].Text?.ToString()?.Trim() ?? "")
                : (sheet.Cells[r, 2].Text?.Trim() ?? "");
            var ranges = new Dictionary<string, (string Start, string End)>();
            foreach (var (table, startCol) in tableStartCols)
            {
                var start = sheet.Cells[r, startCol].Text?.Trim();
                var end = sheet.Cells[r, startCol + 1].Text?.Trim();
                if (!string.IsNullOrEmpty(start) && !string.IsNullOrEmpty(end))
                    ranges[table] = (start, end);
            }
            activities.Add(new Activity(r, id, name, ranges, isIgnored, status));
        }
        return activities;
    }

    private static int FindColByText(ExcelWorksheet sheet, int row, string text)
    {
        if (sheet.Dimension is null)
            return -1;
        for (var c = 1; c <= sheet.Dimension.End.Column; c++)
            if (sheet.Cells[row, c].Text?.Trim() == text)
                return c;
        return -1;
    }

    // 给「生成活动」功能用：只要 活动id→数据状态 的映射，不需要区间/Ignore 那些字段。
    // 「#【A自动填表】创新活动【数值模板】.xlsm」跟 ActivityServerData.xlsm 在同一个目录，
    // 传目录进来按文件名找，不用假设固定路径。找不到就返回空字典（调用方视为"查不到就不拦"）。
    internal static Dictionary<string, string> LoadActivityStatusById(string sameDir)
    {
        var result = new Dictionary<string, string>();
        var templatePath = Directory
            .EnumerateFiles(
                sameDir,
                "*A自动填表*创新活动*数值模板*.xlsm",
                SearchOption.TopDirectoryOnly
            )
            .FirstOrDefault();
        if (templatePath is null)
            return result;

        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var pkg = new ExcelPackage(new FileInfo(templatePath));
        var sheet = pkg.Workbook.Worksheets[TemplateSheetName];
        if (sheet?.Dimension is null)
            return result;

        var idCol = FindColByText(sheet, ColumnLabelRow, "活动ID");
        if (idCol == -1)
            return result;

        for (var r = DataStartRow; r <= sheet.Dimension.End.Row; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim();
            if (!string.IsNullOrEmpty(id))
                result[id] = sheet.Cells[r, 2].Text?.Trim() ?? "";
        }
        return result;
    }

    private static void Run(bool delete, List<Activity> activities)
    {
        var (backupRoot, liveRoot) = LoadRoots();
        if (
            !delete
            && (string.IsNullOrWhiteSpace(backupRoot) || string.IsNullOrWhiteSpace(liveRoot))
        )
        {
            MessageBox.Show(
                "还没有配置备份/正式表根目录，请先点右键菜单里的「大文件备份设置」。",
                "大文件备份"
            );
            return;
        }

        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var tableNotes = new List<string>();
        var previewLines = new List<string>();
        var plans =
            new List<(
                string Table,
                string LivePath,
                List<string> BackupCandidates,
                List<(Activity Activity, string Start, string End)> Ranges
            )>();

        foreach (var table in TrackedTables)
        {
            var ranges = activities
                .Where(a => a.RangeByTable.ContainsKey(table))
                .Select(a => (a, a.RangeByTable[table].Start, a.RangeByTable[table].End))
                .ToList();
            if (ranges.Count == 0)
                continue;

            var livePath = FindFileUnder(liveRoot, table);
            if (livePath is null)
            {
                tableNotes.Add($"[{table}] 在正式表根目录「{liveRoot}」下没找到，跳过。");
                continue;
            }

            var backupCandidates = new List<string>();
            if (!delete)
            {
                backupCandidates = FindAllBackups(backupRoot, table);
                if (backupCandidates.Count == 0)
                {
                    tableNotes.Add(
                        $"[{table}] 在备份根目录「{backupRoot}」下没找到对应备份文件，跳过。"
                    );
                    continue;
                }
            }

            plans.Add((table, livePath, backupCandidates, ranges));
        }

        if (plans.Count == 0)
        {
            MessageBox.Show(
                "没有可处理的表。\n" + string.Join("\n", tableNotes),
                delete ? "大文件备份 - 删除" : "大文件备份 - 还原"
            );
            return;
        }

        foreach (var (table, livePath, backupCandidates, ranges) in plans)
        {
            var label = string.Join(
                "、",
                ranges.Select(r => $"{r.Activity.Id}({r.Start}~{r.End})")
            );
            previewLines.Add(
                delete
                    ? $"[{table}] 删除 {ranges.Count} 个活动区间：{label}"
                    : $"[{table}] 从 {backupCandidates.Count} 份备份里按新到旧匹配，还原 {ranges.Count} 个活动区间：{label}"
            );
        }

        var confirmTitle = delete ? "大文件备份 - 删除预览" : "大文件备份 - 还原预览";
        var confirmBody =
            string.Join("\n", previewLines)
            + (tableNotes.Count > 0 ? "\n\n跳过：\n" + string.Join("\n", tableNotes) : "")
            + "\n\n"
            + (delete ? "确认后原地删除这些行" : "确认后原地覆写这些行")
            + "（git 可回溯）。是否继续？";
        if (
            MessageBox.Show(confirmBody, confirmTitle, MessageBoxButtons.OKCancel)
            != DialogResult.OK
        )
            return;

        var resultLines = new List<string>();
        foreach (var (table, livePath, backupCandidates, ranges) in plans)
        {
            resultLines.Add(
                delete
                    ? ApplyDelete(table, livePath, ranges)
                    : ApplyRestore(table, livePath, backupCandidates, ranges)
            );
        }

        UpdateStatusColumn(activities, delete ? "已删除" : "已还原");

        MessageBox.Show(
            string.Join("\n", resultLines),
            delete ? "大文件备份 - 删除完成" : "大文件备份 - 还原完成"
        );
    }

    // 「大文件备份」表第2列（B列）是数据状态字段：删除后标"已删除"，还原后标"已还原"。
    // 直接改当前打开的这个 workbook（COM），不重新读写 xlsm 文件本身，不会跟 Excel 抢文件锁。
    // 表头是空的才补一个"数据状态"，已有内容不覆盖——现在没法确认这一列原来的表头/取值约定，
    // 先按最保守的假设实现，跟约定不一致的话需要再调整。
    private static void UpdateStatusColumn(List<Activity> activities, string status)
    {
        if (AppServices.App.ActiveWorkbook.Worksheets[TemplateSheetName] is not Worksheet sheet)
            return;

        var header = sheet.Cells[ColumnLabelRow, 2].Text;
        if (string.IsNullOrWhiteSpace(header))
            sheet.Cells[ColumnLabelRow, 2].Value = "数据状态";

        foreach (var activity in activities)
            sheet.Cells[activity.Row, 2].Value = status;
    }

    private const string CommentColumnName = "#备注";

    internal static string ApplyDelete(
        string table,
        string livePath,
        List<(Activity Activity, string Start, string End)> ranges
    )
    {
        using var pkg = new ExcelPackage(new FileInfo(livePath));
        var sheet = pkg.Workbook.Worksheets["Sheet1"];
        var idCol = PubMetToExcel.FindSourceCol(
            sheet,
            XlsxCrossSync.HeaderRow,
            XlsxCrossSync.KeyColumnName
        );
        if (idCol == -1)
            return $"[{table}] 没找到 id 列，跳过。";
        var commentCol = PubMetToExcel.FindSourceCol(
            sheet,
            XlsxCrossSync.HeaderRow,
            CommentColumnName
        );

        var blocks = new List<(int Start, int End)>();
        var notFound = new List<string>();
        var mismatched = new List<string>();
        foreach (var (activity, start, end) in ranges)
        {
            var startRow = PubMetToExcel.FindSourceRow(sheet, idCol, start);
            var endRow = PubMetToExcel.FindSourceRow(sheet, idCol, end);
            if (startRow == -1 || endRow == -1 || endRow < startRow)
            {
                notFound.Add(activity.Id);
                continue;
            }
            var mismatch = FindRangeMismatch(
                sheet,
                idCol,
                commentCol,
                startRow,
                endRow,
                start,
                end
            );
            if (mismatch is not null)
            {
                mismatched.Add($"{activity.Id}：{mismatch}");
                continue;
            }
            blocks.Add((startRow, endRow));
        }

        // 从下往上删，避免前面删除后挪动了后面区块的行号
        blocks.Sort((a, b) => b.Start.CompareTo(a.Start));
        var deletedRows = 0;
        foreach (var (start, end) in blocks)
        {
            sheet.DeleteRow(start, end - start + 1);
            deletedRows += end - start + 1;
        }

        // 顺手清一次表尾残留格式（误 Ctrl+A 设过格式留下的"空白行"，跟这次删除本身无关，
        // 但既然文件已经开着要存盘了，搭这一趟比让用户另外跑一次 xlsx瘦身 便宜）。
        var (trimmedRows, _) = deletedRows > 0 ? XlsxSlimmer.TrimTrailingBlank(sheet) : (0, 0);

        if (
            (deletedRows > 0 || trimmedRows > 0)
            && !XlsxCrossSync.SaveWithFriendlyError(pkg, livePath, "大文件备份")
        )
            return $"[{table}] 保存失败（文件被占用）。";

        var notFoundNote =
            notFound.Count > 0 ? $"，找不到区间的活动：{string.Join("、", notFound)}" : "";
        var mismatchNote =
            mismatched.Count > 0
                ? $"，以下区间疑似混入其它活动数据，已跳过未删，请自行核实：{string.Join("；", mismatched)}"
                : "";
        var trimNote = trimmedRows > 0 ? $"，顺手清掉表尾 {trimmedRows} 行残留格式空行" : "";
        return $"[{table}] 删除 {deletedRows} 行{notFoundNote}{mismatchNote}{trimNote}";
    }

    // 区间内每一行理应都属于同一个活动：先看 id 分组前缀（跟 XlsxCrossSync 分组插入用的规则一致）
    // 是否跟起止id一致；前缀不一致时退一步比 #备注 的中文字符前缀（同活动的备注通常共享同一个中文名前缀）。
    // 两边都不一致才当成"混进了别的活动数据"，整段跳过不删——宁可少删，不能错删，交给用户自己核实。
    private static string? FindRangeMismatch(
        ExcelWorksheet sheet,
        int idCol,
        int commentCol,
        int startRow,
        int endRow,
        string startId,
        string endId
    )
    {
        var startPrefix = IdGroupPrefix(startId);
        var endPrefix = IdGroupPrefix(endId);
        var startComment =
            commentCol == -1 ? "" : ChineseNamePrefix(sheet.Cells[startRow, commentCol].Text);
        var endComment =
            commentCol == -1 ? "" : ChineseNamePrefix(sheet.Cells[endRow, commentCol].Text);

        for (var r = startRow; r <= endRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim() ?? "";
            var idPrefix = IdGroupPrefix(id);
            if (idPrefix == startPrefix || idPrefix == endPrefix)
                continue;

            if (commentCol == -1)
                return $"第{r}行 id={id} 前缀跟起止id不一致，且没有「{CommentColumnName}」列可核对";

            var comment = sheet.Cells[r, commentCol].Text?.Trim() ?? "";
            var commentPrefix = ChineseNamePrefix(comment);
            if (
                commentPrefix.Length > 0
                && (commentPrefix == startComment || commentPrefix == endComment)
            )
                continue;

            return $"第{r}行 id={id}（备注「{comment}」）前缀跟起止id及{CommentColumnName}都不一致";
        }
        return null;
    }

    private static string IdGroupPrefix(string id) =>
        id.Length >= XlsxCrossSync.GroupPrefixLen ? id[..XlsxCrossSync.GroupPrefixLen] : id;

    private static string ChineseNamePrefix(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return "";
        var i = 0;
        while (i < text.Length && text[i] >= 0x4E00 && text[i] <= 0x9FFF)
            i++;
        return text[..i];
    }

    // backupCandidates 按修改时间从新到旧排列；每个活动区间独立地从最新的备份开始试，
    // 该备份里没有这段 id 区间（比如那次备份之前活动就已经被删）就换下一份更旧的，直到找到或试完。
    internal static string ApplyRestore(
        string table,
        string livePath,
        List<string> backupCandidates,
        List<(Activity Activity, string Start, string End)> ranges
    )
    {
        using var livePkg = new ExcelPackage(new FileInfo(livePath));
        var liveSheet = livePkg.Workbook.Worksheets["Sheet1"];
        var liveCols = XlsxCrossSync.ReadHeaderColumns(liveSheet);

        var openBackups = new Dictionary<string, ExcelPackage>();
        ExcelPackage GetBackupPkg(string path)
        {
            if (!openBackups.TryGetValue(path, out var pkg))
                openBackups[path] = pkg = new ExcelPackage(new FileInfo(path));
            return pkg;
        }

        var totalUpdates = 0;
        var totalInserts = 0;
        var notFound = new List<string>();
        var fallbackUsed = new List<string>();
        try
        {
            foreach (var (activity, start, end) in ranges)
            {
                ExcelWorksheet? matchedSheet = null;
                var matchedRowStart = -1;
                var matchedRowEnd = -1;
                var matchedIndex = -1;
                for (var i = 0; i < backupCandidates.Count; i++)
                {
                    var sheet = GetBackupPkg(backupCandidates[i]).Workbook.Worksheets["Sheet1"];
                    var idCol = PubMetToExcel.FindSourceCol(
                        sheet,
                        XlsxCrossSync.HeaderRow,
                        XlsxCrossSync.KeyColumnName
                    );
                    if (idCol == -1)
                        continue;
                    var rowStart = PubMetToExcel.FindSourceRow(sheet, idCol, start);
                    var rowEnd = PubMetToExcel.FindSourceRow(sheet, idCol, end);
                    if (rowStart == -1 || rowEnd == -1 || rowEnd < rowStart)
                        continue;
                    matchedSheet = sheet;
                    matchedRowStart = rowStart;
                    matchedRowEnd = rowEnd;
                    matchedIndex = i;
                    break;
                }

                if (matchedSheet is null)
                {
                    notFound.Add(activity.Id);
                    continue;
                }
                if (matchedIndex > 0)
                    fallbackUsed.Add(
                        $"{activity.Id}用了{Path.GetFileName(backupCandidates[matchedIndex])}"
                    );

                var syncCols = XlsxCrossSync
                    .ReadHeaderColumns(matchedSheet)
                    .Intersect(liveCols)
                    .Where(c => c != XlsxCrossSync.KeyColumnName)
                    .ToList();
                var (updates, inserts, _) = XlsxCrossSync.ExecuteSync(
                    matchedSheet,
                    liveSheet,
                    XlsxCrossSync.KeyColumnName,
                    XlsxCrossSync.GroupPrefixLen,
                    syncCols,
                    preview: false,
                    matchedRowStart,
                    matchedRowEnd
                );
                totalUpdates += updates;
                totalInserts += inserts;
            }
        }
        finally
        {
            foreach (var pkg in openBackups.Values)
                pkg.Dispose();
        }

        if (
            (totalUpdates > 0 || totalInserts > 0)
            && !XlsxCrossSync.SaveWithFriendlyError(livePkg, livePath, "大文件备份")
        )
            return $"[{table}] 保存失败（文件被占用）。";

        var notFoundNote =
            notFound.Count > 0
                ? $"，所有备份里都找不到区间的活动：{string.Join("、", notFound)}"
                : "";
        var fallbackNote =
            fallbackUsed.Count > 0 ? $"（{string.Join("；", fallbackUsed)}用了较旧的备份）" : "";
        return $"[{table}] 还原：更新 {totalUpdates} 行 / 插回 {totalInserts} 行{notFoundNote}{fallbackNote}";
    }

    private static string? FindFileUnder(string root, string fileName) =>
        Directory.Exists(root)
            ? Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories).FirstOrDefault()
            : null;

    // backup 文件名约定 {stem}_backup_{日期}.xlsx，同一个表可能有多份不同日期的备份。
    // 按修改时间从新到旧排好序返回，ApplyRestore 会按这个顺序试，最新那份没有这段 id 区间就换下一份。
    private static List<string> FindAllBackups(string backupRoot, string liveFileName)
    {
        if (!Directory.Exists(backupRoot))
            return [];
        var stem = Path.GetFileNameWithoutExtension(liveFileName);
        var pattern = $"{stem}_backup_*{Path.GetExtension(liveFileName)}";
        return Directory
            .EnumerateFiles(backupRoot, pattern, SearchOption.AllDirectories)
            .Select(p => new FileInfo(p))
            .OrderByDescending(f => f.LastWriteTimeUtc)
            .Select(f => f.FullName)
            .ToList();
    }
}
