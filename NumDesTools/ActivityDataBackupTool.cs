using System.Text.RegularExpressions;
using System.Threading.Tasks;
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

    // 区间数超过这个门槛才把该表的删除/还原丢去并行跑：单表操作本身有独立开线程的开销，区间少的表
    // 串行更快，不值得为它多开一个线程。
    private const int ParallelThreshold = 10;
    private static readonly string[] TrackedTables = { "Type.xlsx", "Icon.xlsx", "Item.xlsx" };

    internal static bool IsTrackedTableFile(string fileName) =>
        TrackedTables.Any(table =>
            string.Equals(Path.GetFileName(fileName), table, StringComparison.OrdinalIgnoreCase)
        );

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

    // 没手动配过备份根目录时，跟正式表根目录联动推算（Excels\Tables → Excels_Backup\Tables_backup
    // 是同级目录，换机器/换环境时正式表根目录一变，备份根目录默认值跟着变，不用写死绝对路径）。
    private static string DeriveDefaultBackupRoot(string liveRoot)
    {
        const string liveMarker = @"Excels\Tables";
        const string backupMarker = @"Excels_Backup\Tables_backup";
        return liveRoot.Contains(liveMarker, StringComparison.OrdinalIgnoreCase)
            ? liveRoot.Replace(liveMarker, backupMarker, StringComparison.OrdinalIgnoreCase)
            : Path.Combine(Path.GetDirectoryName(liveRoot) ?? liveRoot, "Tables_backup");
    }

    internal static void SaveRoots(string backupRoot, string liveRoot)
    {
        AppServices.GlobalValue.SaveValue(BackupRootKey, backupRoot);
        AppServices.GlobalValue.SaveValue(LiveRootKey, liveRoot);
    }

    internal static void OpenSettings() => new UI.ActivityBackupSettingsWindow().Show();

    private const string StatusDeleted = "已删除";

    // internal：便于 NumDesTools.Tests 直接构造样例数据测试 ApplyDelete/ApplyRestore。
    // IsIgnored 对应「数据状态」左边那一列（Ignore），只要单元格里有任何非空字符，
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

    internal static bool HasIgnoreValue(string? text) => !string.IsNullOrWhiteSpace(text);

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
        var liveStatusSheet = wb.Worksheets[TemplateSheetName] as Worksheet;
        return LoadTemplateActivitiesFromFile(templatePath, liveStatusSheet);
    }

    // 方便诊断/测试直接从文件读模板，不依赖 ActiveWorkbook。
    // 数据状态列是本工具自己用 COM 直接改活的 workbook（不强制存盘），
    // 从磁盘用 EPPlus 重新读这一列可能还是上一次存盘前的旧值，所以单独从 COM 读，跟别的字段区分开。
    internal static List<Activity>? LoadTemplateActivitiesFromFile(
        string templatePath,
        Worksheet? liveStatusSheet
    )
    {
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
            var isIgnored = ignoreCol != -1 && HasIgnoreValue(sheet.Cells[r, ignoreCol].Text);
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

        // 新方案：只要 Item.xlsx 有区间，Icon/Type 也自动纳入处理（即使模板没填它们的起止id）
        var itemPlan = plans.Find(p => p.Table == "Item.xlsx");
        if (itemPlan.Table != null)
        {
            foreach (var table in new[] { "Type.xlsx", "Icon.xlsx" })
            {
                if (plans.Any(p => p.Table == table))
                    continue;
                var livePath = FindFileUnder(liveRoot, table);
                if (livePath is null)
                {
                    tableNotes.Add($"[{table}] 在正式表根目录下没找到，跳过。");
                    continue;
                }
                var backupCandidates = !delete ? FindAllBackups(backupRoot, table) : [];
                if (!delete && backupCandidates.Count == 0)
                {
                    tableNotes.Add($"[{table}] 在备份根目录下没找到备份文件，跳过。");
                    continue;
                }
                plans.Add((table, livePath, backupCandidates, itemPlan.Ranges));
            }
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
        if (!UI.ActivityBackupReportWindow.Confirm(confirmTitle, confirmBody))
            return;

        // 删除模式：为 Type/Icon/Item 构建 belongMapType 可删 id 集合，数据源用 live 表。
        // 还原模式：镜像同一套算法算"本该被还原"的id集合，只是数据源换成备份表——因为这批
        // id对应的live数据已经被删掉了，只能从备份里的Type.xlsx查belongMapType。两边用同一个
        // BuildDeletableIdSet，删除/还原才能完全对称闭合：删的是这批id，还原也精确还原这批id，
        // 不会把从来没被删过的id也拿备份数据去覆盖（可能覆盖掉live上更新的内容）。
        var deletableIdsByActivity = new Dictionary<Activity, HashSet<string>>();
        var preservedDetails = new List<string>();
        var itemPlanForIds = plans.Find(p => p.Table == "Item.xlsx");
        var typePlanForIds = plans.Find(p => p.Table == "Type.xlsx");
        if (itemPlanForIds.Table != null && itemPlanForIds.Ranges.Count > 0)
        {
            var itemIdSourcePath = delete
                ? FindFileUnder(liveRoot, "Item.xlsx")
                : itemPlanForIds.BackupCandidates.FirstOrDefault();
            var typeIdSourcePath = delete
                ? FindFileUnder(liveRoot, "Type.xlsx")
                : typePlanForIds.BackupCandidates?.FirstOrDefault();
            if (itemIdSourcePath is not null && typeIdSourcePath is not null)
            {
                (deletableIdsByActivity, preservedDetails) = BuildDeletableIdSet(
                    itemIdSourcePath,
                    typeIdSourcePath,
                    itemPlanForIds.Ranges
                );
            }
        }
        // 删除不区分活动，直接拍平成一个集合按 id 过滤删除即可。
        var deletableIds = deletableIdsByActivity.Values.SelectMany(ids => ids).ToHashSet();

        var planResults = new (string Summary, string MismatchDetail)[plans.Count];
        var parallelPlans = plans
            .Select((plan, index) => (plan, index))
            .Where(x => x.plan.Ranges.Count > ParallelThreshold)
            .ToList();
        var sequentialPlans = plans
            .Select((plan, index) => (plan, index))
            .Where(x => x.plan.Ranges.Count <= ParallelThreshold)
            .ToList();

        Parallel.ForEach(
            parallelPlans,
            item =>
            {
                var (table, livePath, backupCandidates, ranges) = item.plan;
                planResults[item.index] = delete
                    ? ApplyDeleteFiltered(table, livePath, ranges, deletableIds, preservedDetails)
                    : (
                        ApplyRestore(
                            table,
                            livePath,
                            backupCandidates,
                            ranges,
                            deletableIdsByActivity
                        ),
                        ""
                    );
            }
        );
        foreach (var (plan, index) in sequentialPlans)
        {
            var (table, livePath, backupCandidates, ranges) = plan;
            planResults[index] = delete
                ? ApplyDeleteFiltered(table, livePath, ranges, deletableIds, preservedDetails)
                : (
                    ApplyRestore(table, livePath, backupCandidates, ranges, deletableIdsByActivity),
                    ""
                );
        }

        var resultLines = planResults.Select(r => r.Summary).ToList();
        var mismatchDetails = planResults
            .Select(r => r.MismatchDetail)
            .Where(m => !string.IsNullOrEmpty(m))
            .ToList();

        // 续开链路检测
        if (delete)
        {
            var followUpWarnings = CheckFollowUpChainForActivities(
                liveRoot,
                activities.Select(a => a.Id).ToList()
            );
            if (followUpWarnings.Count > 0)
                mismatchDetails.AddRange(followUpWarnings);
        }

        UpdateStatusColumn(activities, delete ? "已删除" : "已还原");

        if (mismatchDetails.Count > 0)
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(string.Join("\n\n", mismatchDetails));
        }

        UI.ActivityBackupReportWindow.ShowResult(
            delete ? "大文件备份 - 删除完成" : "大文件备份 - 还原完成",
            string.Join("\n", resultLines)
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

    internal static (string Summary, string MismatchDetail) ApplyDelete(
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
            return ($"[{table}] 没找到 id 列，跳过。", "");
        var commentCol = FindDescriptiveHashCol(sheet, XlsxCrossSync.HeaderRow);

        var blocks = new List<(int Start, int End)>();
        var notFound = new List<string>();
        var mismatched = new List<string>();
        foreach (var (activity, start, end) in ranges)
        {
            var (startRow, endRow) = FindRowRangeByIdValue(sheet, idCol, start, end);
            if (startRow == -1 || endRow == -1)
            {
                notFound.Add(activity.Id);
                continue;
            }
            var mismatch = FindRangeMismatch(sheet, idCol, commentCol, startRow, endRow);
            if (mismatch is not null)
                mismatched.Add($"{activity.Id}（{start}~{end}）：{mismatch}");
            // 不再因为疑似混入就跳过整段——起止id本身常是预留占位行，误报率太高，删不删交给用户
            // 看日志自己判断，这里只管照常删，不做阻断。
            blocks.Add((startRow, endRow));
        }

        // 不用多次 sheet.DeleteRow：EPPlus 8.2.0~8.6.1（含最新版）在同一个打开的 ExcelWorksheet 上连续
        // 执行大量、跨度很大的 DeleteRow 调用时，内部行索引/偏移维护会出错——实测在236个真实区间上复现，
        // 某一段本该被"挤"上来的数据行会整体错位，中间留出一大段完全空白（详见
        // 2026-07-08-epplus-deleterow-bug-fix-task.md 的复现记录），Icon.xlsx 上出现过 2 万多行的假空白。
        // 改成一次性重建：只把"没被删除的行"按原顺序搬到新行号，最后裁掉多余的尾部行，完全绕开
        // DeleteRow 的偏移计算路径。只搬单元格的值，不搬格式/批注/超链接——ExcelRangeBase.Copy 会连带
        // 处理批注，实测遇到某些异常批注对象（Text 为 null）会直接抛异常，逐格搬值可以完全规避这条路径；
        // 目标行沿用它原来自己的格式，跟被删表整体的样式排布规律一致，不需要跟着值搬。
        var deleteRows = new HashSet<int>();
        foreach (var (start, end) in blocks)
            for (var r = start; r <= end; r++)
                deleteRows.Add(r);
        var deletedRows = deleteRows.Count;

        if (deletedRows > 0)
        {
            var lastRow = sheet.Dimension!.End.Row;
            var lastCol = sheet.Dimension.End.Column;
            var writeRow = 1;
            for (var readRow = 1; readRow <= lastRow; readRow++)
            {
                if (deleteRows.Contains(readRow))
                    continue;
                if (writeRow != readRow)
                    for (var col = 1; col <= lastCol; col++)
                        sheet.Cells[writeRow, col].Value = sheet.Cells[readRow, col].Value;
                writeRow++;
            }
            var newLastRow = writeRow - 1;
            if (newLastRow < lastRow)
                sheet.DeleteRow(newLastRow + 1, lastRow - newLastRow);
        }

        // 顺手清一次表尾残留格式（误 Ctrl+A 设过格式留下的"空白行"，跟这次删除本身无关，
        // 但既然文件已经开着要存盘了，搭这一趟比让用户另外跑一次 xlsx瘦身 便宜）。
        var (trimmedRows, _) = deletedRows > 0 ? XlsxSlimmer.TrimTrailingBlank(sheet) : (0, 0);

        if (
            (deletedRows > 0 || trimmedRows > 0)
            && !XlsxCrossSync.SaveWithFriendlyError(pkg, livePath, "大文件备份")
        )
            return ($"[{table}] 保存失败（文件被占用）。", "");

        var notFoundNote =
            notFound.Count > 0 ? $"，找不到区间的活动：{string.Join("、", notFound)}" : "";
        var mismatchNote =
            mismatched.Count > 0
                ? $"，{mismatched.Count} 个区间有id/说明核实提示（已照常删除），详情见错误日志面板"
                : "";
        var trimNote = trimmedRows > 0 ? $"，顺手清掉表尾 {trimmedRows} 行残留格式空行" : "";
        var summary = $"[{table}] 删除 {deletedRows} 行{notFoundNote}{mismatchNote}{trimNote}";
        var mismatchDetail =
            mismatched.Count > 0
                ? $"[{table}] 以下区间的id/说明跟区间内主流数据有差异（已照常删除，非阻断，多为起止占位行，也可能是混入嫌疑，请自行核实）：\n"
                    + string.Join("\n", mismatched)
                : "";
        return (summary, mismatchDetail);
    }

    // 从左到右找第一个文本以 # 开头且长度 > 1 的列，作为区间核实用的"说明列"。这一列在各表里命名
    // 并不统一（Type/Icon.xlsx 叫 #备注，Item.xlsx 叫 #竞品名称/#新品名称），唯一共同点是都带 # 前缀；
    // 纯 "#" 一个字符的那一列是行标记列，不是活动说明，长度过滤会自动跳过它。
    private static int FindDescriptiveHashCol(ExcelWorksheet sheet, int headerRow)
    {
        if (sheet.Dimension is null)
            return -1;
        for (var c = 1; c <= sheet.Dimension.End.Column; c++)
        {
            var text = sheet.Cells[headerRow, c].Text?.Trim();
            if (text?.Length > 1 && text[0] == '#')
                return c;
        }
        return -1;
    }

    // 拿区间内出现次数最多的 id分组前缀 / 说明列中文签名 / _sub_table_id 作基准，而不是起止id自己的
    // 前缀/备注/子表 id——起止id经常只是预留占位行，内容跟实际数据无关，拿它当基准会把绝大多数正常行
    // 误判成混入。只有三层都偏离"主流"的行才算疑似混入，返回描述交给调用方记日志，不阻塞删除。
    private static string? FindRangeMismatch(
        ExcelWorksheet sheet,
        int idCol,
        int commentCol,
        int startRow,
        int endRow
    )
    {
        var commentColName =
            commentCol == -1
                ? ""
                : sheet.Cells[XlsxCrossSync.HeaderRow, commentCol].Text?.Trim() ?? "";
        var subTableIdCol = sheet.Dimension is null
            ? -1
            : PubMetToExcel.FindSourceCol(sheet, XlsxCrossSync.HeaderRow, "_sub_table_id");

        var rows =
            new List<(
                int Row,
                string Id,
                string IdPrefix,
                string Comment,
                string Sig,
                string SubTableId
            )>();
        var idPrefixCounts = new Dictionary<string, int>();
        var sigCounts = new Dictionary<string, int>();
        var subTableIdCounts = new Dictionary<string, int>();
        for (var r = startRow; r <= endRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim() ?? "";
            var idPrefix = ActivityIdPrefix(id);
            var comment = commentCol == -1 ? "" : sheet.Cells[r, commentCol].Text?.Trim() ?? "";
            var sig = FirstChineseRun(comment);
            var subTableId =
                subTableIdCol == -1 ? "" : sheet.Cells[r, subTableIdCol].Text?.Trim() ?? "";
            rows.Add((r, id, idPrefix, comment, sig, subTableId));
            if (idPrefix.Length > 0)
                idPrefixCounts[idPrefix] = idPrefixCounts.GetValueOrDefault(idPrefix) + 1;
            if (sig.Length > 0)
                sigCounts[sig] = sigCounts.GetValueOrDefault(sig) + 1;
            if (subTableId.Length > 0)
                subTableIdCounts[subTableId] = subTableIdCounts.GetValueOrDefault(subTableId) + 1;
        }

        var majorIdPrefix =
            idPrefixCounts.Count > 0
                ? idPrefixCounts.OrderByDescending(kv => kv.Value).First().Key
                : "";
        var majorSig =
            sigCounts.Count > 0 ? sigCounts.OrderByDescending(kv => kv.Value).First().Key : "";
        var majorSubTableId =
            subTableIdCounts.Count > 0
                ? subTableIdCounts.OrderByDescending(kv => kv.Value).First().Key
                : "";

        foreach (var row in rows)
        {
            // id前缀一致就不用再查中文签名——同活动内不同道具的备注命名可能完全不搭（"阿拉丁"vs
            // "阿拉丁副本纪念品"、纯中文 vs "Lte-xxx-yyy"），id前缀已经能确认是同一活动就不用较真备注。
            if (row.IdPrefix == majorIdPrefix)
                continue;

            if (commentCol != -1 && SigMatches(row.Sig, majorSig))
                continue;

            if (
                subTableIdCol != -1
                && !string.IsNullOrEmpty(majorSubTableId)
                && row.SubTableId == majorSubTableId
            )
                continue;

            var isBoundary = row.Row == startRow || row.Row == endRow;

            if (commentCol == -1)
            {
                return isBoundary
                    ? $"第{row.Row}行 id={row.Id} 是起止id本身，前缀跟区间内主流id前缀「{majorIdPrefix}」不完全一致，大概率是预留占位行，非混入风险，仅供参考"
                    : $"第{row.Row}行 id={row.Id} 前缀跟区间内主流id前缀「{majorIdPrefix}」不一致，且这个表没有#开头的说明列可核对";
            }

            // 起止id往往是预留占位/特殊道具行（比如"XX副本纪念品"），跟主流数据不一致大概率不是
            // 混入，只是降级提示；区间中间的行才值得当真警告。
            return isBoundary
                ? $"第{row.Row}行 id={row.Id}（{commentColName}「{row.Comment}」）是起止id本身，前缀跟区间内主流id前缀「{majorIdPrefix}」及{commentColName}主流签名「{majorSig}」不完全一致——大概率是预留占位/特殊行，非混入风险，仅供参考"
                : $"第{row.Row}行 id={row.Id}（{commentColName}「{row.Comment}」）前缀跟区间内主流id前缀「{majorIdPrefix}」及{commentColName}主流签名「{majorSig}」都不一致";
        }
        return null;
    }

    // 只取前4位作"同一活动"的判定粒度，跟 XlsxCrossSync.GroupPrefixLen（6位，给跨表同步插入定位用）
    // 是两个不同语义，不能共用：活动id惯例是前4位标活动本身（比如 2101=阿拉丁、4506=某活动），
    // 第5-6位往往是活动内部的道具/预留分桶，用6位比会把同一活动的不同分桶误判成混入。
    private const int ActivityIdPrefixLen = 4;

    private static string ActivityIdPrefix(string id) =>
        id.Length >= ActivityIdPrefixLen ? id[..ActivityIdPrefixLen] : id;

    // 中文签名允许"谁是谁的前缀"就算匹配，不要求完全相等——同活动里有的道具备注是"阿拉丁-神灯1"
    // (签名"阿拉丁")，有的是"阿拉丁副本纪念品"(签名整段都是中文，没有分隔符可截断)，后者虽然比
    // 前者长，但确实包含前者，应该算同一活动。
    private static bool SigMatches(string sig, string majorSig) =>
        sig.Length > 0
        && majorSig.Length > 0
        && (
            sig.StartsWith(majorSig, StringComparison.Ordinal)
            || majorSig.StartsWith(sig, StringComparison.Ordinal)
        );

    // 找字符串里第一段连续的中文字符，不要求从第0个字符开始——老数据里常见"Lte-阿拉丁-神灯1"这种
    // 英文/数字前缀+中文名的格式，只从头找会永远匹配不到，必须跳过前面的非中文字符再开始找。
    private static string FirstChineseRun(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return "";
        var start = 0;
        while (start < text.Length && !IsChineseChar(text[start]))
            start++;
        var end = start;
        while (end < text.Length && IsChineseChar(text[end]))
            end++;
        return text[start..end];
    }

    private static bool IsChineseChar(char c) => c is >= (char)0x4E00 and <= (char)0x9FFF;

    // backupCandidates 按修改时间从新到旧排列。
    // deletableIdsByActivity 里有这个活动时（belongMapType 判定推导出来的可还原范围，跟删除时
    // 用的是同一套结果），按具体id集合还原——这跟删除完全对称闭合：删的是这批id，还原也精确
    // 还原这批id，不会把从来没被删过的id也拿备份数据去覆盖，可能覆盖掉live上更新的内容。
    // 某份备份缺了部分id就用它能覆盖的那些，剩下的id换更旧的备份接着找，直到找到或试完。
    // deletableIdsByActivity 里没有这个活动时（Icon/Type 自己配了独立起止区间的历史活动，不是
    // 从 Item.xlsx 推导出来的），走旧的整段行区间还原逻辑兜底。
    internal static string ApplyRestore(
        string table,
        string livePath,
        List<string> backupCandidates,
        List<(Activity Activity, string Start, string End)> ranges,
        Dictionary<Activity, HashSet<string>>? deletableIdsByActivity = null
    )
    {
        deletableIdsByActivity ??= [];
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
        var notFoundIdsByActivity = new List<string>();
        var notFoundActivities = new List<string>();
        var fallbackUsed = new List<string>();
        try
        {
            foreach (var (activity, start, end) in ranges)
            {
                if (
                    deletableIdsByActivity.TryGetValue(activity, out var targetIds)
                    && targetIds.Count > 0
                )
                {
                    var remaining = new HashSet<string>(targetIds, StringComparer.Ordinal);
                    for (var i = 0; i < backupCandidates.Count && remaining.Count > 0; i++)
                    {
                        var sheet = GetBackupPkg(backupCandidates[i]).Workbook.Worksheets["Sheet1"];
                        var idCol = PubMetToExcel.FindSourceCol(
                            sheet,
                            XlsxCrossSync.HeaderRow,
                            XlsxCrossSync.KeyColumnName
                        );
                        if (idCol == -1)
                            continue;
                        var foundHere = ScanIdsPresent(sheet, idCol, remaining);
                        if (foundHere.Count == 0)
                            continue;
                        var syncCols = XlsxCrossSync
                            .ReadHeaderColumns(sheet)
                            .Intersect(liveCols)
                            .Where(c => c != XlsxCrossSync.KeyColumnName)
                            .ToList();
                        var (updates, inserts, _) = XlsxCrossSync.ExecuteSync(
                            sheet,
                            liveSheet,
                            XlsxCrossSync.KeyColumnName,
                            XlsxCrossSync.GroupPrefixLen,
                            syncCols,
                            preview: false,
                            keyFilter: foundHere
                        );
                        totalUpdates += updates;
                        totalInserts += inserts;
                        if (i > 0)
                            fallbackUsed.Add(
                                $"{activity.Id}部分id用了{Path.GetFileName(backupCandidates[i])}"
                            );
                        remaining.ExceptWith(foundHere);
                    }
                    if (remaining.Count > 0)
                        notFoundIdsByActivity.Add(
                            $"{activity.Id}（{string.Join("、", remaining)}）"
                        );
                    continue;
                }

                // 旧逻辑兜底：不是从Item.xlsx推导出来的belongMapType可还原范围，按整段行区间还原。
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
                    var (rowStart, rowEnd) = FindRowRangeByIdValue(sheet, idCol, start, end);
                    if (rowStart == -1 || rowEnd == -1)
                        continue;
                    matchedSheet = sheet;
                    matchedRowStart = rowStart;
                    matchedRowEnd = rowEnd;
                    matchedIndex = i;
                    break;
                }

                if (matchedSheet is null)
                {
                    notFoundActivities.Add(activity.Id);
                    continue;
                }
                if (matchedIndex > 0)
                    fallbackUsed.Add(
                        $"{activity.Id}用了{Path.GetFileName(backupCandidates[matchedIndex])}"
                    );

                var legacySyncCols = XlsxCrossSync
                    .ReadHeaderColumns(matchedSheet)
                    .Intersect(liveCols)
                    .Where(c => c != XlsxCrossSync.KeyColumnName)
                    .ToList();
                var (legacyUpdates, legacyInserts, _) = XlsxCrossSync.ExecuteSync(
                    matchedSheet,
                    liveSheet,
                    XlsxCrossSync.KeyColumnName,
                    XlsxCrossSync.GroupPrefixLen,
                    legacySyncCols,
                    preview: false,
                    matchedRowStart,
                    matchedRowEnd
                );
                totalUpdates += legacyUpdates;
                totalInserts += legacyInserts;
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
            notFoundIdsByActivity.Count > 0
                ? $"，以下id在本表所有备份里都不存在（多为本表本身不覆盖该id，比如不是所有Item都需要Icon数据，通常非异常，仅供核对）：{string.Join("、", notFoundIdsByActivity)}"
                : "";
        var notFoundActivityNote =
            notFoundActivities.Count > 0
                ? $"，所有备份里都找不到区间的活动：{string.Join("、", notFoundActivities)}"
                : "";
        var fallbackNote =
            fallbackUsed.Count > 0 ? $"（{string.Join("；", fallbackUsed)}用了较旧的备份）" : "";
        return $"[{table}] 还原：更新 {totalUpdates} 行 / 插回 {totalInserts} 行{notFoundNote}{notFoundActivityNote}{fallbackNote}";
    }

    private static HashSet<string> ScanIdsPresent(
        ExcelWorksheet sheet,
        int idCol,
        HashSet<string> candidateIds
    )
    {
        var found = new HashSet<string>(StringComparer.Ordinal);
        var lastRow = sheet.Dimension?.End.Row ?? 0;
        for (var r = XlsxCrossSync.HeaderRow + 1; r <= lastRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim();
            if (!string.IsNullOrEmpty(id) && candidateIds.Contains(id))
                found.Add(id);
        }
        return found;
    }

    // 用起止id各自的行号确定区间，取小的作为rowStart、大的作为rowEnd——不能按id数值大小
    // 排序：实测同一个活动的id段经常不是严格数值升序插入的（同一批行号连续的区间里，id会
    // 忽大忽小交替出现，比如8002172501之后紧跟8002174050），起始id对应的行号在前、结束id
    // 对应的行号在后才是真正可靠的语义，数值大小顺序不能当依据。
    // Icon.xlsx这类表本身id分布不连续（不是每个Item都配图标），Start/End两个精确id有一个
    // 缺失就要整段判"找不到"太严格，所以只要求两个id本身都能在表里精确定位到即可，行号顺序
    // 由代码里的min/max保证正确，不依赖模板里两列谁填在前面。
    private static (int RowStart, int RowEnd) FindRowRangeByIdValue(
        ExcelWorksheet sheet,
        int idCol,
        string start,
        string end
    )
    {
        var startRow = PubMetToExcel.FindSourceRow(sheet, idCol, start);
        var endRow = PubMetToExcel.FindSourceRow(sheet, idCol, end);
        if (startRow == -1 || endRow == -1)
            return (-1, -1);
        return (Math.Min(startRow, endRow), Math.Max(startRow, endRow));
    }

    internal static string? FindFileUnder(string root, string fileName) =>
        Directory.Exists(root)
            ? Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories).FirstOrDefault()
            : null;

    // backup 文件名约定 {stem}_backup_{yyyy-M-d}.xlsx（月份/日期不强制补 0），同一个表可能有多份不同日期的备份。
    // 这里也兼容老的零补齐命名，按修改时间从新到旧排好序返回，ApplyRestore 会按这个顺序试，最新那份没有
    // 这段 id 区间就换下一份。
    internal static List<string> FindAllBackups(string backupRoot, string liveFileName)
    {
        if (!Directory.Exists(backupRoot))
            return [];
        var stem = Path.GetFileNameWithoutExtension(liveFileName);
        var extension = Path.GetExtension(liveFileName);
        return Directory
            .EnumerateFiles(backupRoot, $"{stem}_backup_*{extension}", SearchOption.AllDirectories)
            .Where(path =>
            {
                var fileName = Path.GetFileNameWithoutExtension(path);
                return Regex.IsMatch(
                        fileName,
                        $"^{Regex.Escape(stem)}_backup_\\d{{4}}-\\d{{1,2}}-\\d{{1,2}}$",
                        RegexOptions.IgnoreCase
                    )
                    || Regex.IsMatch(
                        fileName,
                        $"^{Regex.Escape(stem)}_backup_\\d{{4}}-\\d{{2}}-\\d{{2}}$",
                        RegexOptions.IgnoreCase
                    );
            })
            .Select(p => new FileInfo(p))
            .OrderByDescending(f => f.LastWriteTimeUtc)
            .Select(f => f.FullName)
            .ToList();
    }

    // 按活动分组返回可删/可还原的 id 集合——删除时用 live 数据源，还原时用备份数据源，两边跑
    // 同一套 belongMapType 判定算法，不额外存一份"曾被删除的id清单"，删除/还原完全对称闭合。
    // 按活动分组（而不是拍平成一个集合）是因为 ApplyRestore 还原时要知道"这个活动该还原哪些
    // 具体id"，落不到任何活动名下的历史遗留数据（Icon/Type 自己配了独立起止区间、不是从
    // Item.xlsx 推导出来的）不在这个函数处理范围内，那部分继续走旧的按行区间还原逻辑。
    internal static (
        Dictionary<Activity, HashSet<string>> DeletableIdsByActivity,
        List<string> PreservedDetails
    ) BuildDeletableIdSet(
        string itemLivePath,
        string typeLivePath,
        List<(Activity Activity, string Start, string End)> itemRanges
    )
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var deletableByActivity = new Dictionary<Activity, HashSet<string>>();
        var preserved = new List<string>();

        var itemIdsByActivity = new Dictionary<Activity, HashSet<string>>();
        var allItemIds = new HashSet<string>();
        using (var itemPkg = new ExcelPackage(new FileInfo(itemLivePath)))
        {
            var itemSheet = itemPkg.Workbook.Worksheets["Sheet1"];
            var itemIdCol = PubMetToExcel.FindSourceCol(
                itemSheet,
                XlsxCrossSync.HeaderRow,
                XlsxCrossSync.KeyColumnName
            );
            if (itemIdCol == -1)
                return (deletableByActivity, preserved);
            // 全量删除/还原时活动数可能有几百个，逐活动调 FindSourceRow（线性扫描）两次
            // 会是 O(活动数 × 行数)，几百活动 × 十几万行代价很高。改成一次扫描建
            // id→行号索引，之后按id查行号是 O(1)，总代价降到 O(行数 + 活动数)。
            var lastItemRow = itemSheet.Dimension?.End.Row ?? 0;
            var itemRowById = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var r = XlsxCrossSync.HeaderRow + 1; r <= lastItemRow; r++)
            {
                var rowId = itemSheet.Cells[r, itemIdCol].Text?.Trim();
                if (!string.IsNullOrEmpty(rowId))
                    itemRowById[rowId] = r;
            }
            foreach (var (activity, start, end) in itemRanges)
            {
                if (
                    !itemRowById.TryGetValue(start, out var startRow)
                    || !itemRowById.TryGetValue(end, out var endRow)
                    || endRow < startRow
                )
                    continue;
                var idsForActivity = itemIdsByActivity.TryGetValue(activity, out var existing)
                    ? existing
                    : itemIdsByActivity[activity] = [];
                for (var r = startRow; r <= endRow; r++)
                {
                    var id = itemSheet.Cells[r, itemIdCol].Text?.Trim();
                    if (string.IsNullOrEmpty(id))
                        continue;
                    idsForActivity.Add(id);
                    allItemIds.Add(id);
                }
            }
        }
        if (allItemIds.Count == 0)
            return (deletableByActivity, preserved);

        var typeBelongMap = new Dictionary<string, string>();
        using (var typePkg = new ExcelPackage(new FileInfo(typeLivePath)))
        {
            var typeSheet = typePkg.Workbook.Worksheets["Sheet1"];
            var typeIdCol = PubMetToExcel.FindSourceCol(
                typeSheet,
                XlsxCrossSync.HeaderRow,
                XlsxCrossSync.KeyColumnName
            );
            if (typeIdCol == -1)
            {
                preserved.Add("Type.xlsx 中没找到 id 列");
                return (deletableByActivity, preserved);
            }
            var bmtCol = FindColByText(typeSheet, XlsxCrossSync.HeaderRow, "belongMapType");
            if (bmtCol == -1)
            {
                preserved.Add("Type.xlsx 中没找到 belongMapType 列");
                return (deletableByActivity, preserved);
            }
            var lastRow = typeSheet.Dimension?.End.Row ?? 0;
            for (var r = XlsxCrossSync.HeaderRow + 1; r <= lastRow; r++)
            {
                var id = typeSheet.Cells[r, typeIdCol].Text?.Trim();
                if (string.IsNullOrEmpty(id) || !allItemIds.Contains(id))
                    continue;
                typeBelongMap[id] = typeSheet.Cells[r, bmtCol].Text?.Trim() ?? "";
            }
        }
        foreach (var (activity, itemIds) in itemIdsByActivity)
        foreach (var id in itemIds)
        {
            if (!typeBelongMap.TryGetValue(id, out var bmt))
            {
                preserved.Add(
                    $"id={id}：Type.xlsx 中不存在该id，无法判断所属场景，Item/Icon 中保留该id不删"
                );
            }
            else if (string.IsNullOrEmpty(bmt))
            {
                preserved.Add(
                    $"id={id}：belongMapType 为空，无法判断所属场景，Item/Icon 中保留该id不删"
                );
            }
            else if (bmt == "[4]")
            {
                (
                    deletableByActivity.TryGetValue(activity, out var existing)
                        ? existing
                        : deletableByActivity[activity] = []
                ).Add(id);
            }
            else
            {
                preserved.Add($"id={id}：belongMapType={bmt}（非纯[4]），Item/Icon 中保留该id不删");
            }
        }
        return (deletableByActivity, preserved);
    }

    internal static (string Summary, string Detail) ApplyDeleteWithIdFilter(
        string table,
        string livePath,
        HashSet<string> deletableIds,
        List<string> preservedDetails
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
            return ($"[{table}] 没找到 id 列，跳过。", "");

        var deleteRows = new HashSet<int>();
        var lastRow = sheet.Dimension?.End.Row ?? 0;
        for (var r = XlsxCrossSync.HeaderRow + 1; r <= lastRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim();
            if (!string.IsNullOrEmpty(id) && deletableIds.Contains(id))
                deleteRows.Add(r);
        }
        var deletedRows = deleteRows.Count;
        if (deletedRows > 0)
        {
            var lastCol = sheet.Dimension!.End.Column;
            var writeRow = 1;
            for (var readRow = 1; readRow <= lastRow; readRow++)
            {
                if (deleteRows.Contains(readRow))
                    continue;
                if (writeRow != readRow)
                    for (var col = 1; col <= lastCol; col++)
                        sheet.Cells[writeRow, col].Value = sheet.Cells[readRow, col].Value;
                writeRow++;
            }
            var newLastRow = writeRow - 1;
            if (newLastRow < lastRow)
                sheet.DeleteRow(newLastRow + 1, lastRow - newLastRow);
        }
        var (trimmedRows, _) = deletedRows > 0 ? XlsxSlimmer.TrimTrailingBlank(sheet) : (0, 0);
        if (
            (deletedRows > 0 || trimmedRows > 0)
            && !XlsxCrossSync.SaveWithFriendlyError(pkg, livePath, "大文件备份")
        )
            return ($"[{table}] 保存失败（文件被占用）。", "");

        var preservedNote =
            preservedDetails.Count > 0
                ? $"，保留 {preservedDetails.Count} 个 id（belongMapType 非纯[4] 或查不到）"
                : "";
        var trimNote = trimmedRows > 0 ? $"，清掉表尾 {trimmedRows} 行残留" : "";
        var summary = $"[{table}] 删除 {deletedRows} 行{preservedNote}{trimNote}";
        var detail =
            preservedDetails.Count > 0
                ? $"[{table}] 保留的 id：\n" + string.Join("\n", preservedDetails)
                : "";
        return (summary, detail);
    }

    internal static (string, string) ApplyDeleteFiltered(
        string table,
        string livePath,
        List<(Activity Activity, string Start, string End)> ranges,
        HashSet<string> deletableIds,
        List<string> preservedDetails
    )
    {
        if (deletableIds.Count == 0)
            return ApplyDelete(table, livePath, ranges);
        return ApplyDeleteWithIdFilter(table, livePath, deletableIds, preservedDetails);
    }

    private static List<string> CheckFollowUpChainForActivities(
        string liveRoot,
        List<string> activityIds
    )
    {
        var warnings = new List<string>();
        var batchSet = new HashSet<string>(activityIds);
        try
        {
            var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(liveRoot);
            foreach (var id in activityIds)
            {
                if (predMap.TryGetValue(id, out var targets) && targets.Count > 0)
                {
                    var missing = targets.Where(t => !batchSet.Contains(t)).Distinct().ToList();
                    if (missing.Count > 0)
                        warnings.Add(
                            $"续开链路警告：活动 {id} 是续开前驱，其续开目标 {string.Join("、", missing)} 未同批处理，续开可能触发这些活动重读 Type/Icon/Item，请确认数据完整"
                        );
                }
            }
        }
        catch { }
        return warnings;
    }
}
