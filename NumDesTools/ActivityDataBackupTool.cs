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
        if (!UI.ActivityBackupReportWindow.Confirm(confirmTitle, confirmBody))
            return;

        var resultLines = new List<string>();
        var mismatchDetails = new List<string>();
        foreach (var (table, livePath, backupCandidates, ranges) in plans)
        {
            if (delete)
            {
                var (summary, mismatchDetail) = ApplyDelete(table, livePath, ranges);
                resultLines.Add(summary);
                if (!string.IsNullOrEmpty(mismatchDetail))
                    mismatchDetails.Add(mismatchDetail);
            }
            else
            {
                resultLines.Add(ApplyRestore(table, livePath, backupCandidates, ranges));
            }
        }

        UpdateStatusColumn(activities, delete ? "已删除" : "已还原");

        // 混入数据的详情单独走 CTP 日志面板列举，不塞进结果弹窗打断操作；三个表都跑完才统一写一次，
        // 避免逐表调用互相覆盖同一个全局 CTP 面板。
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
            var startRow = PubMetToExcel.FindSourceRow(sheet, idCol, start);
            var endRow = PubMetToExcel.FindSourceRow(sheet, idCol, end);
            if (startRow == -1 || endRow == -1 || endRow < startRow)
            {
                notFound.Add(activity.Id);
                continue;
            }
            var mismatch = FindRangeMismatch(sheet, idCol, commentCol, startRow, endRow);
            if (mismatch is not null)
                mismatched.Add($"{activity.Id}：{mismatch}");
            // 不再因为疑似混入就跳过整段——起止id本身常是预留占位行，误报率太高，删不删交给用户
            // 看日志自己判断，这里只管照常删，不做阻断。
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

    // 拿区间内出现次数最多的 id分组前缀 / 说明列中文签名作基准，而不是起止id自己的前缀/备注——
    // 起止id经常只是预留占位行，内容跟实际数据无关，拿它当基准会把绝大多数正常行误判成混入。
    // 只有 id前缀、中文签名都偏离"主流"的行才算疑似混入，返回描述交给调用方记日志，不阻塞删除。
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

        var rows = new List<(int Row, string Id, string IdPrefix, string Comment, string Sig)>();
        var idPrefixCounts = new Dictionary<string, int>();
        var sigCounts = new Dictionary<string, int>();
        for (var r = startRow; r <= endRow; r++)
        {
            var id = sheet.Cells[r, idCol].Text?.Trim() ?? "";
            var idPrefix = ActivityIdPrefix(id);
            var comment = commentCol == -1 ? "" : sheet.Cells[r, commentCol].Text?.Trim() ?? "";
            var sig = FirstChineseRun(comment);
            rows.Add((r, id, idPrefix, comment, sig));
            if (idPrefix.Length > 0)
                idPrefixCounts[idPrefix] = idPrefixCounts.GetValueOrDefault(idPrefix) + 1;
            if (sig.Length > 0)
                sigCounts[sig] = sigCounts.GetValueOrDefault(sig) + 1;
        }

        var majorIdPrefix =
            idPrefixCounts.Count > 0
                ? idPrefixCounts.OrderByDescending(kv => kv.Value).First().Key
                : "";
        var majorSig =
            sigCounts.Count > 0 ? sigCounts.OrderByDescending(kv => kv.Value).First().Key : "";

        foreach (var row in rows)
        {
            // id前缀一致就不用再查中文签名——同活动内不同道具的备注命名可能完全不搭（"阿拉丁"vs
            // "阿拉丁副本纪念品"、纯中文 vs "Lte-xxx-yyy"），id前缀已经能确认是同一活动就不用较真备注。
            if (row.IdPrefix == majorIdPrefix)
                continue;

            var isBoundary = row.Row == startRow || row.Row == endRow;

            if (commentCol == -1)
            {
                return isBoundary
                    ? $"第{row.Row}行 id={row.Id} 是起止id本身，前缀跟区间内主流id前缀「{majorIdPrefix}」不完全一致，大概率是预留占位行，非混入风险，仅供参考"
                    : $"第{row.Row}行 id={row.Id} 前缀跟区间内主流id前缀「{majorIdPrefix}」不一致，且这个表没有#开头的说明列可核对";
            }

            if (SigMatches(row.Sig, majorSig))
                continue;

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
