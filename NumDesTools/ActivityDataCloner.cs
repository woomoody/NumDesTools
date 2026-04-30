using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Text;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 活动数据克隆工具。
///
/// 读取 ActivityTableRules.json 中的 typeSubTableRules / typeTableMap，
/// 以指定的源活动ID为模板，自动克隆所有相关表的数据行并按 addValue 自增ID，
/// 一次操作完成跨表写入。
///
/// 使用方式：
///   1. 在活动自动填表模板中，填写「源活动ID」与「目标活动ID」
///   2. 点击「克隆活动」按钮，工具自动按类型查表链路完成写入
///   3. 也可通过对话框直接输入源ID + addValue 运行
/// </summary>
public static class ActivityDataCloner
{
    // ── 规则配置文件路径（与 ActivityConfigTester 共用同一份 JSON）─────────────
    private static string RulesFilePath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesTools", "Config", "ActivityTableRules.json"
        );

    // ── 数据模型（只解析克隆所需字段，其余字段忽略）─────────────────────────
    private class RulesRoot
    {
        [JsonProperty("tables")]
        public List<TableDef> Tables { get; set; } = new();

        [JsonProperty("typeTableMap")]
        public Dictionary<string, string> TypeTableMap { get; set; } = new();

        [JsonProperty("typeSubTableRules")]
        public Dictionary<string, SubTableRule> TypeSubTableRules { get; set; } = new();

        [JsonProperty("typeMultiSubTableRules")]
        public Dictionary<string, List<string>> TypeMultiSubTableRules { get; set; } = new();
    }

    private class TableDef
    {
        [JsonProperty("name")]      public string Name      { get; set; }
        [JsonProperty("excelFile")] public string ExcelFile { get; set; }
        [JsonProperty("luaKey")]    public string LuaKey    { get; set; }
    }

    private class SubTableRule
    {
        [JsonProperty("table")]       public string Table       { get; set; }
        [JsonProperty("lookupField")] public string LookupField { get; set; } = "activityID";
    }

    // ── 克隆描述项：一张表要做一次克隆 ──────────────────────────────────────
    private record CloneTarget(
        string ExcelName,   // 文件名（含 #SheetName 后缀时拆分处理）
        string LookupField, // 用于定位源行的字段名
        string LookupValue, // 源行在该字段的值
        bool IsSuspect = false,  // 疑似匹配（非标准列），需用户确认
        bool IsManual = false    // 来自 manualIdMaps 历史补充，找不到源行时静默跳过
    );

    // 探测结果：Certain=标准列精确/前缀匹配，Suspect=非标准列前缀匹配
    private record DetectResult(string FieldName, bool IsSuspect);

    // ── 历史记录 & 手动映射方案 ───────────────────────────────────────────
    private static string HistoryFilePath =>
        Path.Combine(Path.GetDirectoryName(NumDesAddIn.App.ActiveWorkbook.Path)!, "TablesTools", "AliceConfig", "clone_history.json");

    private class CloneHistory
    {
        // 最近 N 条克隆输入记录（原始输入字符串）
        [JsonProperty("recent")]
        public List<string> Recent { get; set; } = new();

        // 手动映射方案：key=ExcelFile，value=List<ManualIdScheme>
        // 每条方案记录「源前缀→目标前缀」及期号语义，便于下次推算
        [JsonProperty("manualIdMaps")]
        public Dictionary<string, List<ManualIdScheme>> ManualIdMaps { get; set; } = new();

    }

    private class ManualIdScheme
    {
        [JsonProperty("src")]   public string Src   { get; set; } = ""; // 源ID前缀
        [JsonProperty("dst")]   public string Dst   { get; set; } = ""; // 目标ID前缀
        [JsonProperty("label")] public string Label { get; set; } = ""; // 期号语义，如"第1期→第2期"
        [JsonProperty("remark")]public string Remark{ get; set; } = ""; // 用于展示的备注
    }

    private static CloneHistory LoadHistory()
    {
        if (!File.Exists(HistoryFilePath)) return new CloneHistory();
        try { return JsonConvert.DeserializeObject<CloneHistory>(File.ReadAllText(HistoryFilePath)) ?? new CloneHistory(); }
        catch { return new CloneHistory(); }
    }

    private static void SaveHistory(CloneHistory h)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(HistoryFilePath)!);
        File.WriteAllText(HistoryFilePath, JsonConvert.SerializeObject(h, Formatting.Indented), Encoding.UTF8);
    }


    // ═══════════════════════════════════════════════════════════════════════
    // 公共入口
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 对话框输入模式 —— 打开 WPF 克隆窗口，支持多行 ID 对照表、历史记录、绑定表映射。
    /// </summary>
    public static void RunDialog()
        => RunDialogWithPrefill(null);

    internal static void RunDialogWithPrefill(List<UI.CloneIdRow>? prefillRows)
    {
        try { RunDialogWithPrefillInternal(prefillRows); }
        catch (Exception ex) { PluginLog.Write($"[Clone] UNHANDLED: {ex}"); }
    }

    private static void RunDialogWithPrefillInternal(List<UI.CloneIdRow>? prefillRows)
    {
        var excelPath = NumDesAddIn.App.ActiveWorkbook.Path;
        PluginLog.Write($"[Clone] RunDialogWithPrefill excelPath={excelPath}");
        var win = new UI.CloneActivityWindow(prefillRows, excelPath);
        win.ShowDialog();
        if (!win.Confirmed) return;

        var (global, perTable) = win.ParseResult();
        if (global.Count == 0 && perTable.Count == 0) return;

        var sourceIds = global.Select(p => long.Parse(p.src)).ToList();
        var targetIds = global.Select(p => long.Parse(p.dst)).ToList();

        var history = LoadHistory();

        // 将 perTable 映射预存入 history.ManualIdMaps
        foreach (var (src, dst, table) in perTable)
        {
            var rawTable = table.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                ? table : table + ".xlsx";
            if (!history.ManualIdMaps.ContainsKey(rawTable))
                history.ManualIdMaps[rawTable] = new List<ManualIdScheme>();
            var schemes = history.ManualIdMaps[rawTable];
            if (!schemes.Any(s => s.Src == src))
                schemes.Add(new ManualIdScheme { Src = src, Dst = dst });
        }
        SaveHistory(history);

        RunInternal(sourceIds, targetIds, win.ResultRemark, history,
                    win.ResultReplaceArt, win.ResultReplaceSubTable,
                    win.ResultSuspectDecisions,
                    win.SavedTableSelections.Count > 0 ? win.SavedTableSelections : null,
                    win,
                    win.ResultRemarkDst, win.ResultIncrementRemark);
    }

    /// <summary>
    /// 直接调用（单ID兼容入口）。
    /// </summary>
    public static void Run(long sourceActivityId, int addValue, bool fullClone = true,
                           string remarkKeyword = "")
    {
        var src = sourceActivityId;
        var dst = sourceActivityId + addValue;
        RunInternal(new List<long> { src }, new List<long> { dst }, remarkKeyword, LoadHistory());
    }

    /// <summary>
    /// 核心入口：sourceIds 与 targetIds 一一对应，构成替换映射。
    /// </summary>
    public static void Run(List<long> sourceIds, List<long> targetIds,
                           string remarkKeyword = "")
        => RunInternal(sourceIds, targetIds, remarkKeyword, null);

    private static void RunInternal(List<long> sourceIds, List<long> targetIds,
                           string remarkKeyword, CloneHistory? history,
                           bool? presetReplaceArt = null,
                           bool? presetReplaceSubTable = null,
                           List<UI.SuspectDecision>? presetSuspect = null,
                           List<UI.TableSelection>? savedTableSelections = null,
                           UI.CloneActivityWindow? originWindow = null,
                           string remarkDst = "", bool incrementRemark = true)
    {
        PluginLog.Write($"[Clone] RunInternal src={string.Join(",", sourceIds)} dst={string.Join(",", targetIds)}");
        if (sourceIds.Count != targetIds.Count || sourceIds.Count == 0) return;
        history ??= LoadHistory();

        // 建立 oldStr→newStr 替换字典（字符串级别，最长优先避免短前缀误替换）
        var idMap = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int i = 0; i < sourceIds.Count; i++)
            idMap[sourceIds[i].ToString()] = targetIds[i].ToString();

        // 期号：从第一对ID差值推算（若差值>0且<100视为期号步进）；受 incrementRemark 控制
        int phaseStep = 0;
        if (incrementRemark && targetIds.Count > 0)
        {
            var diff = (int)(targetIds[0] - sourceIds[0]);
            if (diff != 0 && Math.Abs(diff) < 100) phaseStep = diff;
        }

        NumDesAddIn.App.StatusBar = "活动克隆：读取规则...";
        var rules = LoadRules();
        PluginLog.Write($"[Clone] rules={rules != null}");
        if (rules == null) return;

        var excelPath = NumDesAddIn.App.ActiveWorkbook.Path;

        // 增量更新前缀索引，供后续扫描使用（仅解析第2列数字前缀，远快于全量 DB 同步）
        NumDesAddIn.App.StatusBar = "活动克隆：更新前缀索引...";
        Advance.IdPrefixIndex.Sync(excelPath);
        // Public.db 保留供手动同步/搜索功能使用，克隆扫描不再依赖它
        var dbPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Public.db");
        var report = new StringBuilder();
        report.AppendLine("═══════════════ 活动数据克隆报告 ═══════════════");
        report.AppendLine($"时间    ：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        report.AppendLine($"替换映射：{string.Join("  ", idMap.Select(kv => $"{kv.Key}→{kv.Value}"))}");
        if (!string.IsNullOrEmpty(remarkKeyword)) report.AppendLine($"备注关键字：{remarkKeyword}" + (string.IsNullOrEmpty(remarkDst) ? "" : $" → {remarkDst}"));
        report.AppendLine($"备注序号自增：{(incrementRemark ? "是" : "否")}");
        report.AppendLine();

        // 1. 查第一个源ID确定活动 type
        int activityType = -1;
        long activityDataId = -1;
        foreach (var sid in sourceIds)
        {
            var (t, d) = LookupActivityType(excelPath, sid, report, remarkKeyword);
            if (t >= 0) { activityType = t; activityDataId = d; break; }
        }
        PluginLog.Write($"[Clone] activityType={activityType} activityDataId={activityDataId}");
        if (activityType < 0)
        {
            ErrorLogCtp.DisposeCtp();
            PluginLog.Write(report.ToString());
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }
        report.AppendLine($"活动类型：type={activityType}，activityID={activityDataId}");
        report.AppendLine();

        // 2. 构建目标表列表
        NumDesAddIn.App.StatusBar = "活动克隆：分析关联表...";

        // 预加载一次 ActivityTypeMap.xlsx，避免每次 BuildCloneTargets 重复打开
        var xlsxMapPath   = ActivityTypeMapLoader.ResolveXlsxPath(excelPath);
        var cachedTypeMap = xlsxMapPath != null
            ? ActivityTypeMapLoader.LoadTypeTables(xlsxMapPath)
            : null;

        var allTargets   = new List<CloneTarget>();
        var seenKeys     = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // ExcelName|LookupValue
        var seenSuspects = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // ExcelName（疑似表只加一次）
        foreach (var sid in sourceIds)
        {
            var (_, curDataId) = LookupActivityType(excelPath, sid, new StringBuilder(), remarkKeyword);
            if (curDataId < 0) curDataId = activityDataId;
            var perTargets = BuildCloneTargets(activityType, curDataId, sid, rules, excelPath, report,
                                               xlsxMapPath, cachedTypeMap, dbPath);
            foreach (var ct in perTargets)
            {
                if (ct.IsSuspect)
                {
                    // 疑似表（未配置的扫描结果）：同一文件名只询问并克隆一次
                    if (seenSuspects.Add(ct.ExcelName)) allTargets.Add(ct);
                }
                else
                {
                    // 正常表：同一文件名+同一LookupValue才去重（允许不同活动ID各克隆一份）
                    if (seenKeys.Add($"{ct.ExcelName}|{ct.LookupValue}")) allTargets.Add(ct);
                }
            }
        }
        // 2b. 补充 manualIdMaps 中有记录但尚未进入 allTargets 的表
        // 这些表使用非标准ID前缀（如 Field地组表 用 73601 而非 763601）。
        // LookupValue 选取规则：
        //   若 scheme.Src 在本次 idMap 的 keys 中（即当次克隆 src 就是 scheme.Src）→ 用 Src（源行是当期数据）
        //   否则 → 用 Dst（上次克隆已写入 Dst 行，Dst 就是本次源行）
        foreach (var (tableName, schemes) in history.ManualIdMaps)
        {
            if (allTargets.Any(t => t.ExcelName.Equals(tableName, StringComparison.OrdinalIgnoreCase)))
                continue; // 已在列表中
            var filePath = Path.Combine(excelPath, tableName);
            if (!File.Exists(filePath)) continue;
            var firstScheme = schemes.FirstOrDefault(s => !string.IsNullOrEmpty(s.Dst) || !string.IsNullOrEmpty(s.Src));
            if (firstScheme == null) continue;
            // 用 scheme.Src/Dst 作为探针（这张表用的是异构ID，不能用本次 sourceId 探测）
            var probeId = firstScheme.Src.Length > 0 ? firstScheme.Src : firstScheme.Dst;
            var detected = !string.IsNullOrEmpty(probeId)
                ? DetectLookupField(excelPath, tableName, probeId) : null;
            var lookupField = detected?.FieldName ?? "__col2";
            foreach (var scheme in schemes)
            {
                // 确定 LookupValue：
                //   scheme.Src 末尾期号 == 本次任意 sourceId 末尾期号 → 用 Src（当期源行）
                //   否则 → 用 Dst（上次克隆写入的行即为本次源行）
                var srcPhase = ExtractPhaseNum(scheme.Src);
                var useSrc = sourceIds.Any(sid => ExtractPhaseNum(sid.ToString()) == srcPhase);
                var lookupValue = useSrc ? scheme.Src : scheme.Dst;
                if (string.IsNullOrEmpty(lookupValue)) continue;
                var key = $"{tableName}|{lookupValue}";
                if (!seenKeys.Add(key)) continue;
                allTargets.Add(new CloneTarget(tableName, lookupField, lookupValue, IsSuspect: false, IsManual: true));
                report.AppendLine($"  + manualIdMaps 补充：{tableName}  [LookupValue={lookupValue}  {(useSrc ? $"Src={scheme.Src}" : $"Dst={scheme.Dst}（由{scheme.Src}映射而来）")}]");
            }
        }

        PluginLog.Write($"[Clone] allTargets={allTargets.Count}");
        if (allTargets.Count == 0)
        {
            ErrorLogCtp.DisposeCtp();
            PluginLog.Write(report.ToString());
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }

        // 3. 疑似表确认：有预设（来自窗口历史）则直接用，否则逐一弹框，结果写回历史
        var suspectTargets = allTargets.Where(t => t.IsSuspect).ToList();
        if (suspectTargets.Count > 0)
        {
            var suspectMap = presetSuspect?.ToDictionary(d => d.TableName, d => d.Include,
                                 StringComparer.OrdinalIgnoreCase)
                             ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            var decisions  = new List<UI.SuspectDecision>();
            var skipped    = new List<CloneTarget>();
            foreach (var st in suspectTargets)
            {
                bool include;
                if (suspectMap.TryGetValue(st.ExcelName, out var preset))
                {
                    include = preset;
                    report.AppendLine($"  {(include ? "✓" : "✗")} 历史预设：{st.ExcelName}");
                }
                else
                {
                    var msg = $"疑似匹配（非标准列「{st.LookupField}」含 {st.LookupValue}* 数据）：\n\n" +
                              $"  {st.ExcelName}\n\n是否将该表纳入克隆？";
                    include = MessageBox.Show(msg, "确认克隆",
                        System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
                        == System.Windows.MessageBoxResult.Yes;
                    report.AppendLine($"  {(include ? "✓" : "✗")} 用户{(include ? "确认" : "跳过")}：{st.ExcelName}");
                }
                decisions.Add(new UI.SuspectDecision { TableName = st.ExcelName, Include = include });
                if (!include) skipped.Add(st);
            }
            foreach (var s in skipped) allTargets.Remove(s);
            // 把本次决策写回历史，下次同类活动不再重复询问
            originWindow?.UpdateHistoryWithSuspectDecisions(decisions);
            report.AppendLine();
        }

        // 4. 目标表选择：弹窗让用户勾选/取消要克隆的表（有历史则恢复上次选择）
        NumDesAddIn.App.StatusBar = "活动克隆：选择目标表...";
        var tableNames = allTargets.Select(t => t.ExcelName).ToList();
        var selWin = new UI.CloneTableSelectionWindow(tableNames, savedTableSelections);
        selWin.ShowDialog();
        PluginLog.Write($"[Clone] selWin.Confirmed={selWin.Confirmed}");
        if (!selWin.Confirmed) return; // 用户取消

        var tableSelResult = selWin.Result;
        var selectedSet    = new HashSet<string>(
            tableSelResult.Where(t => t.Selected).Select(t => t.TableName),
            StringComparer.OrdinalIgnoreCase);
        allTargets = allTargets.Where(t => selectedSet.Contains(t.ExcelName)).ToList();

        // 将选择结果回写到来源窗口的历史记录
        PluginLog.Write($"[Clone] UpdateHistoryWithTableSelections...");
        try { originWindow?.UpdateHistoryWithTableSelections(tableSelResult); }
        catch (Exception ex) { PluginLog.Write($"[Clone] UpdateHistoryWithTableSelections ERROR: {ex.Message}"); }
        PluginLog.Write($"[Clone] UpdateHistoryWithTableSelections done, allTargets={allTargets.Count}");

        report.AppendLine($"表选择：共 {tableNames.Count} 张，选中 {allTargets.Count} 张，跳过 {tableNames.Count - allTargets.Count} 张");
        if (allTargets.Count == 0)
        {
            report.AppendLine("所有表均已跳过，克隆取消。");
            ErrorLogCtp.DisposeCtp();
            PluginLog.Write(report.ToString());
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }

        // 4b. 直接用窗口复选框的值，null=未勾选时用默认值（美术=true，子表=false）
        bool replaceArtFields   = presetReplaceArt     ?? true;
        bool replaceSubTableIds = presetReplaceSubTable ?? false;
        report.AppendLine($"特殊字段：美术资源={( replaceArtFields ? "替换" : "保留")}  子表ID={( replaceSubTableIds ? "替换" : "保留")}");

        // 5. 依次执行克隆写入，StatusBar 显示进度
        var total      = allTargets.Count;
        var errorCount = 0;
        for (int i = 0; i < total; i++)
        {
            var target = allTargets[i];
            NumDesAddIn.App.StatusBar =
                $"活动克隆 [{i + 1}/{total}]：{target.ExcelName}";
            var err = CloneTableRow(excelPath, target, idMap, phaseStep, report, history,
                                    replaceArtFields, replaceSubTableIds, remarkKeyword, remarkDst, originWindow);
            if (err) errorCount++;
        }

        SaveHistory(history);
        report.AppendLine();
        report.AppendLine($"═════ 完成：{total} 张表，{errorCount} 个错误 ══════");

        // 先恢复屏幕刷新，再弹 CTP，否则面板不显示（外层 wrapper 的 ScreenUpdating=false 尚未结束）
        NumDesAddIn.App.ScreenUpdating = true;
        NumDesAddIn.App.StatusBar = errorCount == 0
            ? $"活动克隆完成（{total} 张表）"
            : $"活动克隆完成（{total} 张表，{errorCount} 个错误）";
        PluginLog.Write($"[Clone] DisposeCtp...");
        ErrorLogCtp.DisposeCtp();
        PluginLog.Write(report.ToString());
        // 只有出现错误或警告时才弹 CTP 面板，正常完成静默
        var reportStr = report.ToString();
        if (errorCount > 0 || reportStr.Contains('❌') || reportStr.Contains('⚠'))
        {
            PluginLog.Write($"[Clone] CreateCtpNormal, report len={reportStr.Length}");
            ErrorLogCtp.CreateCtpNormal(reportStr);
            PluginLog.Write($"[Clone] CreateCtpNormal done");
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部逻辑
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 从 ActivityClientData.xlsx 读取源活动行，返回 (type, activityID)。
    /// 匹配规则：id 精确匹配；若有 remarkKeyword 则额外允许「id 前缀匹配 AND 备注含关键字」。
    /// 找不到返回 (-1, -1)。
    /// </summary>
    private static (int type, long activityDataId) LookupActivityType(
        string excelPath, long sourceId, StringBuilder report, string remarkKeyword = "")
    {
        var path = Path.Combine(excelPath, "ActivityClientData.xlsx");
        if (!File.Exists(path))
        {
            report.AppendLine($"❌ 找不到 ActivityClientData.xlsx（路径：{excelPath}）");
            return (-1, -1);
        }

        using var pkg = new ExcelPackage(new FileInfo(path));
        var sheet = pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0];
        if (sheet?.Dimension == null)
        {
            report.AppendLine("❌ ActivityClientData.xlsx Sheet1 为空");
            return (-1, -1);
        }

        var idCol     = FindHeaderCol(sheet, "id");
        var typeCol   = FindHeaderCol(sheet, "type");
        var dataCol   = FindHeaderCol(sheet, "activityID");
        var remarkCol = FindHeaderCol(sheet, "#备注");

        if (idCol < 0 || typeCol < 0 || dataCol < 0)
        {
            report.AppendLine($"❌ ActivityClientData.xlsx 缺少必要列（id/type/activityID）");
            return (-1, -1);
        }

        var sourcePrefix = sourceId.ToString();
        var useRemark    = !string.IsNullOrEmpty(remarkKeyword) && remarkCol > 0;

        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
        {
            var cellVal = CellStr(sheet.Cells[row, idCol].Value);
            if (string.IsNullOrEmpty(cellVal)) continue;

            bool matched;
            if (cellVal == sourcePrefix)
            {
                matched = true;
            }
            else if (useRemark && cellVal.StartsWith(sourcePrefix))
            {
                var remark = CellStr(sheet.Cells[row, remarkCol].Value) ?? string.Empty;
                matched = remark.Contains(remarkKeyword);
            }
            else
            {
                continue;
            }

            if (!matched) continue;

            var typeStr = CellStr(sheet.Cells[row, typeCol].Value);
            var dataStr = CellStr(sheet.Cells[row, dataCol].Value);
            if (!int.TryParse(typeStr, out var t) || !long.TryParse(dataStr, out var d))
            {
                report.AppendLine($"❌ 源行 id={cellVal} 的 type/activityID 解析失败");
                return (-1, -1);
            }
            if (cellVal != sourcePrefix)
                report.AppendLine($"⚠ 通过备注关键字「{remarkKeyword}」匹配到 id={cellVal}（前缀={sourcePrefix}）");
            return (t, d);
        }

        report.AppendLine($"❌ ActivityClientData.xlsx 中找不到 id={sourceId}" +
                          (useRemark ? $"（含备注「{remarkKeyword}」前缀匹配）" : string.Empty) + " 的行");
        return (-1, -1);
    }

    /// <summary>
    /// 根据 activityType 解析出所有需要克隆的目标表。
    /// 数据来源优先级：ActivityTypeMap.xlsx（TableTools） > JSON typeMultiSubTableRules > 全目录扫描。
    /// 固定包含 ActivityClientData（主表）。
    /// </summary>
    private static List<CloneTarget> BuildCloneTargets(
        int activityType, long activityDataId, long sourceClientId,
        RulesRoot rules, string excelPath, StringBuilder report,
        string? xlsxPath = null,
        Dictionary<int, List<ActivityTypeMapLoader.TableEntry>>? cachedTypeMap = null,
        string? dbPath = null)
    {
        var targets = new List<CloneTarget>
        {
            new("ActivityClientData.xlsx", "id", sourceClientId.ToString()),
        };

        var typeKey      = activityType.ToString();
        var activityIdStr = activityDataId.ToString();

        // ── 1. typeSubTableRules（JSON 深度规则，最高优先）─────────────────────
        if (rules.TypeSubTableRules.TryGetValue(typeKey, out var sub))
        {
            var excelName = ResolveExcelName(sub.Table, rules);
            if (excelName != null)
                targets.Add(new(excelName, sub.LookupField, activityIdStr));
            else
                report.AppendLine($"⚠ typeSubTableRules[{typeKey}] 表名 {sub.Table} 未在 tables[] 找到对应 excelFile，跳过");
        }
        // 回退 typeTableMap（JSON）
        else if (rules.TypeTableMap.TryGetValue(typeKey, out var luaKey))
        {
            var excelName = ResolveExcelNameByLuaKey(luaKey, rules);
            if (excelName != null)
                targets.Add(new(excelName, "activityID", activityIdStr));
            else
                report.AppendLine($"⚠ typeTableMap[{typeKey}] 的 luaKey={luaKey} 未在 tables[] 找到对应 excelFile，跳过");
        }

        // ── 2. ActivityTypeMap.xlsx（使用调用方预加载的缓存，避免重复打开文件）────
        // xlsxPath / cachedTypeMap 由 RunInternal 在循环外预加载一次传入
        List<string>? xlsxTables = null;

        if (cachedTypeMap != null && cachedTypeMap.TryGetValue(activityType, out var entries))
        {
            xlsxTables = entries.Select(e => e.ExcelFile).ToList();
            report.AppendLine($"ActivityTypeMap.xlsx：type={activityType} 共 {xlsxTables.Count} 张配置表");
        }

        // 3. 回退到 JSON multiRules（xlsx 找不到时使用）
        var multiTables = xlsxTables;
        if (multiTables == null && rules.TypeMultiSubTableRules.TryGetValue(typeKey, out var jsonMulti))
        {
            multiTables = jsonMulti;
            report.AppendLine($"回退到 JSON typeMultiSubTableRules：type={activityType} 共 {multiTables.Count} 张");
        }

        // ── 4. type=2 LTE → 全目录扫描（不依赖白名单）────────────────────────
        if (activityType == 2)
        {
            NumDesAddIn.App.StatusBar = $"活动克隆：type=2 扫描关联表...";
            ScanIndexForTargets(excelPath, activityIdStr, targets, report);
        }
        else if (multiTables != null && multiTables.Count > 0)
        {
            NumDesAddIn.App.StatusBar = $"活动克隆：匹配关联表（共{multiTables.Count}张）...";
            // 优先使用 xlsx entries（含 lookupField），避免逐文件探测
            var xlsxEntries = (cachedTypeMap != null && cachedTypeMap.TryGetValue(activityType, out var e2)) ? e2 : null;
            AddMultiTargets(excelPath, multiTables, activityIdStr, targets, report, xlsxEntries);
        }
        else if (xlsxPath == null && multiTables == null)
        {
            report.AppendLine($"⚠ 未找到 ActivityTypeMap.xlsx，且 JSON 中无 type={activityType} 的多表规则");
        }

        // ── 5. 全目录扫描找未配置的关联表（仅当 ActivityTypeMap.xlsx 未能覆盖该类型时才扫）
        // xlsxTables != null 说明 ActivityTypeMap.xlsx 已有该 type 的完整配置，无需再扫
        if (activityType != 2 && xlsxPath != null && xlsxTables == null)
        {
            NumDesAddIn.App.StatusBar = $"活动克隆：扫描补充未配置表...";
            var configuredFiles = new HashSet<string>(
                targets.Select(t => t.ExcelName.Contains('#')
                    ? t.ExcelName.Split('#')[0]
                    : t.ExcelName),
                StringComparer.OrdinalIgnoreCase);

            var hints = ScanIndexForMissingTables(excelPath, activityIdStr, configuredFiles, report);

            if (hints.Count > 0)
            {
                report.AppendLine($"── 发现 {hints.Count} 张未在配置中的疑似关联表（待外层确认） ──");
                foreach (var hint in hints)
                {
                    var sampleStr = string.Join("  ", hint.Samples.Select(s =>
                        string.IsNullOrEmpty(s.Remark) ? s.Id : $"{s.Id}({s.Remark})"));
                    report.AppendLine($"  ? {hint.FileName}  字段={hint.MatchedField}  样本={sampleStr}");
                    targets.Add(new(hint.FileName, hint.MatchedField, activityIdStr, IsSuspect: true));
                }
            }
        }

        report.AppendLine($"克隆目标表（共 {targets.Count} 张，其中疑似 {targets.Count(t => t.IsSuspect)} 张待确认）：");
        foreach (var t in targets)
            report.AppendLine($"  {(t.IsSuspect ? "？" : "✓")} {t.ExcelName}  [查找字段={t.LookupField}  源值={t.LookupValue}]");
        report.AppendLine();

        return targets;
    }

    /// <summary>
    /// 将 multiTables 列表中的每张表加入 targets。
    /// 若 xlsxEntries 提供了 lookupField 则直接使用，无需打开文件探测。
    /// </summary>
    private static void AddMultiTargets(
        string excelPath, List<string> multiTables,
        string activityIdStr, List<CloneTarget> targets, StringBuilder report,
        List<ActivityTypeMapLoader.TableEntry>? xlsxEntries = null)
    {
        var seenExcel = new HashSet<string>(
            targets.Select(t => t.ExcelName), StringComparer.OrdinalIgnoreCase);

        // 建立 excelFile → lookupField 快查表（来自 ActivityTypeMap.xlsx）
        var entryMap = xlsxEntries != null
            ? xlsxEntries.ToDictionary(e => e.ExcelFile, e => e.LookupField, StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var excelName in multiTables)
        {
            if (seenExcel.Contains(excelName)) continue;

            // 若 ActivityTypeMap.xlsx 已记录 lookupField，直接使用，跳过文件探测
            if (entryMap.TryGetValue(excelName, out var knownField) && !string.IsNullOrEmpty(knownField))
            {
                targets.Add(new(excelName, knownField, activityIdStr, IsSuspect: false));
                seenExcel.Add(excelName);
                continue;
            }

            // 回退：打开文件探测（仅当 xlsx 中没有 lookupField 信息时）
            var detected = DetectLookupField(excelPath, excelName, activityIdStr);
            if (detected == null)
            {
                report.AppendLine($"⚠ {excelName}  未找到包含 activityID={activityIdStr} 的关联列，跳过");
                continue;
            }
            targets.Add(new(excelName, detected.FieldName, activityIdStr, detected.IsSuspect));
            seenExcel.Add(excelName);
        }
    }

    // 不应作为活动关联字段的列名（内部序号，非活动ID）
    private static readonly HashSet<string> _skipFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "sub_table_id", "subtableid", "sub_id"
    };

    // 标准活动关联列（精确/前缀匹配时视为「确定」）
    private static readonly HashSet<string> _standardFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "id", "activityID", "activityId"
    };

    /// <summary>
    /// 自动探测该表中用于关联活动ID的字段名。
    /// Certain：标准列（id/activityID）精确或前缀匹配。
    /// Suspect：非标准列前缀匹配，需用户确认。
    /// 找不到返回 null。
    /// </summary>
    private static DetectResult DetectLookupField(string excelDir, string excelName, string lookupValue)
    {
        var path = ResolveFilePath(excelDir, excelName);
        if (!File.Exists(path)) return null;

        try
        {
            using var pkg = new ExcelPackage(new FileInfo(path));
            var sheet = pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0];
            if (sheet?.Dimension == null) return null;

            // 规则：第2列（key列）前缀匹配 → 确定；其他列匹配 → 疑似
            // 先扫第2列
            var col2Header = CellStr(sheet.Cells[2, 2].Value) ?? string.Empty;
            PluginLog.Verbose($"[Detect] {excelName} col2Header={col2Header} endRow={sheet.Dimension.End.Row} lookupValue={lookupValue}");
            if (!_skipFields.Contains(col2Header))
            {
                for (int row = 3; row <= sheet.Dimension.End.Row; row++)
                {
                    var v = CellStr(sheet.Cells[row, 2].Value);
                    if (v.Length > 0 && (v == lookupValue || v.StartsWith(lookupValue)))
                    {
                        PluginLog.Verbose($"[Detect] {excelName} col2 HIT row={row} v={v}");
                        // 列名为空时退用第2列实际内容，仍找不到则记录列号位置用"__col2"占位
                        var fieldName = col2Header.Length > 0 ? col2Header : "__col2";
                        return new DetectResult(fieldName, IsSuspect: false);
                    }
                }
                PluginLog.Verbose($"[Detect] {excelName} col2 NO HIT after scanning {sheet.Dimension.End.Row - 2} rows");
            }

            // 第2列未命中，扫其余列作为疑似
            for (int col = 3; col <= sheet.Dimension.End.Column; col++)
            {
                var header = CellStr(sheet.Cells[2, col].Value) ?? string.Empty;
                if (_skipFields.Contains(header)) continue;

                for (int row = 3; row <= sheet.Dimension.End.Row; row++)
                {
                    var cellStr = CellStr(sheet.Cells[row, col].Value);
                    if (cellStr != null && (cellStr == lookupValue || cellStr.StartsWith(lookupValue)))
                        return new DetectResult(header, IsSuspect: true);
                }
            }
            return null;
        }
        catch { return null; }
    }

    // 全量目录扫描已知的主表文件（克隆已由 ActivityClientData 处理，不重复扫描）
    private static readonly HashSet<string> _alwaysSkipScan = new(StringComparer.OrdinalIgnoreCase)
    {
        "ActivityClientData.xlsx", "ActivityServerData.xlsx"
    };

    /// <summary>
    /// type=2 LTE 全量扫描：遍历 excelDir 下所有 .xlsx，检测是否含有以 activityDataId 开头的行，
    /// 确定的直接加入 targets，疑似的也加入但标记 IsSuspect=true。
    /// </summary>
    private static void ScanDirectoryForTargets(
        string excelDir, string activityDataId,
        List<CloneTarget> targets, StringBuilder report)
    {
        var existingNames = new HashSet<string>(
            targets.Select(t => Path.GetFileNameWithoutExtension(t.ExcelName) + ".xlsx"),
            StringComparer.OrdinalIgnoreCase);

        var allXlsx = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
        int certain = 0, suspect = 0, skipped = 0;

        foreach (var filePath in allXlsx.OrderBy(p => p))
        {
            var fileName = Path.GetFileName(filePath);
            if (_alwaysSkipScan.Contains(fileName)) continue;
            if (existingNames.Contains(fileName)) continue;

            var detected = DetectLookupField(excelDir, fileName, activityDataId);
            if (detected == null || detected.IsSuspect) { skipped++; continue; }

            targets.Add(new(fileName, detected.FieldName, activityDataId, IsSuspect: false));
            certain++;
        }

        report.AppendLine($"全量扫描：确定 {certain} 张，跳过 {skipped} 张");
    }

    // ── DB 加速：增量同步 + 快速扫描 ────────────────────────────────────────────

    /// <summary>
    /// 将 excelDir 目录下有更新的 xlsx 增量同步到 Public.db。
    /// 只同步比 DB 里已有记录更新（文件修改时间更新）的文件，保持速度。
    /// </summary>
    private static void SyncExcelDirToDb(string excelDir, string dbPath)
    {
        try
        {
            var files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            new Advance.ExcelDataToDb().SyncDirectory(files, dbPath);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[Clone] SyncExcelDirToDb failed: {ex.Message}");
        }
    }

    /// <summary>
    /// 用 IdPrefixIndex 快速找出含 activityDataId 前缀数据的表（type=2 LTE 全目录扫描）。
    /// </summary>
    private static void ScanIndexForTargets(
        string excelDir, string activityDataId,
        List<CloneTarget> targets, StringBuilder report)
    {
        var existingNames = new HashSet<string>(
            targets.Select(t => Path.GetFileNameWithoutExtension(t.ExcelName) + ".xlsx"),
            StringComparer.OrdinalIgnoreCase);

        var candidates = Advance.IdPrefixIndex.FindFiles(activityDataId);
        int certain = 0, skipped = 0;

        foreach (var baseName in candidates)
        {
            if (_alwaysSkipScan.Contains(baseName)) { skipped++; continue; }
            if (existingNames.Contains(baseName)) { skipped++; continue; }

            var detected = DetectLookupField(excelDir, baseName, activityDataId);
            if (detected == null || detected.IsSuspect) { skipped++; continue; }

            targets.Add(new(baseName, detected.FieldName, activityDataId, IsSuspect: false));
            existingNames.Add(baseName);
            certain++;
        }
        report.AppendLine($"索引筛选：命中 {candidates.Count} 张，跳过 {skipped} 张；确定 {certain} 张");
    }

    /// <summary>
    /// 用 IdPrefixIndex 快速找出未配置的关联表（替代 ScanForMissingTables 的文件逐一打开）。
    /// </summary>
    private static List<ActivityTypeMapLoader.MissingTableHint> ScanIndexForMissingTables(
        string excelDir, string activityDataId,
        HashSet<string> configuredFiles, StringBuilder? report)
    {
        var candidates = Advance.IdPrefixIndex.FindFiles(activityDataId)
            .Where(n => !configuredFiles.Contains(n))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (candidates.Count == 0) return [];

        // 把索引未命中的文件加入 configured，让 ScanForMissingTables 只处理候选集
        var allXlsx = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Select(Path.GetFileName)
            .Where(n => !candidates.Contains(n!));
        var reducedConfigured = new HashSet<string>(configuredFiles, StringComparer.OrdinalIgnoreCase);
        foreach (var n in allXlsx) reducedConfigured.Add(n!);

        return ActivityTypeMapLoader.ScanForMissingTables(excelDir, activityDataId, reducedConfigured, report);
    }

    /// <summary>
    /// 用 Public.db 替代全目录 xlsx 扫描，找出包含 activityDataId 前缀数据的表。
    /// 一条 SQL UNION ALL 查所有表，毫秒级完成。
    /// </summary>
    private static void ScanDbForTargets(
        string dbPath, string excelDir, string activityDataId,
        List<CloneTarget> targets, StringBuilder report)
    {
        var existingNames = new HashSet<string>(
            targets.Select(t => Path.GetFileNameWithoutExtension(t.ExcelName) + ".xlsx"),
            StringComparer.OrdinalIgnoreCase);

        // DB 只做快速筛选：找出任意列含 activityDataId 前缀的文件列表
        // 列名是 Excel 列字母（A/B/C），不是字段名，所以只能判断"有没有"，
        // 真实字段名（row2）由后续 DetectLookupField 从 xlsx 读取
        var candidateFiles = new List<string>(); // 文件完整路径
        int dbSkipped = 0;

        try
        {
            using var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}");
            conn.Open();

            var tableNames = GetDbTableNames(conn);
            var prefix = activityDataId + "%";

            foreach (var tbl in tableNames)
            {
                var fileName = GetFileNameFromDb(conn, tbl);
                if (string.IsNullOrEmpty(fileName)) continue;

                var baseName = Path.GetFileName(fileName);
                if (_alwaysSkipScan.Contains(baseName)) { dbSkipped++; continue; }
                if (existingNames.Contains(baseName)) { dbSkipped++; continue; }

                var cols = GetDbTableColumns(conn, tbl);
                bool hit = false;
                foreach (var col in cols)
                {
                    var sql = $"SELECT 1 FROM [{tbl}] WHERE CAST([{col}] AS TEXT) LIKE @p LIMIT 1";
                    using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@p", prefix);
                    if (cmd.ExecuteScalar() != null) { hit = true; break; }
                }

                if (hit)
                    candidateFiles.Add(fileName);
                else
                    dbSkipped++;
            }
        }
        catch (Exception ex)
        {
            report.AppendLine($"⚠ DB扫描失败，回退到文件扫描：{ex.Message}");
            ScanDirectoryForTargets(excelDir, activityDataId, targets, report);
            return;
        }

        // 对命中文件用 DetectLookupField 读真实 header（row2），只取第2列确定命中
        // DB 查询的候选文件可能因为关联引用列而误命中（如 mapTaskGiftID 恰好等于某活动ID），
        // 只认第2列（主key列）命中为"确定"，其余列命中属于误报，直接跳过
        int certain = 0, skipped = 0;
        foreach (var filePath in candidateFiles)
        {
            var baseName = Path.GetFileName(filePath);
            var detected = DetectLookupField(excelDir, baseName, activityDataId);
            if (detected == null || detected.IsSuspect) { skipped++; continue; }

            targets.Add(new(baseName, detected.FieldName, activityDataId, IsSuspect: false));
            existingNames.Add(baseName);
            certain++;
        }

        report.AppendLine($"DB筛选：命中 {candidateFiles.Count} 张，跳过 {dbSkipped} 张；确定 {certain} 张");
    }

    /// <summary>
    /// 用 Public.db 找未配置的关联表（替代 ScanForMissingTables 的文件逐一打开）。
    /// </summary>
    private static List<ActivityTypeMapLoader.MissingTableHint> ScanDbForMissingTables(
        string dbPath, string excelDir, string activityDataId,
        HashSet<string> configuredFiles, StringBuilder? report)
    {
        // DB 只做快速筛选（列名是列字母，不是字段名），找出命中文件名后交给 ScanForMissingTables 处理
        var candidateFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            using var conn = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}");
            conn.Open();
            var prefix = activityDataId + "%";

            foreach (var tbl in GetDbTableNames(conn))
            {
                var fileName = GetFileNameFromDb(conn, tbl);
                if (string.IsNullOrEmpty(fileName)) continue;

                var baseName = Path.GetFileName(fileName);
                if (configuredFiles.Contains(baseName)) continue;
                if (baseName.StartsWith('#') || baseName.StartsWith('~')) continue;
                if (candidateFiles.Contains(baseName)) continue;

                foreach (var col in GetDbTableColumns(conn, tbl))
                {
                    var sql = $"SELECT 1 FROM [{tbl}] WHERE CAST([{col}] AS TEXT) LIKE @p LIMIT 1";
                    using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@p", prefix);
                    if (cmd.ExecuteScalar() != null) { candidateFiles.Add(baseName); break; }
                }
            }
        }
        catch (Exception ex)
        {
            report?.AppendLine($"⚠ DB筛选失败，回退到文件扫描：{ex.Message}");
            return ActivityTypeMapLoader.ScanForMissingTables(excelDir, activityDataId, configuredFiles, report);
        }

        // 只对 DB 筛出的候选文件做 xlsx 探测，而不是全目录
        var reducedConfigured = new HashSet<string>(configuredFiles, StringComparer.OrdinalIgnoreCase);
        // 把 DB 未命中的文件也加入 configured，让 ScanForMissingTables 只处理候选文件
        var allXlsx = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly)
            .Select(Path.GetFileName)
            .Where(n => !candidateFiles.Contains(n!));
        foreach (var n in allXlsx) reducedConfigured.Add(n!);

        return ActivityTypeMapLoader.ScanForMissingTables(excelDir, activityDataId, reducedConfigured, report);
    }

    // DB 工具辅助
    private static readonly HashSet<string> _skipProbeFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "sub_table_id", "subtableid", "sub_id"
    };

    private static List<string> GetDbTableNames(Microsoft.Data.Sqlite.SqliteConnection conn)
    {
        var names = new List<string>();
        using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE '\\_%' ESCAPE '\\'", conn);
        using var rdr = cmd.ExecuteReader();
        while (rdr.Read()) names.Add(rdr.GetString(0));
        return names;
    }

    private static string? GetFileNameFromDb(Microsoft.Data.Sqlite.SqliteConnection conn, string tableName)
    {
        using var cmd = new Microsoft.Data.Sqlite.SqliteCommand(
            "SELECT file_full_path FROM _file_metadata WHERE table_name = @t", conn);
        cmd.Parameters.AddWithValue("@t", tableName);
        return cmd.ExecuteScalar()?.ToString();
    }

    private static List<string> GetDbTableColumns(Microsoft.Data.Sqlite.SqliteConnection conn, string tableName)
    {
        var cols = new List<string>();
        using var cmd = new Microsoft.Data.Sqlite.SqliteCommand($"PRAGMA table_info([{tableName}])", conn);
        using var rdr = cmd.ExecuteReader();
        while (rdr.Read()) cols.Add(rdr.GetString(1));
        return cols;
    }

    /// <summary>
    /// 通过 luaKey（如 "Tables.ActivityRichManV2StageGroups"）查对应 excelFile。
    /// </summary>
    private static string ResolveExcelName(string tableName, RulesRoot rules)
    {
        // tableName 可能是 "Tables.Xxx" 也可能直接是 "Xxx"
        var luaKey = tableName.StartsWith("Tables.") ? tableName : "Tables." + tableName;
        return rules.Tables.FirstOrDefault(t => t.LuaKey == luaKey)?.ExcelFile;
    }

    private static string ResolveExcelNameByLuaKey(string luaKey, RulesRoot rules)
        => rules.Tables.FirstOrDefault(t => t.LuaKey == luaKey)?.ExcelFile;

    // ═══════════════════════════════════════════════════════════════════════
    // 核心 EPPlus 克隆写入
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 在指定 Excel 中找到 lookupField=lookupValue 的源行，
    /// 克隆到同一活动块末尾后用 idMap 做字符串替换。
    /// 返回 true 表示有错误。
    /// </summary>
    // 美术资源相关字段名关键词
    private static readonly HashSet<string> _artFieldKeywords = new(StringComparer.OrdinalIgnoreCase)
    {
        "iconid", "spineid", "textureid", "skinid", "imageid", "spriteid",
        "atlasid", "animid", "animationid", "modelid", "resourceid", "resid",
        "assetid", "bgid", "frameid", "effectid", "soundid", "audioid"
    };

    private static bool IsArtField(string header)
    {
        if (string.IsNullOrEmpty(header)) return false;
        var h = header.ToLower();
        return _artFieldKeywords.Contains(h) || _artFieldKeywords.Any(k => h.EndsWith(k));
    }

    private static readonly HashSet<string> _subTableIdFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "sub_table_id", "subtableid", "sub_id",
        "#_sub_table_id", "_sub_table_id"
    };

    private static bool IsSubTableIdField(string header)
        => !string.IsNullOrEmpty(header) && _subTableIdFields.Contains(header);

    /// <summary>
    /// 一次遍历所有目标表，同时检测美术资源字段和子表ID字段，各弹一次询问。
    /// 美术资源默认替换，子表ID默认不替换。找不到对应字段则直接用默认值。
    /// </summary>
    private static (bool replaceArt, bool replaceSubTable) ShouldReplaceSpecialFields(
        string excelDir, List<CloneTarget> targets, Dictionary<string, string> idMap)
    {
        if (idMap.Count == 0) return (false, false);
        var knownSorted = idMap.Keys.OrderBy(k => k).ToList();

        string? artSample = null, artFile = null, artHeader = null;
        string? subSample = null, subFile = null, subHeader = null;
        int scanIdx = 0, scanTotal = targets.Count;

        foreach (var target in targets)
        {
            if (artSample != null && subSample != null) break; // 两类都找到了

            scanIdx++;
            var (rawName, sheetName) = ParseExcelName(target.ExcelName);
            NumDesAddIn.App.StatusBar = $"活动克隆：检测特殊字段 [{scanIdx}/{scanTotal}] {rawName}";
            var path = ResolveFilePath(excelDir, rawName);
            if (!File.Exists(path)) continue;

            try
            {
                using var pkg = new ExcelPackage(new FileInfo(path));
                var sheet = string.IsNullOrEmpty(sheetName)
                    ? (pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0])
                    : pkg.Workbook.Worksheets[sheetName];
                if (sheet?.Dimension == null) continue;

                int scanEnd = Math.Min(sheet.Dimension.End.Row, 52);
                for (int col = 2; col <= sheet.Dimension.End.Column; col++)
                {
                    var header = CellStr(sheet.Cells[2, col].Value) ?? "";

                    bool wantArt = artSample == null && IsIdField(header) && IsArtField(header);
                    bool wantSub = subSample == null && IsSubTableIdField(header);
                    if (!wantArt && !wantSub) continue;

                    for (int row = 3; row <= scanEnd; row++)
                    {
                        var v = CellStr(sheet.Cells[row, col].Value);
                        if (v == null) continue;
                        if (!idMap.ContainsKey(v) && !IsKnownPrefix(v, knownSorted)) continue;

                        if (wantArt && artSample == null) { artSample = v; artFile = rawName; artHeader = header; }
                        if (wantSub && subSample == null) { subSample = v; subFile = rawName; subHeader = header; }
                        break;
                    }
                }
            }
            catch { }
        }

        bool replaceArt = artSample != null && MessageBox.Show(
            $"在「{artFile}」中检测到美术资源字段「{artHeader}」包含与映射匹配的ID值（如：{artSample}）。\n\n" +
            "这类字段通常指向美术资源，克隆后可能需要单独指定新资源。\n\n" +
            "是否将美术资源字段也纳入 ID 替换？\n「是（默认）」= 同步替换  「否」= 保留原始值",
            "美术资源字段处理", System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question, System.Windows.MessageBoxResult.Yes)
            == System.Windows.MessageBoxResult.Yes;

        bool replaceSubTable = subSample != null && MessageBox.Show(
            $"在「{subFile}」中检测到子表ID字段「{subHeader}」包含与映射匹配的ID值（如：{subSample}）。\n\n" +
            "子表ID通常随主表一起自增，但也可能是独立编号。\n\n" +
            "是否将子表ID字段也纳入 ID 替换？\n「是」= 同步替换  「否（默认）」= 保留原始值",
            "子表ID字段处理", System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Question, System.Windows.MessageBoxResult.No)
            == System.Windows.MessageBoxResult.Yes;

        return (replaceArt, replaceSubTable);
    }

    private static bool CloneTableRow(
        string excelDir, CloneTarget target,
        Dictionary<string, string> idMap, int phaseStep,
        StringBuilder report, CloneHistory history,
        bool replaceArtFields = false, bool replaceSubTableIds = false,
        string remarkKeyword = "", string remarkDst = "",
        UI.CloneActivityWindow? originWindow = null)
    {
        var (rawName, sheetName) = ParseExcelName(target.ExcelName);
        var path = ResolveFilePath(excelDir, rawName);
        if (!File.Exists(path))
        { report.AppendLine($"❌ {rawName}  文件不存在：{path}"); return true; }

        ExcelPackage pkg;
        try { pkg = new ExcelPackage(new FileInfo(path)); }
        catch (Exception ex)
        { report.AppendLine($"❌ {rawName}  打开失败：{ex.Message}"); return true; }

        using (pkg)
        {
            var sheet = string.IsNullOrEmpty(sheetName)
                ? (pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0])
                : pkg.Workbook.Worksheets[sheetName];

            if (sheet?.Dimension == null)
            { report.AppendLine($"❌ {target.ExcelName}  Sheet 为空或不存在"); return true; }


            var lookupCol = target.LookupField == "__col2"
                ? 2
                : FindHeaderCol(sheet, target.LookupField);
            PluginLog.Write($"[Clone] {rawName} sheet={sheet.Name} lookupField={target.LookupField} lookupCol={lookupCol} lookupValue={target.LookupValue}");
            if (lookupCol < 0)
            { report.AppendLine($"❌ {target.ExcelName}  找不到列「{target.LookupField}」"); return true; }

            var remarkCol = FindHeaderCol(sheet, "#备注");
            var colCount  = sheet.Dimension.End.Column;

            // ── 构建该表的有效映射（全局 idMap + 已保存手动方案 + 跨sourceId推算）─────────────
            var effectiveMap = new Dictionary<string, string>(idMap, StringComparer.Ordinal);
            if (history.ManualIdMaps.TryGetValue(rawName, out var savedSchemes))
            {
                foreach (var s in savedSchemes)
                {
                    effectiveMap.TryAdd(s.Src, s.Dst);
                    // 如果 scheme.Dst 就是本次克隆的 LookupValue（即上期已写入的前缀），
                    // 用 phaseStep 推算出当期→下期的映射并加入 effectiveMap
                    if (phaseStep != 0
                        && target.LookupValue.StartsWith(s.Dst, StringComparison.Ordinal)
                        && long.TryParse(s.Dst, out var dstNum))
                    {
                        var nextDst = (dstNum + phaseStep).ToString();
                        effectiveMap.TryAdd(s.Dst, nextDst);
                    }
                }

                // 对 idMap 里其他 oldId→newId 对，尝试用已有 scheme 推算异构前缀
                // 规则：找 scheme.Src 与某个 idMap key 的最长公共后缀，
                //       用同样后缀在其他 idMap key 上拼出新的异构前缀
                foreach (var scheme in savedSchemes)
                {
                    foreach (var (baseOld, baseNew) in idMap)
                    {
                        // 找 scheme.Src 和 baseOld 的公共后缀长度
                        int suffixLen = 0;
                        for (int i = 1; i <= Math.Min(scheme.Src.Length, baseOld.Length); i++)
                        {
                            if (scheme.Src[^i] == baseOld[^i]) suffixLen = i;
                            else break;
                        }
                        if (suffixLen < 3) continue; // 公共后缀太短，不可靠
                        // 对 idMap 里每个不同的 oldId，用相同后缀推算对应异构前缀
                        foreach (var (otherOld, otherNew) in idMap)
                        {
                            if (otherOld == baseOld) continue;
                            if (!otherOld.EndsWith(baseOld[^suffixLen..])) continue;
                            var newSrc = otherOld[..^suffixLen] + scheme.Src[^suffixLen..];
                            // scheme.Dst 和 baseNew 的对应后缀
                            int dstSuffixLen = 0;
                            for (int i = 1; i <= Math.Min(scheme.Dst.Length, baseNew.Length); i++)
                            {
                                if (scheme.Dst[^i] == baseNew[^i]) dstSuffixLen = i;
                                else break;
                            }
                            if (dstSuffixLen < 3) continue;
                            if (!otherNew.EndsWith(baseNew[^dstSuffixLen..])) continue;
                            var newDst = otherNew[..^dstSuffixLen] + scheme.Dst[^dstSuffixLen..];
                            effectiveMap.TryAdd(newSrc, newDst);
                        }
                    }
                }
            }

            // ── 第1步：找到 lookupField 能匹配的源行（精确或前缀）──────────────
            var srcRows     = FindRowsByValueOrPrefix(sheet, lookupCol, target.LookupValue);
            var alienSrcKeys = new HashSet<string>(StringComparer.Ordinal); // 真正有行的 alien 前缀

            // ── 第2步：找 alien ID ───────────────────────────────────────────────
            // 场景A：lookupField 是非ID列（如 activityID），srcRows 已匹配，但第2列（真实ID列）
            //        的值不在 effectiveMap 范围内 → 需要用户补充第2列的映射
            // 场景B：lookupField 是ID列，但有些行的备注含活动关键字、lookupField 值不在 effectiveMap
            //        → 原有逻辑
            var keyCol = 2; // 第2列始终是 key/ID 列
            var alienPrefixes = lookupCol != keyCol
                ? CollectAlienKeyPrefixesFromKeyCol(sheet, keyCol, lookupCol, remarkCol,
                                                     target.LookupValue, effectiveMap, srcRows)
                : CollectAlienKeyPrefixes(
                    sheet, lookupCol, remarkCol, target.LookupValue, effectiveMap, srcRows);


            // ── 第3步：对异构前缀弹窗，让用户提供映射 ──────────────────────────
            // 过滤掉当前活动组已屏蔽的前缀
            var blockedPrefixes = originWindow?.CurrentBlockedAlienPrefixes ?? new();
            blockedPrefixes.TryGetValue(rawName, out var blockedForTable);
            var visibleAliens = alienPrefixes
                .Where(p => blockedForTable == null ||
                            !blockedForTable.Any(b => p.SampleKey.StartsWith(b, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            if (visibleAliens.Count > 0)
            {
                var phaseLabel = phaseStep != 0
                    ? $"第{ExtractPhaseNum(target.LookupValue)}期 → 第{ExtractPhaseNum(target.LookupValue) + phaseStep}期"
                    : $"{target.LookupValue}→?";

                var prefill = visibleAliens.Select(p => new UI.CloneIdRow
                {
                    SourceId   = p.SampleKey,
                    TargetId   = "",
                    Remark     = p.SampleRemark,
                    BoundTable = rawName,
                }).ToList();

                var win = new UI.CloneActivityWindow(prefill, excelDir);
                win.Title = $"异构ID映射 — {rawName}（{phaseLabel}）";
                win.ShowDialog();

                if (win.Confirmed && win.ResultBlockAlienTable)
                {
                    // 屏蔽当前弹窗展示的所有前缀（不是全表，只是这批ID）
                    originWindow?.UpdateHistoryWithBlockedAlienPrefixes(rawName, win.ResultBlockAlienPrefixes);
                    report.AppendLine($"  屏蔽异构前缀：{rawName}  [{string.Join(", ", win.ResultBlockAlienPrefixes)}]");
                }
                else if (win.Confirmed)
                {
                    foreach (var r in win.ResultRows)
                    {
                        var src = r.SourceId.Trim();
                        var dst = r.TargetId.Trim();
                        if (string.IsNullOrEmpty(src) || string.IsNullOrEmpty(dst)) continue;

                        effectiveMap.TryAdd(src, dst);

                        // 保存：记录 src前缀→dst前缀 + 期号标签，供下次推算
                        var tbl = rawName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                            ? rawName : rawName + ".xlsx";
                        if (!history.ManualIdMaps.ContainsKey(tbl))
                            history.ManualIdMaps[tbl] = new List<ManualIdScheme>();
                        var schemes = history.ManualIdMaps[tbl];
                        if (!schemes.Any(s => s.Src == src))
                            schemes.Add(new ManualIdScheme
                            {
                                Src    = src,
                                Dst    = dst,
                                Label  = phaseLabel,
                                Remark = r.Remark,
                            });

                        // 用异构映射扩充 srcRows：找到以 src 为前缀的行也纳入克隆
                        var alienRows = FindRowsByValueOrPrefix(sheet, lookupCol, src);
                        foreach (var ar in alienRows)
                            if (!srcRows.Contains(ar)) srcRows.Add(ar);

                        // 只记录真正有行的 alien 前缀，供后续 insertAfter 计算使用
                        if (alienRows.Count > 0)
                            alienSrcKeys.Add(src);

                        report.AppendLine($"  异构映射：{tbl}  {src}→{dst}  [{phaseLabel}]");
                    }
                }
            }

            if (srcRows.Count == 0)
            {
                if (target.IsManual)
                {
                    // manualIdMaps 补充的表：本期可能没有这个异构前缀的数据，静默跳过
                    report.AppendLine($"  跳过 {target.ExcelName}  [LookupValue={target.LookupValue} 无匹配行]");
                    return false;
                }
                report.AppendLine($"❌ {target.ExcelName}  找不到可克隆的源行");
                return true;
            }

            // ── 第4步：按源行顺序克隆并替换 ────────────────────────────────────
            // insertAfter = 所有 source 前缀（idMap.Keys + alienKeys）和已写入 clone 前缀（idMap.Values）
            // 的最后一行，使克隆行紧跟在所有源数据及已写克隆之后。
            // 找不到某前缀时返回 -1，不会错误退化到表末尾。
            var insertAfter = FindLastRowWithPrefix(sheet, lookupCol, target.LookupValue);
            foreach (var key in idMap.Keys)
            {
                var r = FindLastRowWithPrefix(sheet, lookupCol, key);
                if (r > 0) insertAfter = Math.Max(insertAfter, r);
            }
            foreach (var val in idMap.Values)
            {
                var r = FindLastRowWithPrefix(sheet, lookupCol, val);
                if (r > 0) insertAfter = Math.Max(insertAfter, r);
            }
            foreach (var alienKey in alienSrcKeys)
            {
                var r = FindLastRowWithPrefix(sheet, lookupCol, alienKey);
                if (r > 0) insertAfter = Math.Max(insertAfter, r);
            }
            // 退化保底：如果所有前缀都找不到（insertAfter<=0），追加到表末尾
            if (insertAfter <= 0) insertAfter = sheet.Dimension.End.Row;

            srcRows.Sort();
            foreach (var srcRow in srcRows)
            {
                var destRow = insertAfter + 1;
                sheet.InsertRow(destRow, 1);
                EpPlusRowWriter.CloneRow(sheet, srcRow, destRow, colCount);
                // 替换：仅对新复制的行做 effectiveMap 替换，不收集 unknownIds
                // forcedIdCol：当 lookup 列无列头（__col2）时强制将该列视为 ID 列参与替换
                var forcedIdCol = target.LookupField == "__col2" ? lookupCol : -1;
                PluginLog.Verbose($"[Clone] ApplyIdMap {rawName} srcRow={srcRow} destRow={destRow} forcedIdCol={forcedIdCol} phaseStep={phaseStep} remarkKeyword={remarkKeyword} remarkDst={remarkDst} effectiveMap=[{string.Join(",", effectiveMap.Select(kv => $"{kv.Key}→{kv.Value}"))}]");
                ApplyIdMap(sheet, srcRow, destRow, colCount, effectiveMap, phaseStep, replaceArtFields, replaceSubTableIds, remarkKeyword, remarkDst, forcedIdCol);
                var newId = sheet.Cells[destRow, lookupCol].Value;
                report.AppendLine($"✓ {target.ExcelName}  源行={srcRow}  插入行={destRow}  新{target.LookupField}={newId}");
                insertAfter = destRow;
            }

            pkg.Save();
        }
        return false;
    }

    // 表示一个「异构前缀」：lookupField key 不含已知映射前缀，但备注与活动相关
    private record AlienPrefix(string SampleKey, string SampleRemark);

    /// <summary>
    /// 扫描表中备注包含 lookupValue 相关关键字、但 lookupField 值不在 effectiveMap 前缀范围内的行，
    /// 提取这些行 lookupField 值的最短公共前缀（不同前缀各取一条代表），供用户确认映射。
    /// </summary>
    private static List<AlienPrefix> CollectAlienKeyPrefixes(
        ExcelWorksheet sheet, int lookupCol, int remarkCol,
        string lookupValue, Dictionary<string, string> effectiveMap,
        List<int> alreadyMatchedRows)
    {
        if (sheet.Dimension == null || remarkCol < 0) return [];

        var alreadySet   = alreadyMatchedRows.ToHashSet();
        var knownSorted  = effectiveMap.Keys.OrderBy(k => k).ToList();
        var knownLenRange = GetKnownLenRange(effectiveMap.Keys);

        // 备注关键字：取 lookupValue 前4位作为模糊匹配依据
        var remarkHint = lookupValue.Length >= 3 ? lookupValue[..Math.Min(4, lookupValue.Length)] : lookupValue;

        var seenPrefixes = new Dictionary<string, AlienPrefix>(StringComparer.Ordinal);

        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
        {
            if (alreadySet.Contains(row)) continue;

            var keyVal = CellStr(sheet.Cells[row, lookupCol].Value);
            if (string.IsNullOrEmpty(keyVal)) continue;

            if (IsKnownPrefix(keyVal, knownSorted)) continue;
            // 位数与已知 key 相差超过 2 → 不是同类活动 ID，跳过
            if (!IsCompatibleLength(keyVal.Length, knownLenRange)) continue;

            var remark = CellStr(sheet.Cells[row, remarkCol].Value) ?? "";
            if (!remark.Contains(remarkHint)) continue;

            var bucket = keyVal.Length >= 5 ? keyVal[..5] : keyVal;
            if (!seenPrefixes.ContainsKey(bucket))
                seenPrefixes[bucket] = new AlienPrefix(keyVal, remark);
        }

        return seenPrefixes.Values.ToList();
    }

    /// <summary>
    /// lookupField 不是第2列时（如 activityID 是 I 列，真实 ID 在第2列 B）：
    /// 扫描 srcRows 已匹配行的第2列，找出不在 effectiveMap 前缀范围内的值，
    /// 按5位前缀分桶，每桶取一条代表行，供用户补充映射。
    /// </summary>
    private static List<AlienPrefix> CollectAlienKeyPrefixesFromKeyCol(
        ExcelWorksheet sheet, int keyCol, int lookupCol, int remarkCol,
        string lookupValue, Dictionary<string, string> effectiveMap,
        List<int> srcRows)
    {
        if (sheet.Dimension == null || srcRows.Count == 0) return [];

        var knownSorted  = effectiveMap.Keys.OrderBy(k => k).ToList();
        var knownLenRange = GetKnownLenRange(effectiveMap.Keys);
        var seenPrefixes = new Dictionary<string, AlienPrefix>(StringComparer.Ordinal);

        foreach (var row in srcRows)
        {
            var keyVal = CellStr(sheet.Cells[row, keyCol].Value);
            if (string.IsNullOrEmpty(keyVal)) continue;
            if (IsKnownPrefix(keyVal, knownSorted)) continue;
            if (!IsCompatibleLength(keyVal.Length, knownLenRange)) continue;

            var remark = remarkCol > 0
                ? CellStr(sheet.Cells[row, remarkCol].Value) ?? ""
                : "";
            var bucket = keyVal.Length >= 5 ? keyVal[..5] : keyVal;
            if (!seenPrefixes.ContainsKey(bucket))
                seenPrefixes[bucket] = new AlienPrefix(keyVal, remark);
        }

        return seenPrefixes.Values.ToList();
    }

    // 返回已知 key 集合的位数范围 (min, max)，用于过滤明显不同类的 ID
    private static (int min, int max) GetKnownLenRange(IEnumerable<string> keys)
    {
        int min = int.MaxValue, max = 0;
        foreach (var k in keys)
        {
            if (k.Length < min) min = k.Length;
            if (k.Length > max) max = k.Length;
        }
        return min == int.MaxValue ? (0, 99) : (min, max);
    }

    // 候选 ID 位数是否与已知 key 位数范围兼容（允许 ±2 的浮动）
    private static bool IsCompatibleLength(int len, (int min, int max) range)
        => len >= range.min - 2 && len <= range.max + 2;

    // 检查 keyVal 是否以 sortedKnown 中某个 key 为前缀（sortedKnown 已排序，可提前终止）
    private static bool IsKnownPrefix(string keyVal, List<string> sortedKnown)
    {
        foreach (var k in sortedKnown)
        {
            if (k.Length > keyVal.Length) continue;
            if (keyVal.StartsWith(k)) return true;
        }
        return false;
    }

    // 从活动ID字符串中提取期号（末尾数字，如 "73601" → 1，"73602" → 2）
    private static int ExtractPhaseNum(string id)
    {
        if (string.IsNullOrEmpty(id)) return 1;
        // 尝试取最后一位数字作为期号
        for (int i = id.Length - 1; i >= 0; i--)
            if (char.IsDigit(id[i]))
                return id[i] - '0';
        return 1;
    }

    /// <summary>
    /// 找到最后一个在 lookupCol 列值以 prefix 开头的行号（同一活动前缀的最末行）。
    /// 找不到时返回 -1。
    /// </summary>
    private static int FindLastRowWithPrefix(ExcelWorksheet sheet, int lookupCol, string prefix)
    {
        for (int row = sheet.Dimension.End.Row; row >= 3; row--)
        {
            var val = CellStr(sheet.Cells[row, lookupCol].Value);
            if (val != null && (val == prefix || val.StartsWith(prefix)))
                return row;
        }
        return -1;
    }

    /// <summary>
    /// 对新行所有字段做 idMap 字符串替换，同时处理期号和子行ID前缀。
    /// 包含 idMap 中 key 的值则替换，否则不变。美术字段受 replaceArtFields 控制。
    /// </summary>
    private static void ApplyIdMap(
        ExcelWorksheet sheet, int srcRow, int destRow, int colCount,
        Dictionary<string, string> idMap, int phaseStep,
        bool replaceArtFields = false, bool replaceSubTableIds = false,
        string remarkKeyword = "", string remarkDst = "",
        int forcedIdCol = -1)  // 无列头的 lookup 列（如 __col2），强制视为 ID 列
    {
        // 预建排序后的 oldId 列表（长的先匹配，防止短前缀误替换）
        var sortedKeys = idMap.Keys.OrderByDescending(k => k.Length).ToList();
        var hasRemarkReplace = !string.IsNullOrEmpty(remarkKeyword) && !string.IsNullOrEmpty(remarkDst);

        for (int col = 2; col <= colCount; col++)
        {
            var header = CellStr(sheet.Cells[2, col].Value) ?? "";
            var cell   = sheet.Cells[destRow, col];
            var val    = cell.Value;
            if (val == null) continue;

            var valStr = val.ToString()!;

            // ── 备注列：做关键字替换（remarkKeyword → remarkDst），不走ID替换逻辑 ──
            if (header == "#备注")
            {
                var updated = valStr;
                var keywordHit = hasRemarkReplace && updated.Contains(remarkKeyword);
                if (keywordHit)
                    updated = updated.Replace(remarkKeyword, remarkDst);
                // 关键字替换命中时 remarkDst 已含目标期号，不再自增；否则按 phaseStep 自增
                if (!keywordHit && phaseStep != 0)
                    updated = ReplacePhaseNumber(updated, phaseStep);
                PluginLog.Verbose($"[Remark] row={destRow} orig={valStr} updated={updated} keywordHit={keywordHit} keyword={remarkKeyword}");
                if (updated != valStr) cell.Value = updated;
                continue;
            }

            // ── 纯数字字段（可能是ID）─────────────────────────────────────────
            if (long.TryParse(valStr, out var numVal) && numVal != 0)
            {
                // 子表ID字段：受 replaceSubTableIds 开关控制
                if (IsSubTableIdField(header) && !replaceSubTableIds) continue;
                // 美术ID字段：受 replaceArtFields 开关控制
                if (IsArtField(header) && !replaceArtFields) continue;

                // 非ID类列名的普通数字字段：只有在 effectiveMap 中能精确或前缀命中时才替换，
                // 否则跳过（避免误改金额、数量等普通数字）
                var isId = IsIdField(header) || IsSubTableIdField(header) || col == forcedIdCol;

                // 1. 精确匹配 idMap → 直接替换
                if (idMap.TryGetValue(valStr, out var mapped))
                { cell.Value = long.Parse(mapped); continue; }

                // 2. 子行前缀替换：若该数字以某个 oldId 为前缀且更长
                var prefixHit = false;
                foreach (var oldKey in sortedKeys)
                {
                    if (valStr.Length > oldKey.Length && valStr.StartsWith(oldKey))
                    {
                        cell.Value = long.Parse(idMap[oldKey] + valStr[oldKey.Length..]);
                        prefixHit = true;
                        break;
                    }
                }
                if (prefixHit) continue;

                // 3. 非ID列名且无法命中 → 不替换
                if (!isId) continue;
                continue;
            }

            // ── 字符串字段：做多次子串替换（ID、期号）─────────────────────────
            var upd = valStr;
            foreach (var oldKey in sortedKeys)
                upd = upd.Replace(oldKey, idMap[oldKey]);

            // 期号替换（用 phaseStep 计算）
            if (phaseStep != 0)
                upd = ReplacePhaseNumber(upd, phaseStep);

            if (upd != valStr) cell.Value = upd;
        }
    }

    /// <summary>针对已插入行的单列回写（手动映射补充替换）。</summary>
    private static void ApplyIdMapSingle(
        ExcelWorksheet sheet, int row, int colCount, string oldId, string newId)
    {
        for (int col = 2; col <= colCount; col++)
        {
            var cell = sheet.Cells[row, col];
            var val  = CellStr(cell.Value);
            if (string.IsNullOrEmpty(val)) continue;
            if (val == oldId) { cell.Value = long.TryParse(newId, out var n) ? (object)n : newId; continue; }
            if (val.StartsWith(oldId) && long.TryParse(val, out _))
            { cell.Value = long.Parse(newId + val[oldId.Length..]); continue; }
            if (val.Contains(oldId)) cell.Value = val.Replace(oldId, newId);
        }
    }

    private static readonly System.Text.RegularExpressions.Regex RxPhaseNum =
        new(@"第(\d+)期", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string ReplacePhaseNumber(string text, int step)
    {
        if (!text.Contains("第") || !text.Contains("期")) return text;
        return RxPhaseNum.Replace(text, m =>
            $"第{int.Parse(m.Groups[1].Value) + step}期");
    }

    private static string CellStr(object? value) => value?.ToString() ?? "";

    private static bool IsIdField(string header)
    {
        if (string.IsNullOrEmpty(header)) return false;
        var h = header.ToLower();
        return h == "id" || h.EndsWith("id") || h.EndsWith("_id")
            || h == "activityid" || h == "activityID";
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 通用 EPPlus 辅助：供其他模块复用
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 找到第2行（表头行）中值等于 headerName 的列号（1-based），找不到返回 -1。
    /// </summary>
    public static int FindHeaderCol(ExcelWorksheet sheet, string headerName)
    {
        if (sheet.Dimension == null) return -1;
        var endCol = sheet.Dimension.End.Column;
        // 优先第2行（标准表头行），找不到再查第1行（部分表把 #备注 放在第1行）
        for (int col = 1; col <= endCol; col++)
            if (CellStr(sheet.Cells[2, col].Value) == headerName)
                return col;
        for (int col = 1; col <= endCol; col++)
            if (CellStr(sheet.Cells[1, col].Value) == headerName)
                return col;
        return -1;
    }

    /// <summary>
    /// 从第3行起按列值查找行号（1-based），找不到返回 -1。
    /// </summary>
    public static int FindRowByValue(ExcelWorksheet sheet, int col, string value)
    {
        if (sheet.Dimension == null) return -1;
        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
            if (CellStr(sheet.Cells[row, col].Value) == value)
                return row;
        return -1;
    }

    /// <summary>
    /// 收集所有精确匹配或以 prefix 开头的行（Item 类表多行场景）。
    /// 精确匹配优先；若无精确匹配则返回所有前缀匹配行。
    /// </summary>
    private static List<int> FindRowsByValueOrPrefix(ExcelWorksheet sheet, int col, string prefix)
    {
        if (sheet.Dimension == null) return [];

        var exact  = new List<int>();
        var byPrefix = new List<int>();

        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
        {
            var cellStr = CellStr(sheet.Cells[row, col].Value);
            if (cellStr.Length == 0) continue;
            if (cellStr == prefix)
                exact.Add(row);
            else if (cellStr.StartsWith(prefix))
                byPrefix.Add(row);
        }

        return exact.Count > 0 ? exact : byPrefix;
    }

    // ── 路径 / 名称解析 ──────────────────────────────────────────────────

    private static (string rawName, string sheetName) ParseExcelName(string name)
    {
        if (name.Contains('#'))
        {
            var parts = name.Split('#', 2);
            return (parts[0], parts[1]);
        }
        return (name, string.Empty);
    }

    private static string ResolveFilePath(string baseDir, string fileName)
    {
        var parent = Path.GetDirectoryName(Path.GetDirectoryName(baseDir));
        return fileName switch
        {
            "Localizations.xlsx"  => Path.Combine(parent ?? baseDir, "Excels", "Localizations", fileName),
            "UIConfigs.xlsx"      => Path.Combine(parent ?? baseDir, "Excels", "UIs", fileName),
            "UIItemConfigs.xlsx"  => Path.Combine(parent ?? baseDir, "Excels", "UIs", fileName),
            _                     => Path.Combine(baseDir, fileName)
        };
    }

    // ── 规则加载 ─────────────────────────────────────────────────────────

    private static RulesRoot LoadRules()
    {
        if (!File.Exists(RulesFilePath))
        {
            MessageBox.Show($"找不到规则配置文件：\n{RulesFilePath}\n请先运行一次「验证活动」以生成默认配置。");
            return null;
        }
        try
        {
            return JsonConvert.DeserializeObject<RulesRoot>(
                File.ReadAllText(RulesFilePath, Encoding.UTF8));
        }
        catch (Exception ex)
        {
            MessageBox.Show($"读取规则配置失败：{ex.Message}");
            return null;
        }
    }
}

/// <summary>
/// 通用 EPPlus 行克隆写入工具，供 ActivityDataCloner 及其他模块复用。
/// </summary>
public static class EpPlusRowWriter
{
    /// <summary>
    /// 将 sheet 中 srcRow 行的值复制到 destRow 行（destRow 须已存在或为末尾+1）。
    /// 同时应用统一的字体/对齐风格。
    /// </summary>
    public static void CloneRow(
        ExcelWorksheet sheet, int srcRow, int destRow, int colCount)
    {
        var src  = sheet.Cells[srcRow,  1, srcRow,  colCount];
        var dest = sheet.Cells[destRow, 1, destRow, colCount];

        dest.Value = src.Value;
        dest.Style.Font.Name = "微软雅黑";
        dest.Style.Font.Size = 10;
        dest.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
    }

    /// <summary>
    /// 在 sheet 末尾追加一行克隆数据，返回新行行号。
    /// </summary>
    public static int AppendClonedRow(ExcelWorksheet sheet, int srcRow, int colCount)
    {
        var destRow = (sheet.Dimension?.End.Row ?? 2) + 1;
        CloneRow(sheet, srcRow, destRow, colCount);
        return destRow;
    }

    /// <summary>
    /// 在指定 Excel 文件中查找源行并克隆到末尾，直接保存。
    /// lookupCol 为查找列（1-based），lookupValue 为目标值。
    /// 返回新行号，失败返回 -1。
    /// </summary>
    public static int CloneRowToEnd(
        string filePath, string sheetName,
        int lookupCol, string lookupValue,
        int colCount = 0)
    {
        if (!File.Exists(filePath)) return -1;

        using var pkg  = new ExcelPackage(new FileInfo(filePath));
        var sheet = string.IsNullOrEmpty(sheetName)
            ? (pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0])
            : pkg.Workbook.Worksheets[sheetName];

        if (sheet?.Dimension == null) return -1;

        var srcRow = ActivityDataCloner.FindRowByValue(sheet, lookupCol, lookupValue);
        if (srcRow < 0) return -1;

        var cols    = colCount > 0 ? colCount : sheet.Dimension.End.Column;
        var destRow = AppendClonedRow(sheet, srcRow, cols);
        pkg.Save();
        return destRow;
    }
}
