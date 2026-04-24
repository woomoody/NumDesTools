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
        string LookupValue  // 源行在该字段的值
    );

    // ═══════════════════════════════════════════════════════════════════════
    // 公共入口
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 对话框输入模式：弹出 InputBox 让用户填「源活动ID，addValue」。
    /// 例如输入 "94005,1" 表示以 ID=94005 为模板，新ID = 94005 + 1 = 94006。
    /// </summary>
    public static void RunDialog()
    {
        var input = Microsoft.VisualBasic.Interaction.InputBox(
            "请输入：源活动ID , addValue（ID增量）\n" +
            "例：94005,1  — 以94005为模板克隆一期，ID+1\n" +
            "例：94005,10 — 以94005为模板克隆一期，ID+10",
            "活动数据克隆",
            ","
        );
        if (string.IsNullOrWhiteSpace(input)) return;

        var parts = input.Split(',');
        if (parts.Length < 2
            || !long.TryParse(parts[0].Trim(), out var sourceId)
            || !int.TryParse(parts[1].Trim(), out var addValue)
            || addValue == 0)
        {
            MessageBox.Show("输入格式有误，请填「源活动ID,addValue」，addValue 不能为0。");
            return;
        }

        Run(sourceId, addValue, fullClone: true);
    }

    /// <summary>
    /// 直接调用：已知源ID与增量，全量克隆（所有ID字段自增）。
    /// </summary>
    public static void Run(long sourceActivityId, int addValue, bool fullClone = true)
    {
        NumDesAddIn.App.StatusBar = "活动克隆：读取规则...";

        var rules = LoadRules();
        if (rules == null) return;

        var excelPath = NumDesAddIn.App.ActiveWorkbook.Path;
        var report = new StringBuilder();
        report.AppendLine("═══════════════ 活动数据克隆报告 ═══════════════");
        report.AppendLine($"时间    ：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        report.AppendLine($"源活动ID：{sourceActivityId}");
        report.AppendLine($"ID增量  ：{addValue}");
        report.AppendLine();

        // 1. 从 ActivityClientData 中查源行，确定活动 type
        var (activityType, activityDataId) = LookupActivityType(
            excelPath, sourceActivityId, report);
        if (activityType < 0)
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }

        report.AppendLine($"活动类型：type={activityType}，activityID={activityDataId}");
        report.AppendLine();

        // 2. 构建需要克隆的目标表列表
        var targets = BuildCloneTargets(activityType, activityDataId, sourceActivityId, rules, excelPath, report);
        if (targets.Count == 0)
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }

        // 3. 依次执行克隆写入
        var errorCount = 0;
        foreach (var target in targets)
        {
            NumDesAddIn.App.StatusBar = $"活动克隆：写入 {target.ExcelName}...";
            var err = CloneTableRow(excelPath, target, addValue, fullClone, report);
            if (err) errorCount++;
        }

        report.AppendLine();
        report.AppendLine($"═════ 完成：{targets.Count} 张表，{errorCount} 个错误 ══════");

        NumDesAddIn.App.StatusBar = errorCount == 0 ? "活动克隆完成" : $"活动克隆完成（{errorCount}个错误）";
        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(report.ToString());
    }

    // ═══════════════════════════════════════════════════════════════════════
    // 内部逻辑
    // ═══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// 从 ActivityClientData.xlsx 读取源活动行，返回 (type, activityID)。
    /// 找不到返回 (-1, -1)。
    /// </summary>
    private static (int type, long activityDataId) LookupActivityType(
        string excelPath, long sourceId, StringBuilder report)
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

        var idCol    = FindHeaderCol(sheet, "id");
        var typeCol  = FindHeaderCol(sheet, "type");
        var dataCol  = FindHeaderCol(sheet, "activityID");

        if (idCol < 0 || typeCol < 0 || dataCol < 0)
        {
            report.AppendLine($"❌ ActivityClientData.xlsx 缺少必要列（id/type/activityID）");
            return (-1, -1);
        }

        for (int row = 3; row <= sheet.Dimension.End.Row; row++)
        {
            var cellVal = sheet.Cells[row, idCol].Value?.ToString();
            if (!long.TryParse(cellVal, out var rowId) || rowId != sourceId) continue;

            var typeStr = sheet.Cells[row, typeCol].Value?.ToString();
            var dataStr = sheet.Cells[row, dataCol].Value?.ToString();
            if (!int.TryParse(typeStr, out var t) || !long.TryParse(dataStr, out var d))
            {
                report.AppendLine($"❌ 源行 id={sourceId} 的 type/activityID 解析失败");
                return (-1, -1);
            }
            return (t, d);
        }

        report.AppendLine($"❌ ActivityClientData.xlsx 中找不到 id={sourceId} 的行");
        return (-1, -1);
    }

    /// <summary>
    /// 根据 activityType 从规则文件中解析出所有需要克隆的目标表。
    /// 固定包含 ActivityClientData（主表）和 ActivityServerData（服务端表）。
    /// </summary>
    private static List<CloneTarget> BuildCloneTargets(
        int activityType, long activityDataId, long sourceClientId,
        RulesRoot rules, string excelPath, StringBuilder report)
    {
        var targets = new List<CloneTarget>
        {
            // 主表：用活动主ID定位
            new("ActivityClientData.xlsx", "id", sourceClientId.ToString()),
        };

        var typeKey = activityType.ToString();

        // 优先查 typeSubTableRules
        if (rules.TypeSubTableRules.TryGetValue(typeKey, out var sub))
        {
            var excelName = ResolveExcelName(sub.Table, rules);
            if (excelName != null)
                targets.Add(new(excelName, sub.LookupField, activityDataId.ToString()));
            else
                report.AppendLine($"⚠ typeSubTableRules[{typeKey}] 表名 {sub.Table} 未在 tables[] 找到对应 excelFile，跳过");
        }
        // 回退 typeTableMap
        else if (rules.TypeTableMap.TryGetValue(typeKey, out var luaKey))
        {
            var excelName = ResolveExcelNameByLuaKey(luaKey, rules);
            if (excelName != null)
                targets.Add(new(excelName, "activityID", activityDataId.ToString()));
            else
                report.AppendLine($"⚠ typeTableMap[{typeKey}] 的 luaKey={luaKey} 未在 tables[] 找到对应 excelFile，跳过");
        }
        else
        {
            report.AppendLine($"⚠ ActivityTableRules.json 中没有 type={activityType} 的子表规则，仅克隆主表和服务端表");
        }

        // 追加 typeMultiSubTableRules 中的多张附属表
        if (rules.TypeMultiSubTableRules.TryGetValue(typeKey, out var multiTables))
        {
            foreach (var excelName in multiTables)
            {
                // 自动探测该表用 activityID 还是 id 作为关联字段
                var detectedField = DetectLookupField(excelPath, excelName, activityDataId.ToString());
                if (detectedField == null)
                {
                    report.AppendLine($"⚠ {excelName}  未找到包含 activityID={activityDataId} 的关联列，跳过");
                    continue;
                }
                targets.Add(new(excelName, detectedField, activityDataId.ToString()));
            }
        }

        report.AppendLine($"克隆目标表（共 {targets.Count} 张）：");
        foreach (var t in targets)
            report.AppendLine($"  • {t.ExcelName}  [查找字段={t.LookupField}  源值={t.LookupValue}]");
        report.AppendLine();

        return targets;
    }

    /// <summary>
    /// 自动探测该表中用于关联 activityID 的字段名：
    /// 优先找名为 "activityID" 的列，其次找 "id" 列，
    /// 并验证该列中确实存在 lookupValue 的数据行。
    /// 找不到返回 null。
    /// </summary>
    private static string DetectLookupField(string excelDir, string excelName, string lookupValue)
    {
        var path = ResolveFilePath(excelDir, excelName);
        if (!File.Exists(path)) return null;

        try
        {
            using var pkg = new ExcelPackage(new FileInfo(path));
            var sheet = pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0];
            if (sheet?.Dimension == null) return null;

            foreach (var candidate in new[] { "activityID", "id" })
            {
                var col = FindHeaderCol(sheet, candidate);
                if (col < 0) continue;
                if (FindRowByValue(sheet, col, lookupValue) >= 0)
                    return candidate;
            }
        }
        catch { /* 打开失败时跳过 */ }

        return null;
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
    /// 克隆到末尾并对所有数值型ID字段按 addValue 自增。
    /// 返回 true 表示有错误。
    /// </summary>
    private static bool CloneTableRow(
        string excelDir, CloneTarget target, int addValue,
        bool fullClone, StringBuilder report)
    {
        var (rawName, sheetName) = ParseExcelName(target.ExcelName);

        // 路径解析（兼容特殊表路径）
        var path = ResolveFilePath(excelDir, rawName);
        if (!File.Exists(path))
        {
            report.AppendLine($"❌ {rawName}  文件不存在：{path}");
            return true;
        }

        ExcelPackage pkg;
        try { pkg = new ExcelPackage(new FileInfo(path)); }
        catch (Exception ex)
        {
            report.AppendLine($"❌ {rawName}  打开失败：{ex.Message}");
            return true;
        }

        using (pkg)
        {
            var sheet = string.IsNullOrEmpty(sheetName)
                ? (pkg.Workbook.Worksheets["Sheet1"] ?? pkg.Workbook.Worksheets[0])
                : pkg.Workbook.Worksheets[sheetName];

            if (sheet?.Dimension == null)
            {
                report.AppendLine($"❌ {target.ExcelName}  Sheet 为空或不存在");
                return true;
            }

            // 检查是否有公式（与现有逻辑一致，避免破坏公式）
            foreach (var cell in sheet.Cells)
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    report.AppendLine($"⚠ {target.ExcelName}  含公式单元格 {cell.Address}，跳过写入");
                    return true;
                }

            var lookupCol = FindHeaderCol(sheet, target.LookupField);
            if (lookupCol < 0)
            {
                report.AppendLine($"❌ {target.ExcelName}  找不到列「{target.LookupField}」");
                return true;
            }

            // 找源行（从数据起始行 3 开始）
            var srcRow = FindRowByValue(sheet, lookupCol, target.LookupValue);
            if (srcRow < 0)
            {
                report.AppendLine($"❌ {target.ExcelName}  找不到 {target.LookupField}={target.LookupValue} 的行");
                return true;
            }

            var colCount = sheet.Dimension.End.Column;
            var destRow  = sheet.Dimension.End.Row + 1;

            // 插入新行并复制值
            EpPlusRowWriter.CloneRow(sheet, srcRow, destRow, colCount);

            // 按 addValue 自增所有数值型 ID 列
            IncrementIdFields(sheet, srcRow, destRow, colCount, addValue, fullClone);

            pkg.Save();

            var newId = sheet.Cells[destRow, lookupCol].Value;
            report.AppendLine($"✓ {target.ExcelName}  源行={srcRow}  新行={destRow}  新{target.LookupField}={newId}");
        }

        return false;
    }

    /// <summary>
    /// 对新行的数值型字段按 addValue 自增。
    /// fullClone=true：所有数值列自增；
    /// fullClone=false：仅第2列（id）和 lookupField 列自增，其余保留原值。
    /// </summary>
    private static void IncrementIdFields(
        ExcelWorksheet sheet, int srcRow, int destRow,
        int colCount, int addValue, bool fullClone)
    {
        for (int col = 2; col <= colCount; col++)
        {
            var header = sheet.Cells[2, col].Value?.ToString() ?? "";
            var cell   = sheet.Cells[destRow, col];
            var val    = cell.Value;

            if (val == null) continue;
            if (!double.TryParse(val.ToString(), out var num)) continue;
            if (num == 0) continue;

            // 只有明确是"ID类"字段才自增
            if (!IsIdField(header)) continue;

            // 部分复用模式下只自增 id 本身，其他引用保持原值
            if (!fullClone && !IsDirectIdField(header)) continue;

            cell.Value = (long)(num + addValue);
        }
    }

    // 判断字段名是否为 ID 类（需要自增）
    private static bool IsIdField(string header)
    {
        if (string.IsNullOrEmpty(header)) return false;
        var h = header.ToLower();
        return h == "id" || h.EndsWith("id") || h.EndsWith("_id")
            || h == "activityid" || h == "activityID";
    }

    // 仅主键 id 和 activityID 本身，用于部分复用模式
    private static bool IsDirectIdField(string header)
    {
        var h = header.ToLower();
        return h == "id" || h == "activityid";
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
        for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            if (sheet.Cells[2, col].Value?.ToString() == headerName)
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
            if (sheet.Cells[row, col].Value?.ToString() == value)
                return row;
        return -1;
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
