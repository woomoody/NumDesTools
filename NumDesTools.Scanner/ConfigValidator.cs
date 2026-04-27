using Newtonsoft.Json;

namespace NumDesTools.Scanner;

/// <summary>
/// 统一入口：加载规则 → 遍历所有表 → 跑 L1+L2 → 打印报告 → 可选输出 MD 文件。
/// 支持 --activity ID 只验证单个活动相关表。
/// 支持 --md [path]  同时输出 Markdown 报告（默认路径 ConfigDir）。
/// </summary>
public static class ConfigValidator
{
    private const string TablesDir = @"C:\M1Work\Public\Excels\Tables";
    private const string RulesPath = @"C:\Users\cent\Documents\NumDesTools\Config\ActivityTableRules.json";
    private const string ReportDir = @"C:\Users\cent\Documents\NumDesTools\Config";

    public static int Run(string[] args)
    {
        bool onlyErrors = args.Contains("--errors-only");
        string? activityFilter = null;
        string? mdPath = null;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--activity" && i + 1 < args.Length)
                activityFilter = args[i + 1];
            if (args[i] == "--md")
                // --md 后面可以跟自定义路径，也可以不跟（用默认）
                mdPath = (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                    ? args[i + 1]
                    : Path.Combine(ReportDir, "validate_latest.md");
        }

        Console.OutputEncoding = System.Text.Encoding.UTF8;

        if (!File.Exists(RulesPath))
        { Console.WriteLine($"[ERROR] 规则文件不存在: {RulesPath}"); return 1; }

        ActivityTableRules rules;
        try
        {
            rules = JsonConvert.DeserializeObject<ActivityTableRules>(
                        File.ReadAllText(RulesPath)) ?? new();
        }
        catch (Exception ex)
        { Console.WriteLine($"[ERROR] 规则文件解析失败: {ex.Message}"); return 1; }

        // 交互式选择活动（--activity 未指定时询问）
        if (args.Contains("--pick"))
            activityFilter = PickActivity(TablesDir);

        var tableDefs = rules.Tables
            .Where(t => !string.IsNullOrEmpty(t.ExcelFile))
            .ToList();

        if (activityFilter != null)
            tableDefs = FilterByActivity(activityFilter, tableDefs, rules, TablesDir);

        PrintBanner(activityFilter, mdPath);

        var report = new ValidationReport();

        foreach (var def in tableDefs)
        {
            var excelPath = Path.Combine(TablesDir, def.ExcelFile);
            Console.Write($"  校验 {def.ExcelFile,-50}");

            var r1 = L1_TableValidator.Validate(excelPath, def);
            var r2 = L2_RefValidator.Validate(TablesDir, def, rules);

            var merged = new TableValidationReport { ExcelFile = def.ExcelFile };
            merged.Issues.AddRange(r1.Issues);
            merged.Issues.AddRange(r2.Issues);
            report.Tables.Add(merged);

            int e = merged.Issues.Count(i => i.Level == Severity.Error);
            int w = merged.Issues.Count(i => i.Level == Severity.Warning);
            Console.WriteLine(e > 0 ? $"  ❌ {e} 错误  {w} 警告" : w > 0 ? $"  ⚠  {w} 警告" : "  ✓");
        }

        PrintConsoleReport(report, onlyErrors);

        if (mdPath != null)
        {
            WriteMdReport(report, mdPath, activityFilter, onlyErrors);
            Console.WriteLine($"  📄 报告已保存: {mdPath}");
        }

        return report.ErrorCount > 0 ? 2 : 0;
    }

    // ── 控制台输出 ────────────────────────────────────────────────────────────
    private static void PrintBanner(string? filter, string? mdPath)
    {
        Console.WriteLine();
        Console.WriteLine("╔═══════════════════════════════════════════════════════╗");
        Console.WriteLine("║   NumDesTools — 配置自检  L1(表格) + L2(引用)         ║");
        Console.WriteLine("╚═══════════════════════════════════════════════════════╝");
        if (filter != null) Console.WriteLine($"  过滤: activityID = {filter}");
        Console.WriteLine($"  表目录: {TablesDir}");
        if (mdPath != null) Console.WriteLine($"  报告将输出至: {mdPath}");
        Console.WriteLine();
    }

    private static void PrintConsoleReport(ValidationReport report, bool onlyErrors)
    {
        Console.WriteLine();
        Console.WriteLine(new string('─', 90));

        var issues = onlyErrors
            ? report.AllIssues.Where(i => i.Level == Severity.Error)
            : report.AllIssues;

        foreach (var grp in issues.GroupBy(i => i.ExcelFile))
        {
            Console.WriteLine($"\n  【{grp.Key}】");
            foreach (var iss in grp.OrderBy(i => i.Row))
                Console.WriteLine("    " + iss);
        }

        Console.WriteLine();
        Console.WriteLine(new string('─', 90));
        Console.ForegroundColor = report.ErrorCount > 0 ? ConsoleColor.Red : ConsoleColor.Green;
        Console.WriteLine($"  结果: {report.ErrorCount} 个错误  {report.WarningCount} 个警告" +
                          $"  (扫描 {report.Tables.Count} 张表，{report.RunAt:HH:mm:ss})");
        Console.ResetColor();
        Console.WriteLine();
    }

    // ── Markdown 报告 ─────────────────────────────────────────────────────────
    private static void WriteMdReport(
        ValidationReport report, string path, string? filter, bool onlyErrors)
    {
        var sb = new System.Text.StringBuilder();

        sb.AppendLine("# 配置自检报告");
        sb.AppendLine();
        sb.AppendLine($"| 项目 | 值 |");
        sb.AppendLine($"|------|-----|");
        sb.AppendLine($"| 生成时间 | {report.RunAt:yyyy-MM-dd HH:mm:ss} |");
        sb.AppendLine($"| 扫描表数 | {report.Tables.Count} |");
        sb.AppendLine($"| 错误数 | **{report.ErrorCount}** |");
        sb.AppendLine($"| 警告数 | {report.WarningCount} |");
        if (filter != null)
            sb.AppendLine($"| 活动过滤 | activityID = {filter} |");
        sb.AppendLine();

        // 扫描概览表
        sb.AppendLine("## 扫描概览");
        sb.AppendLine();
        sb.AppendLine("| 表文件 | 错误 | 警告 | 状态 |");
        sb.AppendLine("|--------|------|------|------|");
        foreach (var t in report.Tables)
        {
            int e = t.Issues.Count(i => i.Level == Severity.Error);
            int w = t.Issues.Count(i => i.Level == Severity.Warning);
            string status = e > 0 ? "❌ 有错误" : w > 0 ? "⚠ 有警告" : "✅ 通过";
            sb.AppendLine($"| {t.ExcelFile} | {e} | {w} | {status} |");
        }
        sb.AppendLine();

        // 问题明细
        var allIssues = onlyErrors
            ? report.AllIssues.Where(i => i.Level == Severity.Error)
            : report.AllIssues;

        var grouped = allIssues.GroupBy(i => i.ExcelFile).ToList();
        if (grouped.Count == 0)
        {
            sb.AppendLine("## 问题明细");
            sb.AppendLine();
            sb.AppendLine("> ✅ 未发现任何问题。");
        }
        else
        {
            sb.AppendLine("## 问题明细");
            sb.AppendLine();
            foreach (var grp in grouped)
            {
                sb.AppendLine($"### {grp.Key}");
                sb.AppendLine();
                sb.AppendLine("| 行号 | 层级 | 严重程度 | 字段 | 描述 |");
                sb.AppendLine("|------|------|----------|------|------|");
                foreach (var iss in grp.OrderBy(i => i.Row))
                {
                    string lvl = iss.Level == Severity.Error ? "🔴 Error" : "🟡 Warning";
                    sb.AppendLine($"| {iss.Row} | {iss.Layer} | {lvl} | `{iss.Field}` | {iss.Message} |");
                }
                sb.AppendLine();
            }
        }

        // 规则说明
        sb.AppendLine("## 校验规则说明");
        sb.AppendLine();
        sb.AppendLine("| 层 | 规则 | 说明 |");
        sb.AppendLine("|----|------|------|");
        sb.AppendLine("| L1 | 必填为空 | `required=true` 字段无内容 |");
        sb.AppendLine("| L1 | 主键重复 | 同表内 `id` 出现两次 |");
        sb.AppendLine("| L1 | 数值非法 | 类型声明 int/number 的列含非数字 |");
        sb.AppendLine("| L1 | ID ≤ 0 | id / activityID 值为 0 或负数 |");
        sb.AppendLine("| L2 | 引用不存在 | refTable 字段指向的记录不存在 |");
        sb.AppendLine("| L2 | 子表缺失 | type→子表中找不到对应 activityID |");
        sb.AppendLine();
        sb.AppendLine("---");
        sb.AppendLine($"*由 NumDesTools.Scanner ConfigValidator 自动生成*");

        File.WriteAllText(path, sb.ToString(), System.Text.Encoding.UTF8);
    }

    // ── 交互式选择活动 ────────────────────────────────────────────────────────
    private static string? PickActivity(string tablesDir)
    {
        var mainPath = Path.Combine(tablesDir, "ActivityClientData.xlsx");
        if (!File.Exists(mainPath))
        {
            Console.WriteLine("[WARN] ActivityClientData.xlsx 不存在，无法列出活动");
            return null;
        }

        var (_, rows) = ExcelReader.Read(mainPath);

        // 取最近 30 条（按 id 降序）
        // 注意：#备注 列在 ExcelReader 中被跳过，改用 name 列作备注
        var entries = rows
            .Select(r => (
                Id:      r.Data.GetValueOrDefault("id", ""),
                ActId:   r.Data.GetValueOrDefault("activityID", ""),
                Type:    r.Data.GetValueOrDefault("type", ""),
                Remark:  r.Data.GetValueOrDefault("name", "")
            ))
            .Where(e => !string.IsNullOrEmpty(e.Id) && !string.IsNullOrEmpty(e.ActId))
            .OrderByDescending(e => long.TryParse(e.Id, out long v) ? v : 0)
            .Take(30)
            .ToList();

        Console.WriteLine();
        Console.WriteLine("  最近 30 条活动（按 id 降序）：");
        Console.WriteLine($"  {"序号",-4} {"id",-10} {"activityID",-12} {"type",-6} 备注");
        Console.WriteLine("  " + new string('─', 60));
        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            Console.WriteLine($"  {i + 1,-4} {e.Id,-10} {e.ActId,-12} {e.Type,-6} {e.Remark}");
        }
        Console.WriteLine();
        Console.Write("  输入序号或直接输入 activityID（回车跳过=全量校验）: ");

        var input = Console.ReadLine()?.Trim() ?? "";
        if (string.IsNullOrEmpty(input)) return null;

        // 输入的是序号
        if (int.TryParse(input, out int idx) && idx >= 1 && idx <= entries.Count)
            return entries[idx - 1].ActId;

        // 输入的是 activityID 字符串
        return input;
    }

    // ── 按 activityID 过滤相关表 ─────────────────────────────────────────────
    private static List<TableDef> FilterByActivity(
        string actId, List<TableDef> all, ActivityTableRules rules, string tablesDir)
    {
        // 先读主表找 type
        var mainPath = Path.Combine(tablesDir, "ActivityClientData.xlsx");
        string? typeStr = null;
        if (File.Exists(mainPath))
        {
            var (_, rows) = ExcelReader.Read(mainPath);
            foreach (var (_, data) in rows)
            {
                if (data.TryGetValue("activityID", out var aid) && aid == actId)
                { data.TryGetValue("type", out typeStr); break; }
            }
        }

        var needed = new HashSet<string> { "ActivityClientData.xlsx" };
        if (typeStr != null)
        {
            if (rules.TypeSubTableRules.TryGetValue(typeStr, out var sub) && sub.Table != null)
                needed.Add(sub.Table.Replace("Tables.", "") + ".xlsx");
            else if (rules.TypeTableMap.TryGetValue(typeStr, out var mapped))
                needed.Add(mapped.Replace("Tables.", "") + ".xlsx");

            if (rules.TypeMultiSubTableRules.TryGetValue(typeStr, out var multi))
                foreach (var xls in multi) needed.Add(xls);
        }

        return all.Where(d => needed.Contains(d.ExcelFile)).ToList();
    }
}
