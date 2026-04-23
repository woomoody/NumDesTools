using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 扫描 EnumCmds.lua.txt / ActivityManager.lua.txt / 各 LogicBase 文件，
/// 自动推断每个 ActivityType 对应的子表名，
/// 将缺失项追加到 ActivityTableRules.json 的 typeTableMap（只增不改）。
/// </summary>
public static class ActivityRulesUpdater
{
    // ── 路径 ─────────────────────────────────────────────────────────────────
    private static string RulesFilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                     "NumDesTools", "Config", "ActivityTableRules.json");

    private static readonly string LuaRoot =
        @"C:\M1Work\Code\Assets\LuaScripts";

    private static string EnumCmdsPath =>
        Path.Combine(LuaRoot, "Logics", "DataStructures", "Definitions", "EnumCmds.lua.txt");

    private static string ActivityManagerPath =>
        Path.Combine(LuaRoot, "Logics", "Controller", "Activity", "ActivityManager.lua.txt");

    // ── 正则 ─────────────────────────────────────────────────────────────────
    // ActivityType 枚举成员：  EnumName = 数字,
    private static readonly Regex RxEnumEntry =
        new(@"(\w+)\s*=\s*(\d+)", RegexOptions.Compiled);

    // ActivityManager if/elseif 块：
    //   data.type == EnumCmds.ActivityType.XxxName then
    //       logic = SomeLogicBase:Create(...)
    private static readonly Regex RxTypeBlock =
        new(@"ActivityType\.(\w+)\s*then\s*\r?\n\s*logic\s*=\s*(\w+):Create",
            RegexOptions.Compiled | RegexOptions.Multiline);

    // LogicBase 文件里首行 activityID 子表引用：
    //   self.config = Tables.XxxData[self.data.activityID]
    private static readonly Regex RxTableRef =
        new(@"Tables\.(\w+)\[self\.data\.activityID\]", RegexOptions.Compiled);

    // ═════════════════════════════════════════════════════════════════════════
    // 公共入口
    // ═════════════════════════════════════════════════════════════════════════

    public static void Run()
    {
        var report = new StringBuilder();
        report.AppendLine("═══════════════ ActivityTableRules 更新报告 ═══════════════");
        report.AppendLine($"时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        report.AppendLine();

        NumDesAddIn.App.StatusBar = "更新规则：读取枚举...";

        // 1. 解析 ActivityType 枚举 → { 枚举名 → type数字 }
        var enumMap = ParseActivityTypeEnum(report);
        if (enumMap.Count == 0)
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtpNormal(report.ToString());
            return;
        }

        // 2. 解析 ActivityManager → { 枚举名 → LogicBase类名 }
        NumDesAddIn.App.StatusBar = "更新规则：读取 ActivityManager...";
        var logicMap = ParseActivityManagerMapping(report);

        // 3. 对每个 LogicBase 扫描 Tables.Xxx[activityID] → { LogicBase类名 → 子表luaKey }
        NumDesAddIn.App.StatusBar = "更新规则：扫描 LogicBase 文件...";
        var tableMap = BuildLogicToTableMap(logicMap.Values.Distinct().ToList(), report);

        // 4. 合并：enumName → typeNum → logicBase → luaKey
        var inferred = new Dictionary<string, string>(); // typeNum(string) → luaKey
        foreach (var (enumName, typeNum) in enumMap)
        {
            if (!logicMap.TryGetValue(enumName, out var logic)) continue;
            if (!tableMap.TryGetValue(logic, out var luaKey)) continue;
            inferred[typeNum.ToString()] = luaKey;
        }

        report.AppendLine($"推断出有子表的 type 数：{inferred.Count}");
        report.AppendLine();

        // 5. 读取现有 JSON，只追加 typeTableMap 中缺失的项
        NumDesAddIn.App.StatusBar = "更新规则：写入 JSON...";
        var (added, skipped) = PatchRulesJson(inferred, report);

        report.AppendLine();
        report.AppendLine($"═════ 完成：新增 {added} 条，已有跳过 {skipped} 条 ══════");
        NumDesAddIn.App.StatusBar = $"规则更新完成（新增 {added} 条）";

        ErrorLogCtp.DisposeCtp();
        ErrorLogCtp.CreateCtpNormal(report.ToString());
    }

    // ═════════════════════════════════════════════════════════════════════════
    // 内部实现
    // ═════════════════════════════════════════════════════════════════════════

    /// <summary>解析 EnumCmds.lua.txt 中的 ActivityType = { ... } 块。</summary>
    private static Dictionary<string, int> ParseActivityTypeEnum(StringBuilder report)
    {
        if (!File.Exists(EnumCmdsPath))
        {
            report.AppendLine($"❌ 找不到 EnumCmds.lua.txt：{EnumCmdsPath}");
            return new();
        }

        var text  = File.ReadAllText(EnumCmdsPath, Encoding.UTF8);
        var start = text.IndexOf("ActivityType = {", StringComparison.Ordinal);
        if (start < 0)
        {
            report.AppendLine("❌ EnumCmds.lua.txt 中找不到 ActivityType = {");
            return new();
        }

        // 截取到对应的 } 结束
        var depth = 0;
        var end   = start;
        for (; end < text.Length; end++)
        {
            if (text[end] == '{') depth++;
            else if (text[end] == '}') { depth--; if (depth == 0) break; }
        }

        var block  = text.Substring(start, end - start + 1);
        var result = new Dictionary<string, int>();
        foreach (Match m in RxEnumEntry.Matches(block))
        {
            var name = m.Groups[1].Value;
            if (int.TryParse(m.Groups[2].Value, out var num))
                result[name] = num;
        }

        report.AppendLine($"ActivityType 枚举条数：{result.Count}");
        return result;
    }

    /// <summary>解析 ActivityManager 中 type → LogicBase 的 if/elseif 映射。</summary>
    private static Dictionary<string, string> ParseActivityManagerMapping(StringBuilder report)
    {
        if (!File.Exists(ActivityManagerPath))
        {
            report.AppendLine($"⚠ 找不到 ActivityManager.lua.txt：{ActivityManagerPath}");
            return new();
        }

        var text   = File.ReadAllText(ActivityManagerPath, Encoding.UTF8);
        var result = new Dictionary<string, string>();
        foreach (Match m in RxTypeBlock.Matches(text))
        {
            var enumName  = m.Groups[1].Value;
            var logicName = m.Groups[2].Value;
            result.TryAdd(enumName, logicName);
        }

        report.AppendLine($"ActivityManager type→Logic 映射条数：{result.Count}");
        return result;
    }

    /// <summary>
    /// 对每个 LogicBase 类名在 LuaRoot 下全量搜索对应 .lua.txt 文件，
    /// 提取 Tables.Xxx[self.data.activityID] 的第一条引用作为子表。
    /// </summary>
    private static Dictionary<string, string> BuildLogicToTableMap(
        List<string> logicNames, StringBuilder report)
    {
        // 建立 logicName(lower) → 文件路径 的索引（一次性扫描目录）
        var fileIndex = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in Directory.EnumerateFiles(LuaRoot, "*.lua.txt", SearchOption.AllDirectories))
        {
            var baseName = Path.GetFileNameWithoutExtension(
                               Path.GetFileNameWithoutExtension(f)); // 去掉 .lua.txt
            fileIndex.TryAdd(baseName, f);
        }

        var result = new Dictionary<string, string>();
        foreach (var logic in logicNames)
        {
            if (!fileIndex.TryGetValue(logic, out var path)) continue;

            var text = File.ReadAllText(path, Encoding.UTF8);
            var m    = RxTableRef.Match(text);
            if (m.Success)
                result.TryAdd(logic, "Tables." + m.Groups[1].Value);
        }

        report.AppendLine($"LogicBase → 子表 推断成功：{result.Count}/{logicNames.Count}");
        return result;
    }

    /// <summary>
    /// 读取 ActivityTableRules.json，将 inferred 中 typeTableMap 缺失的项追加进去并保存。
    /// 返回 (新增数, 跳过数)。
    /// </summary>
    private static (int added, int skipped) PatchRulesJson(
        Dictionary<string, string> inferred, StringBuilder report)
    {
        if (!File.Exists(RulesFilePath))
        {
            report.AppendLine($"❌ 规则文件不存在：{RulesFilePath}");
            return (0, 0);
        }

        JObject root;
        try
        {
            root = JObject.Parse(File.ReadAllText(RulesFilePath, Encoding.UTF8));
        }
        catch (Exception ex)
        {
            report.AppendLine($"❌ 解析 JSON 失败：{ex.Message}");
            return (0, 0);
        }

        if (root["typeTableMap"] is not JObject typeTableMap)
        {
            typeTableMap = new JObject();
            root["typeTableMap"] = typeTableMap;
        }

        var added   = 0;
        var skipped = 0;

        // 按 type 数字排序输出
        foreach (var (typeNum, luaKey) in inferred.OrderBy(kv => int.Parse(kv.Key)))
        {
            if (typeTableMap.ContainsKey(typeNum))
            {
                skipped++;
                continue;
            }
            typeTableMap[typeNum] = luaKey;
            report.AppendLine($"  + type={typeNum,-4} → {luaKey}");
            added++;
        }

        if (added > 0)
        {
            // 按 key 数字排序后写回
            var sorted = new JObject(
                typeTableMap.Properties().OrderBy(p =>
                    int.TryParse(p.Name, out var n) ? n : int.MaxValue));
            root["typeTableMap"] = sorted;

            File.WriteAllText(RulesFilePath,
                root.ToString(Formatting.Indented), Encoding.UTF8);
        }

        return (added, skipped);
    }
}
