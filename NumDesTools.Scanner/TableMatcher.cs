using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace NumDesTools.Scanner;

/// <summary>
/// 根据需求/缺陷文本匹配涉及的配置表格。
/// 对应 Python 版本的 match_types_from_text() + identify_tables()。
/// </summary>
public static class TableMatcher
{
    private static readonly HashSet<string> ExcludedExcel = ["ActivityServerData.xlsx"];

    private static readonly string[] PlannerKeywords =
    [
        "活动", "LTE", "大富翁", "克隆", "配置", "数值", "策划",
        "type=", "ActivityType", "礼包", "体力", "奖励", "排期",
    ];

    private static readonly HashSet<string> SkipNoteWords = ["废弃", "弃用", "占位"];

    private static readonly Regex RxPhase = new(@"第\s*([0-9０-９]+)\s*期", RegexOptions.Compiled);
    private static readonly Regex RxExplicitType = new(@"type\s*[=＝]\s*(\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex RxLteVariant = new(@"([一-鿿]{1,8}?)玩法[-—\s]*lte", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static bool IsPlannerRelated(WorkItem item)
    {
        var text = (item.Name + " " + item.Desc).ToLower();
        return PlannerKeywords.Any(kw => text.Contains(kw.ToLower()));
    }

    /// <summary>
    /// 从文本中匹配 type 列表，返回去重有序字符串列表。
    /// </summary>
    public static List<string> MatchTypesFromText(string text, IReadOnlyList<(string Note, int Type)> typeIndex)
    {
        var textLower = text.ToLower();
        var found = new List<string>();
        var seen  = new HashSet<int>();

        // 1. 显式 type=XX
        foreach (Match m in RxExplicitType.Matches(text))
        {
            if (int.TryParse(m.Groups[1].Value, out int t) && seen.Add(t))
                found.Insert(0, t.ToString());
        }

        // 2+3. 备注核心双向子串匹配
        foreach (var (note, type) in typeIndex)
        {
            if (SkipNoteWords.Any(w => note.Contains(w))) continue;
            if (seen.Contains(type)) continue;

            // 方向A：备注核心 ⊆ 文本
            if (textLower.Contains(note)) { seen.Add(type); found.Add(type.ToString()); continue; }
            // 方向B：文本 ⊆ 备注核心（核心≥4字防误匹配）
            if (note.Length >= 4 && note.Contains(textLower)) { seen.Add(type); found.Add(type.ToString()); }
        }

        // 4. "XXX玩法-LTE" 变体
        var extras = new List<string>();
        foreach (Match m in RxLteVariant.Matches(textLower))
        {
            var word = m.Groups[1].Value.Trim();
            if (!string.IsNullOrEmpty(word)) extras.Add(word + "lte");
        }
        if (extras.Count > 0)
        {
            foreach (var (note, type) in typeIndex)
            {
                if (SkipNoteWords.Any(w => note.Contains(w))) continue;
                if (seen.Contains(type)) continue;
                foreach (var extra in extras)
                {
                    if (note.Contains(extra) || extra.Contains(note))
                    { seen.Add(type); found.Add(type.ToString()); break; }
                }
            }
        }

        return found;
    }

    /// <summary>
    /// 识别需求涉及的配置表。返回 (tables, phase)。
    /// phase=1 首期/全量；phase>1 续期只含主表+主子表。
    /// </summary>
    public static (List<TableMatch> Tables, int Phase) IdentifyTables(
        WorkItem item, ActivityTableRules rules, IReadOnlyList<(string Note, int Type)> typeIndex)
    {
        var text     = item.Name + " " + item.Desc;
        var typeNums = MatchTypesFromText(text, typeIndex);
        if (typeNums.Count == 0) return ([], 1);

        var phase         = DetectPhase(text);
        var includeMulti  = phase == 1;

        var results   = new List<TableMatch>();
        var seenExcel = new HashSet<string>();
        var tablesMap = rules.Tables.Where(t => !string.IsNullOrEmpty(t.LuaKey))
                                    .ToDictionary(t => t.LuaKey);

        // 主表 ActivityClientData
        var mainDef = rules.Tables.FirstOrDefault(t => t.Name == "ActivityClientData");
        if (mainDef != null && seenExcel.Add(mainDef.ExcelFile))
            results.Add(new TableMatch(mainDef.ExcelFile, mainDef.Desc,
                mainDef.Fields.Where(f => f.Required).Select(f => f.Name).ToList(),
                mainDef.KeyField, null));

        foreach (var tNum in typeNums)
        {
            // 主子表
            string luaKey = "";
            List<FieldDef> subFields = [];
            string lookupField = "activityID";

            if (rules.TypeSubTableRules.TryGetValue(tNum, out var sub))
            {
                luaKey      = sub.Table ?? "";
                subFields   = sub.Fields;
                lookupField = sub.LookupField;
            }
            else if (rules.TypeTableMap.TryGetValue(tNum, out var mapped))
            {
                luaKey = mapped;
            }

            if (!string.IsNullOrEmpty(luaKey))
            {
                if (!luaKey.StartsWith("Tables.")) luaKey = "Tables." + luaKey;
                tablesMap.TryGetValue(luaKey, out var tDef);
                var excel = tDef?.ExcelFile ?? (luaKey.Replace("Tables.", "") + ".xlsx");
                if (seenExcel.Add(excel) && !ExcludedExcel.Contains(excel))
                {
                    var fields = subFields.Count > 0
                        ? subFields.Where(f => f.Required).Select(f => f.Name).ToList()
                        : (tDef?.Fields.Where(f => f.Required).Select(f => f.Name).ToList() ?? []);
                    results.Add(new TableMatch(excel, tDef?.Desc ?? $"type={tNum} 子表",
                        fields, lookupField, tNum));
                }
            }

            // 多子表（仅首期）
            if (includeMulti && rules.TypeMultiSubTableRules.TryGetValue(tNum, out var multiList))
            {
                foreach (var excel in multiList)
                {
                    if (!seenExcel.Add(excel) || ExcludedExcel.Contains(excel)) continue;
                    results.Add(new TableMatch(excel, $"type={tNum} 关联表（activityID 引用）",
                        [], "activityID", tNum));
                }
            }
        }

        return (results.Where(r => !ExcludedExcel.Contains(r.Excel)).ToList(), phase);
    }

    private static int DetectPhase(string text)
    {
        var nums = RxPhase.Matches(text)
            .Select(m => int.TryParse(m.Groups[1].Value, out int n) ? n : 1)
            .Where(n => n >= 1).ToList();
        return nums.Count > 0 ? nums.Min() : 1;
    }
}
