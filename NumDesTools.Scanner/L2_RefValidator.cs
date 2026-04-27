namespace NumDesTools.Scanner;

/// <summary>
/// L2 — 跨表引用完整性校验：
///   1. FieldDef.refTable 指定的字段值必须在目标表的 keyField 中存在
///   2. ActivityClientData.type → typeTableMap / typeSubTableRules 对应子表必须存在该 activityID
///   3. 数组字段（refIsArray=true）拆分后逐个校验
/// </summary>
public static class L2_RefValidator
{
    public static TableValidationReport Validate(
        string               tablesDir,
        TableDef             def,
        ActivityTableRules   rules,
        string?              sheetName = null)
    {
        var report = new TableValidationReport { ExcelFile = def.ExcelFile };
        var excelPath = Path.Combine(tablesDir, def.ExcelFile);
        if (!File.Exists(excelPath)) return report;

        var (_, rows) = ExcelReader.Read(excelPath, sheetName);
        if (rows.Count == 0) return report;

        // 缓存已加载的目标表 key 集合，避免重复读文件
        var keySetCache = new Dictionary<string, HashSet<string>>();

        // 找到有 refTable 的字段
        var refFields = def.Fields.Where(f => !string.IsNullOrEmpty(f.RefTable)).ToList();

        // 是否是主表（ActivityClientData），需要额外做 type→子表 校验
        bool isMainTable = def.Name == "ActivityClientData";

        foreach (var (rowNum, data) in rows)
        {
            // ── 1. refTable 字段校验 ─────────────────────────────────────────
            foreach (var fd in refFields)
            {
                if (!data.TryGetValue(fd.Name, out var rawVal) || string.IsNullOrWhiteSpace(rawVal))
                    continue;

                var refTableDef = rules.Tables.FirstOrDefault(t => t.LuaKey == fd.RefTable);
                var refExcel    = refTableDef?.ExcelFile
                                  ?? fd.RefTable.Replace("Tables.", "") + ".xlsx";
                var refKey      = refTableDef?.KeyField ?? "id";

                var keys = GetOrLoadKeys(tablesDir, refExcel, refKey, keySetCache);

                // 支持数组（"#" 分隔 或 json 数组样式）
                var values = fd.RefIsArray
                    ? SplitArray(rawVal)
                    : [rawVal];

                foreach (var v in values)
                {
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    if (v == "0") continue; // 0 为通用"不使用"占位值，跳过引用校验
                    if (!keys.Contains(v))
                        report.Issues.Add(new ValidationIssue(
                            Severity.Error, "L2", def.ExcelFile, rowNum, fd.Name,
                            $"引用 {refExcel}[{refKey}={v}] 不存在"));
                }
            }

            // ── 2. 主表 type → 子表 activityID 校验 ────────────────────────
            if (!isMainTable) continue;
            if (!data.TryGetValue("type", out var typeStr) || string.IsNullOrWhiteSpace(typeStr)) continue;
            if (!data.TryGetValue("activityID", out var actId) || string.IsNullOrWhiteSpace(actId) || actId == "0") continue;

            string? subExcel = null;
            string  lookupFd = "activityID";

            if (rules.TypeSubTableRules.TryGetValue(typeStr, out var subRule))
            {
                subExcel = subRule.Table?.Replace("Tables.", "") + ".xlsx";
                lookupFd = subRule.LookupField;
            }
            else if (rules.TypeTableMap.TryGetValue(typeStr, out var mapped))
            {
                subExcel = mapped.Replace("Tables.", "") + ".xlsx";
            }

            if (string.IsNullOrEmpty(subExcel)) continue;

            var subKeys = GetOrLoadKeys(tablesDir, subExcel, lookupFd, keySetCache);
            if (subKeys.Count == 0) continue; // 子表文件不存在，跳过（不重复报缺失）

            if (!subKeys.Contains(actId))
                report.Issues.Add(new ValidationIssue(
                    Severity.Error, "L2", def.ExcelFile, rowNum, "activityID",
                    $"type={typeStr} 子表 {subExcel} 中找不到 {lookupFd}={actId}"));
        }

        return report;
    }

    // ── 辅助 ────────────────────────────────────────────────────────────────
    private static HashSet<string> GetOrLoadKeys(
        string tablesDir, string excelFile, string keyField,
        Dictionary<string, HashSet<string>> cache)
    {
        var cacheKey = $"{excelFile}::{keyField}";
        if (cache.TryGetValue(cacheKey, out var cached)) return cached;

        var path = Path.Combine(tablesDir, excelFile);
        if (!File.Exists(path))
        {
            cache[cacheKey] = [];
            return [];
        }

        var set = ExcelReader.ReadKeySet(path, keyField);
        cache[cacheKey] = set;
        return set;
    }

    // 解析 "#" 分隔 或 "[1,2,3]" 格式的数组字段
    private static List<string> SplitArray(string raw)
    {
        raw = raw.Trim();
        if (raw.StartsWith('[') && raw.EndsWith(']'))
            raw = raw[1..^1];
        return [.. raw.Split(['#', ','], StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => s.Trim())
                      .Where(s => !string.IsNullOrEmpty(s))];
    }
}
