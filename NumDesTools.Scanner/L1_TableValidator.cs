namespace NumDesTools.Scanner;

/// <summary>
/// L1 — 表格自身校验：
///   1. 必填字段为空
///   2. keyField（主键）重复
///   3. 声明为 number 的字段填了非数字内容
///   4. activityID / id 字段为 0 或负数
/// </summary>
public static class L1_TableValidator
{
    private static readonly HashSet<string> NumberTypeKeywords = ["int", "number", "float", "double"];

    public static TableValidationReport Validate(
        string            excelPath,
        TableDef          def,
        string?           sheetName = null)
    {
        var report = new TableValidationReport { ExcelFile = def.ExcelFile };

        if (!File.Exists(excelPath))
        {
            report.Issues.Add(new ValidationIssue(
                Severity.Warning, "L1", def.ExcelFile, 0, "-",
                $"文件不存在，已跳过: {excelPath}"));
            return report;
        }

        var (fields, rows) = ExcelReader.Read(excelPath, sheetName);
        if (rows.Count == 0) return report;

        // 预先读一次 Row3（类型行）来判断 number 列
        var numberFields = BuildNumberFields(excelPath, fields, sheetName);

        // 主键去重
        var keyField  = def.KeyField;
        var seenKeys  = new Dictionary<string, int>(); // value → 首次出现行

        foreach (var (rowNum, data) in rows)
        {
            // 1. 必填字段
            foreach (var fd in def.Fields.Where(f => f.Required))
            {
                if (!data.TryGetValue(fd.Name, out var v) || string.IsNullOrWhiteSpace(v))
                    report.Issues.Add(new ValidationIssue(
                        Severity.Error, "L1", def.ExcelFile, rowNum, fd.Name,
                        "必填字段为空"));
            }

            // 2. 主键重复
            if (!string.IsNullOrEmpty(keyField)
                && data.TryGetValue(keyField, out var keyVal)
                && !string.IsNullOrWhiteSpace(keyVal))
            {
                if (seenKeys.TryGetValue(keyVal, out int firstRow))
                    report.Issues.Add(new ValidationIssue(
                        Severity.Error, "L1", def.ExcelFile, rowNum, keyField,
                        $"主键重复：值 {keyVal} 首次出现于 Row {firstRow}"));
                else
                    seenKeys[keyVal] = rowNum;
            }

            // 3. number 字段非数字
            foreach (var fn in numberFields)
            {
                if (!data.TryGetValue(fn, out var nv) || string.IsNullOrWhiteSpace(nv)) continue;
                if (!IsNumeric(nv))
                    report.Issues.Add(new ValidationIssue(
                        Severity.Error, "L1", def.ExcelFile, rowNum, fn,
                        $"数值字段含非数字内容: \"{nv}\""));
            }

            // 4. id / activityID 不能为 0 或负
            foreach (var fn in new[] { "id", "activityID" })
            {
                if (!data.TryGetValue(fn, out var iv) || string.IsNullOrWhiteSpace(iv)) continue;
                if (double.TryParse(iv, out double d) && d <= 0)
                    report.Issues.Add(new ValidationIssue(
                        Severity.Error, "L1", def.ExcelFile, rowNum, fn,
                        $"ID 字段值 {iv} ≤ 0，不合法"));
            }
        }

        return report;
    }

    // ── 辅助：从 Row3 读类型声明，找出声明为数值类型的字段名 ───────────────────
    private static HashSet<string> BuildNumberFields(string path, List<string> fields, string? sheetName)
    {
        var result = new HashSet<string>();
        try
        {
            OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
            using var pkg = new OfficeOpenXml.ExcelPackage(new FileInfo(path));
            var ws = sheetName != null
                ? pkg.Workbook.Worksheets[sheetName] ?? pkg.Workbook.Worksheets[0]
                : pkg.Workbook.Worksheets[0];
            if (ws == null) return result;

            int typeRow = 3;
            int colCount = ws.Dimension?.Columns ?? 0;
            // 重建列索引（和 ExcelReader 一致：跳过 # 开头的列）
            int fi = 0;
            for (int c = 1; c <= colCount && fi < fields.Count; c++)
            {
                var header = ws.Cells[ExcelReader.HeaderRow, c].Text?.Trim() ?? "";
                if (string.IsNullOrEmpty(header) || header.StartsWith('#')) continue;
                var typeDecl = ws.Cells[typeRow, c].Text?.Trim().ToLower() ?? "";
                // 排除数组类型（int[]、int[][]等），只保留纯标量数值类型
                if (NumberTypeKeywords.Any(k => typeDecl.StartsWith(k)) && !typeDecl.Contains('['))
                    result.Add(fields[fi]);
                fi++;
            }
        }
        catch { /* 读类型行失败不影响主流程 */ }
        return result;
    }

    private static bool IsNumeric(string s)
        => double.TryParse(s, System.Globalization.NumberStyles.Any,
                           System.Globalization.CultureInfo.InvariantCulture, out _);
}
