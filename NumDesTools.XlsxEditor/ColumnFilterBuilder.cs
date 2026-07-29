using System.Data;
using System.Globalization;
using System.Text;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// Pure (no-WPF) logic that builds DataView.RowFilter strings for the per-column
/// header filter TextBoxes. Extracted from MainWindow so the filter-expression logic
/// is testable without a WPF Dispatcher.
/// </summary>
public static class ColumnFilterBuilder
{
    /// <summary>
    /// Build a DataView.RowFilter from a dictionary of columnName → filterValue.
    /// - Empty filterValue → that column contributes no clause (no filter).
    /// - Non-empty → exact match: [Column] = 'value' (single quotes escaped).
    /// - Multiple columns → AND.
    /// </summary>
    public static string BuildFilter(IReadOnlyDictionary<string, string> filters)
    {
        var parts = new List<string>(filters.Count);
        foreach (var (col, value) in filters)
        {
            if (string.IsNullOrEmpty(value))
                continue;
            // 转义单引号: DataView.RowFilter 用 '' 转义 '
            var escaped = value.Replace("'", "''");
            parts.Add($"[{col}] = '{escaped}'");
        }
        return string.Join(" AND ", parts);
    }

    /// <summary>
    /// Build a DataView.RowFilter from a dictionary of columnName → (filterValue, columnType).
    /// Type-aware filtering:
    /// <list type="bullet">
    ///   <item><description>Text → LIKE '%val%'（contains，转义 LIKE 通配符）</description></item>
    ///   <item><description>Integer/Float → 支持 &gt;=, &lt;=, &gt;, &lt; 前缀，无前缀精确匹配</description></item>
    ///   <item><description>Date → = #yyyy-MM-dd#</description></item>
    ///   <item><description>Enum → = 'val'（精确匹配）</description></item>
    /// </list>
    /// </summary>
    public static string BuildFilter(
        IReadOnlyDictionary<string, (string value, ColumnType type)> filters
    )
    {
        var parts = new List<string>(filters.Count);
        foreach (var (col, (value, type)) in filters)
        {
            if (string.IsNullOrEmpty(value))
                continue;

            var clause = type switch
            {
                ColumnType.Text => BuildTextClause(col, value),
                ColumnType.Integer => BuildNumericClause(col, value),
                ColumnType.Float => BuildNumericClause(col, value),
                ColumnType.Date => BuildDateClause(col, value),
                ColumnType.Enum => BuildEnumClause(col, value),
                _ => BuildTextClause(col, value),
            };
            parts.Add(clause);
        }
        return string.Join(" AND ", parts);
    }

    private static string BuildTextClause(string column, string value)
    {
        // 转义 LIKE 通配符: 先转 [ 再转 % 再转 _
        var escaped = EscapeLikeWildcards(value);
        // 转义单引号
        escaped = escaped.Replace("'", "''");
        return $"[{column}] LIKE '%{escaped}%'";
    }

    private static string BuildNumericClause(string column, string value)
    {
        var (op, number) = ParseRangePrefix(value);
        return $"[{column}] {op} {number}";
    }

    private static string BuildDateClause(string column, string value)
    {
        // 转义单引号不影响 # 包裹的日期，但以防万一
        var escaped = value.Replace("'", "''");
        return $"[{column}] = #{escaped}#";
    }

    private static string BuildEnumClause(string column, string value)
    {
        var escaped = value.Replace("'", "''");
        return $"[{column}] = '{escaped}'";
    }

    /// <summary>
    /// 转义 DataView.RowFilter LIKE 通配符：% → [%]，_ → [_]，[ → [[]。
    /// 顺序：先转 [，再转 %，再转 _，避免引入的新 [ 被二次转义。
    /// </summary>
    private static string EscapeLikeWildcards(string value)
    {
        // 先转 [ 防止后续替换引入的 [ 被重复转义
        var result = value.Replace("[", "[[]");
        result = result.Replace("%", "[%]");
        result = result.Replace("_", "[_]");
        return result;
    }

    /// <summary>
    /// 解析数字范围前缀：&gt;=, &lt;=, &gt;, &lt;，无前缀则精确匹配 (=)。
    /// </summary>
    private static (string op, string number) ParseRangePrefix(string value)
    {
        if (value.StartsWith(">=", StringComparison.Ordinal))
            return (">=", value[2..]);
        if (value.StartsWith("<=", StringComparison.Ordinal))
            return ("<=", value[2..]);
        if (value.StartsWith('>'))
            return (">", value[1..]);
        if (value.StartsWith('<'))
            return ("<", value[1..]);
        return ("=", value);
    }
}
