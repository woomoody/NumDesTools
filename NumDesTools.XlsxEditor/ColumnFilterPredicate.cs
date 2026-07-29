using System.Globalization;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// P5：类型感知的列头筛选谓词。纯逻辑（不依赖 WPF），供 DataGridExtensions 的
/// <c>ICustomFilter</c> 适配器（<see cref="ColumnStoreFilterAdapter"/>）复用并单测。
/// <para>
/// <see cref="Build"/> 把"列筛选值集合"编译成 <c>Func&lt;int,bool&gt;</c>——按行号读 ColumnStore 对应列
/// （不整表复制、不物化 RowView），多列 AND。谓词直接喂给 <see cref="VirtualizingSortableView.ApplyFilter"/>。
/// </para>
/// </summary>
public static class ColumnFilterPredicate
{
    /// <summary>
    /// 类型感知的单格匹配：Text/Date→包含(OrdinalIgnoreCase)，Integer/Float→支持 &gt;=/&lt;=/&gt;/&lt;/= 前缀，
    /// Enum→精确(Ordinal)。空筛选值恒匹配（不过滤）。
    /// </summary>
    public static bool Matches(string cell, string filterValue, ColumnType type)
    {
        if (string.IsNullOrEmpty(filterValue))
        {
            return true;
        }

        switch (type)
        {
            case ColumnType.Integer:
            case ColumnType.Float:
            {
                var (op, num) = ParseRangePrefix(filterValue);
                if (
                    !double.TryParse(
                        cell,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var cellNum
                    )
                    || !double.TryParse(
                        num,
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var target
                    )
                )
                {
                    return false;
                }

                return op switch
                {
                    ">=" => cellNum >= target,
                    "<=" => cellNum <= target,
                    ">" => cellNum > target,
                    "<" => cellNum < target,
                    _ => Math.Abs(cellNum - target) < double.Epsilon,
                };
            }
            case ColumnType.Enum:
                return string.Equals(cell, filterValue, StringComparison.Ordinal);
            case ColumnType.Date:
            case ColumnType.Text:
            default:
                return cell.Contains(filterValue, StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// 把多列筛选编译成行谓词（多列 AND）。谓词按需读 <paramref name="store"/> 对应列的值——
    /// 只做 O(激活列数) 次 GetCell/行，不整表扫、不物化 RowView。空筛选集合返回恒 true。
    /// </summary>
    public static Func<int, bool> Build(
        ColumnStore store,
        IReadOnlyList<(int Col, string Value, ColumnType Type)> columnFilters
    )
    {
        ArgumentNullException.ThrowIfNull(store);
        ArgumentNullException.ThrowIfNull(columnFilters);

        // 只保留列号合法 + 有非空筛选值的条目（快照，避免闭包捕获可变集合）
        var active = columnFilters
            .Where(f => f.Col >= 0 && f.Col < store.ColumnCount && !string.IsNullOrEmpty(f.Value))
            .ToArray();

        if (active.Length is 0)
        {
            return static _ => true;
        }

        return row =>
        {
            foreach (var (col, value, type) in active)
            {
                var cell = store.GetCell(row, col) ?? string.Empty;
                if (!Matches(cell, value, type))
                {
                    return false;
                }
            }

            return true;
        };
    }

    private static (string Op, string Number) ParseRangePrefix(string value)
    {
        if (value.StartsWith(">=", StringComparison.Ordinal))
        {
            return (">=", value[2..]);
        }

        if (value.StartsWith("<=", StringComparison.Ordinal))
        {
            return ("<=", value[2..]);
        }

        if (value.StartsWith('>'))
        {
            return (">", value[1..]);
        }

        if (value.StartsWith('<'))
        {
            return ("<", value[1..]);
        }

        if (value.StartsWith('='))
        {
            return ("=", value[1..]);
        }

        return ("=", value);
    }
}
