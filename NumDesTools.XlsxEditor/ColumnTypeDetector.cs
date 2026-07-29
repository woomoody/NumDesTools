using System.Data;
using System.Globalization;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 根据 DataTable 列数据样本判断列类型。
/// 采样前 N 行非空值，按优先级判断：Integer → Float → Date → Enum → Text。
/// </summary>
public enum ColumnType
{
    Text,
    Integer,
    Float,
    Date,
    Enum,
}

public static class ColumnTypeDetector
{
    /// <summary>
    /// 采样前 <paramref name="sampleSize"/> 行非空值，判断指定列的类型。
    /// </summary>
    /// <param name="table">数据表。</param>
    /// <param name="columnName">列名。</param>
    /// <param name="sampleSize">采样行数，默认 100。</param>
    /// <returns>推断的 <see cref="ColumnType"/>。</returns>
    /// <exception cref="ArgumentException">列不存在时抛出。</exception>
    public static ColumnType Detect(
        DataTable table,
        string columnName,
        int sampleSize = 100
    )
    {
        if (!table.Columns.Contains(columnName))
            throw new ArgumentException(
                $"列 '{columnName}' 不存在于 DataTable 中。",
                nameof(columnName)
            );

        // 采集前 sampleSize 行的非空值
        var nonNullValues = new List<string>(sampleSize);
        var maxRows = Math.Min(sampleSize, table.Rows.Count);
        for (var i = 0; i < maxRows; i++)
        {
            var val = table.Rows[i][columnName];
            if (val is not DBNull && val is not null)
            {
                var str = val.ToString()?.Trim();
                if (!string.IsNullOrEmpty(str))
                    nonNullValues.Add(str);
            }
        }

        // 全空 → Text
        if (nonNullValues.Count == 0)
            return ColumnType.Text;

        // Integer: 全部能解析为 long
        if (nonNullValues.All(CanParseAsLong))
            return ColumnType.Integer;

        // Float: 全部能解析为 double（含整数）
        if (nonNullValues.All(CanParseAsDouble))
            return ColumnType.Float;

        // Date: 全部能解析为 DateTime
        if (nonNullValues.All(CanParseAsDateTime))
            return ColumnType.Date;

        // Enum: 唯一值 ≤ 20 且非空值数 > 20
        var distinctCount = nonNullValues.Distinct(StringComparer.Ordinal).Count();
        if (distinctCount <= 20 && nonNullValues.Count > 20)
            return ColumnType.Enum;

        return ColumnType.Text;
    }

    private static bool CanParseAsLong(string value)
    {
        return long.TryParse(
            value,
            NumberStyles.Integer,
            CultureInfo.InvariantCulture,
            out _
        );
    }

    private static bool CanParseAsDouble(string value)
    {
        return double.TryParse(
            value,
            NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.InvariantCulture,
            out _
        );
    }

    private static bool CanParseAsDateTime(string value)
    {
        return DateTime.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out _
        );
    }
}