using System.Data;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

public sealed class ColumnTypeDetectorTests
{
    private static DataTable MakeTable(string columnName, params string?[] values)
    {
        var table = new DataTable();
        table.Columns.Add(columnName, typeof(string));
        foreach (var v in values)
        {
            var row = table.NewRow();
            row[columnName] = v ?? (object)DBNull.Value;
            table.Rows.Add(row);
        }
        return table;
    }

    [Fact]
    public void Detect_AllIntegers_ReturnsInteger()
    {
        var table = MakeTable("Col", "1", "42", "100", "999", "0");
        Assert.Equal(ColumnType.Integer, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_AllFloatsWithDecimals_ReturnsFloat()
    {
        var table = MakeTable("Col", "1.5", "3.14", "100.0", "0.001");
        Assert.Equal(ColumnType.Float, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_MixedIntAndFloat_ReturnsFloat()
    {
        var table = MakeTable("Col", "1", "2.5", "3", "4.7");
        Assert.Equal(ColumnType.Float, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_AllDates_ReturnsDate()
    {
        var table = MakeTable(
            "Col",
            "2024-01-15",
            "2024-06-30",
            "2024-12-25",
            "2023-01-01"
        );
        Assert.Equal(ColumnType.Date, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_FewUniqueValuesManyRows_ReturnsEnum()
    {
        // 25 行，只有 3 个唯一值 → Enum
        var values = new List<string?>(25);
        for (var i = 0; i < 25; i++)
            values.Add(i % 3 == 0 ? "Red" : i % 3 == 1 ? "Green" : "Blue");
        var table = MakeTable("Col", values.ToArray());
        Assert.Equal(ColumnType.Enum, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_MixedText_ReturnsText()
    {
        var table = MakeTable("Col", "abc", "123", "hello", "world", "42");
        Assert.Equal(ColumnType.Text, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_AllNull_ReturnsText()
    {
        var table = MakeTable("Col", null, null, null);
        Assert.Equal(ColumnType.Text, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_EmptyTable_ReturnsText()
    {
        var table = new DataTable();
        table.Columns.Add("Col", typeof(string));
        Assert.Equal(ColumnType.Text, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_SampleSizeTruncation_OnlyChecksFirstNRows()
    {
        // 前 100 行都是整数，但第 101 行是文本
        // sampleSize=100 时只看前 100 行 → Integer
        var values = new List<string?>(101);
        for (var i = 0; i < 100; i++)
            values.Add("42");
        values.Add("not-a-number");
        var table = MakeTable("Col", values.ToArray());
        Assert.Equal(ColumnType.Integer, ColumnTypeDetector.Detect(table, "Col", sampleSize: 100));
    }

    [Fact]
    public void Detect_NonExistentColumn_Throws()
    {
        var table = MakeTable("Col", "1", "2");
        Assert.Throws<ArgumentException>(() => ColumnTypeDetector.Detect(table, "NonExistent"));
    }

    [Fact]
    public void Detect_IntegersWithNegativeAndZero_ReturnsInteger()
    {
        var table = MakeTable("Col", "-100", "0", "42", "-1", "999");
        Assert.Equal(ColumnType.Integer, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_FloatsWithScientificNotation_ReturnsFloat()
    {
        var table = MakeTable("Col", "1e5", "2.5e-3", "1.0E+2");
        Assert.Equal(ColumnType.Float, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_EnumBoundary_Exactly20UniqueValues_ReturnsEnum()
    {
        // 21 行，20 个唯一值 (Value0..Value18 各 1 个，Value19 重复 2 次) = Enum
        var values = new List<string?>(21);
        for (var i = 0; i < 20; i++)
            values.Add($"Value{i}");
        values.Add("Value19"); // 第 21 行，重复已有值
        var table = MakeTable("Col", values.ToArray());
        Assert.Equal(ColumnType.Enum, ColumnTypeDetector.Detect(table, "Col"));
    }

    [Fact]
    public void Detect_EnumBoundary_21UniqueValues_ReturnsText()
    {
        // 22 行，21 个唯一值 = Text
        var values = new List<string?>(22);
        for (var i = 0; i < 22; i++)
            values.Add($"Value{i}");
        var table = MakeTable("Col", values.ToArray());
        Assert.Equal(ColumnType.Text, ColumnTypeDetector.Detect(table, "Col"));
    }
}