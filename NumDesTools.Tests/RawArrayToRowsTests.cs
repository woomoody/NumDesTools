using System.Collections.Generic;
using Xunit;

namespace NumDesTools.Tests;

public class RawArrayToRowsTests
{
    // ── helpers ──────────────────────────────────────────────────────────────

    /// <summary>
    /// 构造 1-based object[,]，模拟 COM UsedRange.Value2 的返回格式。
    /// values[0] 对应 raw[1,*]，values[0][0] 对应 raw[1,1]。
    /// </summary>
    private static object[,] MakeRaw(params object?[][] values)
    {
        int rows = values.Length;
        int cols = values.Length > 0 ? values[0].Length : 0;
        var raw = new object[rows + 1, cols + 1]; // 1-based → 上界 = count
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < values[r].Length; c++)
                raw[r + 1, c + 1] = values[r][c];
        return raw;
    }

    private static IDictionary<string, object> Row(List<dynamic> rows, int index) =>
        (IDictionary<string, object>)rows[index];

    // ── tracer bullet ────────────────────────────────────────────────────────

    [Fact]
    public void BasicConversion_ReturnsRowsWithLetterKeys()
    {
        var raw = MakeRaw(
            new object?[] { "#", "id", "val" },
            new object?[] { "#", "int", "string" },
            new object?[] { "#", "123", "hello" }
        );

        var result = NumDesAddIn.RawArrayToRows(raw);

        Assert.Equal(3, result.Count);
        Assert.Equal("#",     Row(result, 0)["A"]);
        Assert.Equal("id",    Row(result, 0)["B"]);
        Assert.Equal("val",   Row(result, 0)["C"]);
        Assert.Equal("123",   Row(result, 2)["B"]);
        Assert.Equal("hello", Row(result, 2)["C"]);
    }

    // ── row / col count ──────────────────────────────────────────────────────

    [Fact]
    public void RowAndColCount_MatchInput()
    {
        var raw = MakeRaw(
            new object?[] { "r1c1", "r1c2", "r1c3", "r1c4" },
            new object?[] { "r2c1", "r2c2", "r2c3", "r2c4" },
            new object?[] { "r3c1", "r3c2", "r3c3", "r3c4" }
        );

        var result = NumDesAddIn.RawArrayToRows(raw);

        Assert.Equal(3, result.Count);
        Assert.Equal(4, ((IDictionary<string, object>)result[0]).Count);
    }

    // ── column letter mapping ────────────────────────────────────────────────

    [Fact]
    public void ColumnMapping_Col1IsA_Col26IsZ_Col27IsAA()
    {
        // 构造 27 列的单行数组
        var cols = new object?[27];
        for (int i = 0; i < 27; i++)
            cols[i] = $"v{i + 1}";

        var raw = MakeRaw(cols, cols); // 2行，满足 rowCount >= 2

        var result = NumDesAddIn.RawArrayToRows(raw);
        var row = Row(result, 0);

        Assert.Equal("v1",  row["A"]);
        Assert.Equal("v26", row["Z"]);
        Assert.Equal("v27", row["AA"]);
    }

    // ── null preservation ────────────────────────────────────────────────────

    [Fact]
    public void NullValues_ArePreserved()
    {
        var raw = MakeRaw(
            new object?[] { "#", null, "x" },
            new object?[] { "#", null, "y" }
        );

        var result = NumDesAddIn.RawArrayToRows(raw);

        Assert.Null(Row(result, 0)["B"]);
        Assert.Null(Row(result, 1)["B"]);
    }

    // ── fewer than 2 rows ────────────────────────────────────────────────────

    [Fact]
    public void FewerThanTwoRows_ReturnsEmpty()
    {
        var raw = MakeRaw(
            new object?[] { "#", "id" }   // 只有 1 行
        );

        var result = NumDesAddIn.RawArrayToRows(raw);

        Assert.Empty(result);
    }

    [Fact]
    public void ZeroRows_ReturnsEmpty()
    {
        // GetUpperBound(0) == 0 → rowCount < 2
        var raw = new object[1, 3]; // 1-based 空，只有下标 0
        var result = NumDesAddIn.RawArrayToRows(raw);
        Assert.Empty(result);
    }
}
