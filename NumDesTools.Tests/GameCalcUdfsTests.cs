namespace NumDesTools.Tests;

public class GameCalcUdfsTests
{
    // ── AliceLtePoisonNear ───────────────────────────────────────────────────

    [Fact]
    public void AliceLtePoisonNear_NonIntegerBasePos_ReturnsExcelErrorValue()
    {
        // basePos 含非整数时应返回 #VALUE!，而非以 (0,0) 静默计算
        object[,] basePos =
        {
            { "abc", "xyz" },
        };
        var result = ExcelUdf.AliceLtePoisonNear(basePos, "5,3|2,7", @"(\d+),(\d+)", "1");
        Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorValue, result);
    }

    [Fact]
    public void AliceLtePoisonNear_ValidInput_ReturnsNearestCoord()
    {
        // 基准 (0,0)，目标 (5,3) 和 (2,7) → 距离 34 vs 53 → 最近是 (5,3)
        object[,] basePos =
        {
            { 0, 0 },
        };
        var result = ExcelUdf.AliceLtePoisonNear(basePos, "5,3|2,7", @"(\d+),(\d+)", "1");
        Assert.Equal("5,3", result);
    }

    [Fact]
    public void AliceLtePoisonNear_InvalidGroupValues_DoesNotThrow()
    {
        // Regex 匹配到非数字 group，不应 FormatException
        object[,] basePos =
        {
            { 1, 1 },
        };
        var ex = Record.Exception(() =>
            ExcelUdf.AliceLtePoisonNear(basePos, "ax,by", @"([a-z]+),([a-z]+)", "1")
        );
        Assert.Null(ex);
    }

    // ── AliceLtePoison ───────────────────────────────────────────────────────

    [Fact]
    public void AliceLtePoison_InvalidGroupValues_DoesNotThrow()
    {
        var ex = Record.Exception(() => ExcelUdf.AliceLtePoison("ax,by", @"([a-z]+),([a-z]+)"));
        Assert.Null(ex);
    }

    [Fact]
    public void AliceLtePoison_ValidInput_ReturnsFormattedCoords()
    {
        // 匹配 "3,4" 和 "1,2" → 应包含 {21,3,4} 和 {21,1,2}
        var result = ExcelUdf.AliceLtePoison("3,4|1,2", @"(\d+),(\d+)");
        Assert.Contains("21,3,4", result);
        Assert.Contains("21,1,2", result);
    }
}
