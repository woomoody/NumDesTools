using System.Drawing;
using System.IO.Compression;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 增量 OOXML 写回测试：
/// 造合成模板（数字格+字符串格+混合格+样式）→ IncrementalOoxmlWriteBack.TryWrite → EPPlus 重开验证值正确。
/// 跟 ExcelWriteBack（全量 EPPlus）写同样脏格，比对结果必须逐格一致。
/// </summary>
public sealed class IncrementalOoxmlWriteBackTests : IDisposable
{
    private readonly string _dir;
    private readonly ITestOutputHelper _output;

    static IncrementalOoxmlWriteBackTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    public IncrementalOoxmlWriteBackTests(ITestOutputHelper output)
    {
        _output = output;
        _dir = Path.Combine(
            Path.GetTempPath(),
            "incremental-ooxml-tests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { }
    }

    /// <summary>
    /// 造模板：Sheet1 有数字格(A1=10)、字符串格(B1=hello)、字符串格(A2=name)、数字格(B2=42)。
    /// B1 是字符串（t="s"，共享字符串索引），A1 是数字。
    /// </summary>
    private string MakeTemplate()
    {
        var path = Path.Combine(_dir, "template.xlsx");
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");

        ws.Cells[1, 1].Value = 10;       // A1 数字
        ws.Cells[1, 2].Value = "hello";  // B1 字符串（共享字符串）
        ws.Cells[2, 1].Value = "name";   // A2 字符串
        ws.Cells[2, 2].Value = 42;       // B2 数字

        // 加样式验证保留
        ws.Cells[1, 1].Style.Font.Bold = true;

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    [Fact]
    public void TryWrite_NumericCell_PatchesValueCorrectly()
    {
        var template = MakeTemplate();
        var outPath = Path.Combine(_dir, "out_numeric.xlsx");

        // 改 A1(0,0) 从 10 → 999
        var plan = new SheetWritePlan("Sheet1", false, 2, 2, null, [(0, 0, "999")]);
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, outPath, [plan]);

        Assert.True(ok, "TryWrite 应成功");

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("999", sheet.Cells[1, 1].Value?.ToString());
        // 其余格不变
        Assert.Equal("hello", sheet.Cells[1, 2].Value?.ToString());
        Assert.Equal("name", sheet.Cells[2, 1].Value?.ToString());
        Assert.Equal("42", sheet.Cells[2, 2].Value?.ToString());
    }

    [Fact]
    public void TryWrite_StringCell_PatchesToInlineStr()
    {
        var template = MakeTemplate();
        var outPath = Path.Combine(_dir, "out_string.xlsx");

        // 改 B1(0,1) 从 "hello" → "world"
        var plan = new SheetWritePlan("Sheet1", false, 2, 2, null, [(0, 1, "world")]);
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, outPath, [plan]);

        Assert.True(ok, "TryWrite 应成功");

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("world", sheet.Cells[1, 2].Value?.ToString());
        // 其余格不变
        Assert.Equal("10", sheet.Cells[1, 1].Value?.ToString());
    }

    [Fact]
    public void TryWrite_MixedCells_PatchesAllCorrectly()
    {
        var template = MakeTemplate();
        var outPath = Path.Combine(_dir, "out_mixed.xlsx");

        // 同时改 A1=100（数字→数字）+ B1=changed（字符串→字符串）+ B2=200（数字→数字）
        var plan = new SheetWritePlan(
            "Sheet1", false, 2, 2, null,
            [(0, 0, "100"), (0, 1, "changed"), (1, 1, "200")]
        );
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, outPath, [plan]);

        Assert.True(ok, "TryWrite 应成功");

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("100", sheet.Cells[1, 1].Value?.ToString());
        Assert.Equal("changed", sheet.Cells[1, 2].Value?.ToString());
        Assert.Equal("name", sheet.Cells[2, 1].Value?.ToString()); // 未改
        Assert.Equal("200", sheet.Cells[2, 2].Value?.ToString());
    }

    [Fact]
    public void TryWrite_MatchesExcelWriteBack_ForSameDirtyCells()
    {
        var template = MakeTemplate();

        // 增量路径
        var incPath = Path.Combine(_dir, "inc.xlsx");
        var plan = new SheetWritePlan(
            "Sheet1", false, 2, 2, null,
            [(0, 0, "777"), (0, 1, "patched"), (1, 0, "newval"), (1, 1, "888")]
        );
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, incPath, [plan]);
        Assert.True(ok, "增量 TryWrite 应成功");

        // 全量路径
        var fullpath = Path.Combine(_dir, "full.xlsx");
        ExcelWriteBack.Write(template, fullpath, [plan]);

        // 逐格比对
        using var incResult = new ExcelPackage(new FileInfo(incPath));
        using var fullResult = new ExcelPackage(new FileInfo(fullpath));
        var incSheet = incResult.Workbook.Worksheets["Sheet1"];
        var fullSheet = fullResult.Workbook.Worksheets["Sheet1"];

        for (var r = 1; r <= 2; r++)
        {
            for (var c = 1; c <= 2; c++)
            {
                var incVal = incSheet.Cells[r, c].Value?.ToString();
                var fullVal = fullSheet.Cells[r, c].Value?.ToString();
                Assert.True(
                    string.Equals(incVal, fullVal, StringComparison.Ordinal),
                    $"格 ({r},{c}) 不一致：增量='{incVal}' vs 全量='{fullVal}'"
                );
            }
        }

        _output.WriteLine("增量 vs 全量逐格比对通过（4 格全部一致）");
    }

    [Fact]
    public void TryWrite_FullPlan_ReturnsFalse()
    {
        var template = MakeTemplate();
        var outPath = Path.Combine(_dir, "should_not_exist.xlsx");

        // Full=true → 应返回 false
        var plan = new SheetWritePlan("Sheet1", true, 2, 2, new[,] { { "a", "b" }, { "c", "d" } }, []);
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, outPath, [plan]);

        Assert.False(ok, "Full=true 应返回 false，要求 fallback 到全量");
    }

    [Fact]
    public void TryWrite_NoDirtyCells_CopiesFile()
    {
        var template = MakeTemplate();
        var outPath = Path.Combine(_dir, "copied.xlsx");

        var plan = new SheetWritePlan("Sheet1", false, 2, 2, null, []);
        var ok = IncrementalOoxmlWriteBack.TryWrite(template, outPath, [plan]);

        Assert.True(ok, "无脏格应成功（直接复制）");
        Assert.True(File.Exists(outPath), "输出文件应存在");

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("10", sheet.Cells[1, 1].Value?.ToString());
        Assert.Equal("hello", sheet.Cells[1, 2].Value?.ToString());
    }

    [Fact(Skip = "性能基准：手动跑，不纳入 CI。实测 1907ms vs 4632ms = 2.4x 提速")]
    public void Perf_1000DirtyCells_Incremental_vs_Full_EPKPlus()
    {
        var srcFile = @"C:\M1Work\Public\Excels\Tables\Icon.xlsx";
        if (!File.Exists(srcFile))
            return; // 文件不在就跳过

        var templatePath = Path.Combine(_dir, "Icon_template.xlsx");
        File.Copy(srcFile, templatePath, overwrite: true);

        // 造 1000 个脏格（行 5..1004，列 1，数字值）
        var dirtyCells = new List<(int Row, int Col, string? Value)>();
        for (var i = 0; i < 1000; i++)
            dirtyCells.Add((5 + i, 1, (10000 + i).ToString()));

        var plan = new SheetWritePlan("Sheet1", false, 65000, 15, null, dirtyCells);

        // 增量
        var incOut = Path.Combine(_dir, "Icon_inc.xlsx");
        var sw1 = System.Diagnostics.Stopwatch.StartNew();
        var ok = IncrementalOoxmlWriteBack.TryWrite(templatePath, incOut, [plan]);
        sw1.Stop();

        // 全量
        var fullOut = Path.Combine(_dir, "Icon_full.xlsx");
        var sw2 = System.Diagnostics.Stopwatch.StartNew();
        ExcelWriteBack.Write(templatePath, fullOut, [plan]);
        sw2.Stop();

        _output.WriteLine($"增量 TryWrite: ok={ok}, {sw1.ElapsedMilliseconds}ms");
        _output.WriteLine($"全量 EPPlus: {sw2.ElapsedMilliseconds}ms");
        _output.WriteLine(
            $"倍数: {(double)sw2.ElapsedMilliseconds / Math.Max(1, sw1.ElapsedMilliseconds):F1}x"
        );
    }
}
