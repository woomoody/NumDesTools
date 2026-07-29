using System.Drawing;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// P4 WF2 写回优化测试：
/// ① 合成文件（带 chart + 公式 + 自定义列宽 + 样式）→ ExcelWriteBack.Write → 独立重开断言
///    chart 无、公式无（保留计算值）、列宽/样式在、值正确。
/// ② 增量 vs 全量写：Full=false 只写 DirtyCells，Full=true 整表重写（结构变更 fallback）。
/// 纯 IO，不依赖 WPF。所有文件写在临时目录，测试后清理。
/// </summary>
public sealed class ExcelWriteBackTests : IDisposable
{
    private readonly string _dir;
    private readonly ITestOutputHelper _output;

    static ExcelWriteBackTests() => ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    public ExcelWriteBackTests(ITestOutputHelper output)
    {
        _output = output;
        _dir = Path.Combine(
            Path.GetTempPath(),
            "xlsx-writeback-tests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch
        {
            // 清理失败不致命
        }
    }

    /// <summary>
    /// 造合成模板：Sheet1 有值、一个公式格(C1=A1+... 取和)、自定义列宽、粗体+背景色样式、一个折线图。
    /// 返回文件路径。
    /// </summary>
    private string MakeSyntheticTemplate()
    {
        var path = Path.Combine(_dir, "synthetic.xlsx");
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");

        // 值
        ws.Cells[1, 1].Value = "10"; // A1
        ws.Cells[1, 2].Value = "20"; // B1
        ws.Cells[2, 1].Value = "hdr";
        ws.Cells[2, 2].Value = "data";

        // 公式格 C1 = A1 + B1（数值），EPPlus Calculate 生成缓存值 30
        ws.Cells[1, 3].Formula = "A1+B1";
        ws.Cells[1, 3].Calculate();

        // 自定义列宽（A 列 42，B 列 18）
        ws.Column(1).Width = 42;
        ws.Column(2).Width = 18;

        // 样式：A1 粗体 + 黄色填充
        ws.Cells[1, 1].Style.Font.Bold = true;
        ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

        // 折线图
        var chart = ws.Drawings.AddChart("chart1", eChartType.Line);
        chart.Series.Add(ws.Cells[1, 1, 1, 2], ws.Cells[2, 1, 2, 2]);

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    [Fact]
    public void Write_StripsChartsAndFormulas_PreservesStyleAndColumnWidth_AndWritesValues()
    {
        var template = MakeSyntheticTemplate();

        // 先确认模板确实有 chart + 公式（自证测试有效）
        using (var check = new ExcelPackage(new FileInfo(template)))
        {
            var ws = check.Workbook.Worksheets["Sheet1"];
            Assert.True(ws.Drawings.Count > 0, "模板应有 chart");
            Assert.False(string.IsNullOrEmpty(ws.Cells[1, 3].Formula), "模板 C1 应有公式");
            Assert.Equal(30d, Convert.ToDouble(ws.Cells[1, 3].Value)); // 缓存计算值
        }

        var outPath = Path.Combine(_dir, "out.xlsx");
        // 增量写：改 A1=10→99（脏格 (0,0)），其余不动
        var plan = new SheetWritePlan(
            SheetName: "Sheet1",
            Full: false,
            RowCount: 2,
            ColCount: 3,
            FullData: null,
            DirtyCells: [(0, 0, "99")]
        );

        ExcelWriteBack.Write(template, outPath, [plan]);

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];

        // chart 剥离
        Assert.Equal(0, sheet.Drawings.Count);

        // 公式剥离，但保留计算值 30（C1 不再是公式，值仍是 30）
        Assert.True(string.IsNullOrEmpty(sheet.Cells[1, 3].Formula), "C1 公式应被剥离");
        Assert.Equal(30d, Convert.ToDouble(sheet.Cells[1, 3].Value));

        // 列宽保留（浮点比较容差）
        Assert.True(
            Math.Abs(sheet.Column(1).Width - 42) < 0.5,
            $"A 列宽应≈42，实际 {sheet.Column(1).Width}"
        );
        Assert.True(
            Math.Abs(sheet.Column(2).Width - 18) < 0.5,
            $"B 列宽应≈18，实际 {sheet.Column(2).Width}"
        );

        // 样式保留：A1 粗体 + 黄填充
        Assert.True(sheet.Cells[1, 1].Style.Font.Bold, "A1 应保留粗体");
        Assert.Equal(ExcelFillStyle.Solid, sheet.Cells[1, 1].Style.Fill.PatternType);

        // 增量写生效：A1=99，其余值不变
        Assert.Equal("99", sheet.Cells[1, 1].Value?.ToString());
        Assert.Equal("20", sheet.Cells[1, 2].Value?.ToString());
        Assert.Equal("data", sheet.Cells[2, 2].Value?.ToString());

        _output.WriteLine(
            $"[WriteBack] chart={sheet.Drawings.Count}, C1.Formula='{sheet.Cells[1, 3].Formula}', "
                + $"C1.Value={sheet.Cells[1, 3].Value}, colA.W={sheet.Column(1).Width:F1}, A1='{sheet.Cells[1, 1].Value}'"
        );
    }

    [Fact]
    public void Write_DirtyOnly_DoesNotTouchOtherCells()
    {
        var template = MakeSyntheticTemplate();
        var outPath = Path.Combine(_dir, "dirty.xlsx");

        // 只标脏 B1（(0,1)）→ 改成 "changed"
        var plan = new SheetWritePlan("Sheet1", Full: false, 2, 3, null, [(0, 1, "changed")]);

        ExcelWriteBack.Write(template, outPath, [plan]);

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("changed", sheet.Cells[1, 2].Value?.ToString()); // B1 改了
        Assert.Equal("10", sheet.Cells[1, 1].Value?.ToString()); // A1 未动
        Assert.Equal("hdr", sheet.Cells[2, 1].Value?.ToString()); // A2 未动
    }

    [Fact]
    public void Write_Full_RewritesEntireSheet()
    {
        var template = MakeSyntheticTemplate();
        var outPath = Path.Combine(_dir, "full.xlsx");

        // 全量重写：2×2 新数据（比模板 2×3 少一列，触发 DeleteColumn）
        var data = new string[2, 2]
        {
            { "n00", "n01" },
            { "n10", "n11" },
        };
        var plan = new SheetWritePlan("Sheet1", Full: true, 2, 2, data, []);

        ExcelWriteBack.Write(template, outPath, [plan]);

        using var result = new ExcelPackage(new FileInfo(outPath));
        var sheet = result.Workbook.Worksheets["Sheet1"];
        Assert.Equal("n00", sheet.Cells[1, 1].Value?.ToString());
        Assert.Equal("n11", sheet.Cells[2, 2].Value?.ToString());
        Assert.Equal(0, sheet.Drawings.Count); // chart 仍被剥离
        // 全量写仍保留列宽（模板样式不因值重写丢失）
        Assert.True(Math.Abs(sheet.Column(1).Width - 42) < 0.5);
    }
}
