using System.Diagnostics;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 补全 A/B/C 全链路实测缺的一环：写回。之前的 TuiMigrationBenchmarkC/B 只测了「读入/排序/整列复制」，
/// 没有任何方案测过「改N格→写回已有文件保留结构」——这正是"读和写都基于已有文件"这个场景最重的一步
/// （参见 reference_xlsx_readwrite_benchmarks.md：EPPlus SaveAs ~4868ms 跟改动格数无关）。
/// 本测试只测方案C（当前C#，复用现有 ExcelWriteBack），只读不改原文件（在临时副本上操作）。
/// </summary>
public sealed class TuiMigrationBenchmarkWriteback(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";
    private const int DirtyColumn = 1; // 对齐方案A/B/C的 SORT_COL（B 列/id 列，0-based）
    private const int DirtyRowCount = 1000; // 模拟一次编辑会话改动 1000 个散布的格

    [Fact]
    public void ReportTimings_WriteDirtyCellsBack()
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "wf-writeback-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var templatePath = Path.Combine(tempDir, "Item_copy.xlsx");
        File.Copy(ItemPath, templatePath); // 绝不碰原文件，全程只操作临时副本
        var outputPath = Path.Combine(tempDir, "Item_out.xlsx");

        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
            string sheetName;
            using (var package = new ExcelPackage(new FileInfo(templatePath)))
                sheetName = package.Workbook.Worksheets[0].Name;

            var store = ColumnStoreExcelLoader.Load(templatePath);
            var dirtyCells = new List<(int Row, int Col, string? Value)>(DirtyRowCount);
            for (var r = 0; r < DirtyRowCount && r < store.RowCount; r++)
                dirtyCells.Add((r, DirtyColumn, "888888"));

            var plan = new SheetWritePlan(sheetName, Full: false, store.RowCount, store.ColumnCount, null, dirtyCells);

            var writeMs = Median(
                Times(
                    3,
                    () => ExcelWriteBack.Write(templatePath, outputPath, [plan])
                )
            );

            // 验证：重新打开确认第一个改动的值真的写进去了
            using (var verify = new ExcelPackage(new FileInfo(outputPath)))
            {
                var cellValue = verify.Workbook.Worksheets[sheetName].Cells[1, DirtyColumn + 1].Value?.ToString();
                Assert.Equal("888888", cellValue);
            }

            output.WriteLine(
                $"[方案C 纯C# 写回] rows={store.RowCount} cols={store.ColumnCount} dirtyCells={DirtyRowCount} | "
                    + $"EPPlus写回(median of 3, 模板→输出全量SaveAs)={writeMs:F0}ms"
            );
        }
        finally
        {
            Directory.Delete(tempDir, recursive: true);
        }
    }

    private static IEnumerable<double> Times(int runs, Action action)
    {
        for (var i = 0; i < runs; i++)
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();
            yield return sw.Elapsed.TotalMilliseconds;
        }
    }

    private static double Median(IEnumerable<double> values)
    {
        var xs = values.ToList();
        xs.Sort();
        return xs[xs.Count / 2];
    }
}
