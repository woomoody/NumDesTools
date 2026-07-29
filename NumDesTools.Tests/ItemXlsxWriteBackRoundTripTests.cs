using System.Diagnostics;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// P4 WF2 真实 Item.xlsx 写回 round-trip（在 <b>临时副本</b> 上操作，绝不碰原文件）。
/// 复现生产保存路径的数据层：ColumnStoreExcelLoader.Load → 编辑标脏 → 组装 SheetWritePlan（增量）→
/// ExcelWriteBack.Write（原文件为模板）。断言改的格新值、未改格旧值、行列数不变；
/// 二次保存无脏格应秒过。同时先记录 Item.xlsx 的实际格式（图表/公式/列宽），供报告。
/// </summary>
public sealed class ItemXlsxWriteBackRoundTripTests : IDisposable
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";
    private readonly string _dir;
    private readonly ITestOutputHelper _output;

    static ItemXlsxWriteBackRoundTripTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    public ItemXlsxWriteBackRoundTripTests(ITestOutputHelper output)
    {
        _output = output;
        _dir = Path.Combine(
            Path.GetTempPath(),
            "xlsx-item-roundtrip",
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
        catch { }
    }

    [Fact]
    public void RecordItemXlsxActualFormat()
    {
        Assert.True(File.Exists(ItemPath), $"缺 {ItemPath}");
        using var pkg = new ExcelPackage(new FileInfo(ItemPath));
        var ws = pkg.Workbook.Worksheets[0];
        var dim = ws.Dimension;

        var chartCount = ws.Drawings.Count;
        // 抽样前 200×20 格看有无公式（全表扫太慢，抽样足够记录事实）
        var formulaCount = 0;
        var maxR = Math.Min(dim.End.Row, 200);
        var maxC = Math.Min(dim.End.Column, 20);
        for (var r = 1; r <= maxR; r++)
        for (var c = 1; c <= maxC; c++)
            if (!string.IsNullOrEmpty(ws.Cells[r, c].Formula))
                formulaCount++;

        // 列宽：看前 5 列是否非默认
        var widths = string.Join(
            ", ",
            Enumerable
                .Range(1, Math.Min(5, dim.End.Column))
                .Select(c => $"col{c}={ws.Column(c).Width:F1}")
        );

        _output.WriteLine(
            $"[Item.xlsx 实际格式] sheet='{ws.Name}', dim={dim.End.Row}×{dim.End.Column}, "
                + $"charts={chartCount}, formulas(采样{maxR}×{maxC})={formulaCount}, 列宽[{widths}], "
                + $"默认列宽={ws.DefaultColWidth:F1}"
        );

        Assert.True(dim.End.Row > 0);
    }

    [Fact]
    public void RoundTrip_IncrementalEdit_PreservesUnchanged_And_SecondSaveIsNoOp()
    {
        Assert.True(File.Exists(ItemPath), $"缺 {ItemPath}");

        // 1. 复制到临时副本（绝不碰原文件）
        var copy = Path.Combine(_dir, "Item_copy.xlsx");
        File.Copy(ItemPath, copy, overwrite: true);

        // 2. 加载进 ColumnStore（生产加载路径）
        var store = ColumnStoreExcelLoader.Load(copy);
        var rows = store.RowCount;
        var cols = store.ColumnCount;
        // 行/列数动态取自 EPPlus dimension（上游刷表会增行，硬编码基线脆断）；断言 ColumnStore 与文件形状一致。
        int expectedRows,
            expectedCols;
        using (var shape = new ExcelPackage(new FileInfo(copy)))
        {
            var dim = shape.Workbook.Worksheets[0].Dimension;
            expectedRows = dim.End.Row;
            expectedCols = dim.End.Column;
        }
        Assert.Equal(expectedRows, rows);
        Assert.Equal(expectedCols, cols);

        // 记录 3 个将改格的旧值 + 3 个不改格的旧值（0-based）
        var edit1 = (r: 4, c: 1); // Excel(5,2) 基线 "11010001"
        var edit2 = (r: 99, c: 1); // Excel(100,2) 基线 "13010504"
        var edit3 = (r: 4, c: 2); // Excel(5,3)
        var keepA = (r: 4, c: 3);
        var keepB = (r: 100, c: 1);
        var keepC = (r: 500, c: 5);
        var keepAOld = store.GetCell(keepA.r, keepA.c);
        var keepBOld = store.GetCell(keepB.r, keepB.c);
        var keepCOld = store.GetCell(keepC.r, keepC.c);

        // 3. 编辑 3 格（标脏）
        store.SetCell(edit1.r, edit1.c, "P4_EDIT_1");
        store.SetCell(edit2.r, edit2.c, "P4_EDIT_2");
        store.SetCell(edit3.r, edit3.c, "P4_EDIT_3");
        Assert.Equal(3, store.DirtyCells.Count);
        Assert.False(store.StructureChanged); // 纯编辑，无结构变更 → 增量写

        // 4. 组装增量计划 + 原子写（模板=copy 自身，输出=tmp，再 File.Replace）
        var plan = BuildIncrementalPlan("Sheet1", store);
        var sw1 = Stopwatch.StartNew();
        var res1 = AtomicFileWriter.Write(copy, tmp => ExcelWriteBack.Write(copy, tmp, [plan]));
        sw1.Stop();
        Assert.True(res1.Succeeded, res1.Error?.ToString());
        store.ClearDirty();

        // 5. 独立重开断言
        using (var check = new ExcelPackage(new FileInfo(copy)))
        {
            var ws = check.Workbook.Worksheets["Sheet1"];
            Assert.Equal(expectedRows, ws.Dimension.End.Row); // 行数不变
            Assert.Equal(expectedCols, ws.Dimension.End.Column); // 列数不变
            // 3 个改格新值（EPPlus 1-based）
            Assert.Equal("P4_EDIT_1", ws.Cells[edit1.r + 1, edit1.c + 1].Value?.ToString());
            Assert.Equal("P4_EDIT_2", ws.Cells[edit2.r + 1, edit2.c + 1].Value?.ToString());
            Assert.Equal("P4_EDIT_3", ws.Cells[edit3.r + 1, edit3.c + 1].Value?.ToString());
            // 3 个未改格旧值不变
            Assert.Equal(
                keepAOld ?? string.Empty,
                ws.Cells[keepA.r + 1, keepA.c + 1].Value?.ToString() ?? string.Empty
            );
            Assert.Equal(
                keepBOld ?? string.Empty,
                ws.Cells[keepB.r + 1, keepB.c + 1].Value?.ToString() ?? string.Empty
            );
            Assert.Equal(
                keepCOld ?? string.Empty,
                ws.Cells[keepC.r + 1, keepC.c + 1].Value?.ToString() ?? string.Empty
            );
            Assert.Equal(0, ws.Drawings.Count); // 图表剥离（Item.xlsx 若本无图表也应=0）
        }

        // 6. 二次保存（无新编辑 → 脏集合空 → 增量写 0 格），记录耗时对比
        var plan2 = BuildIncrementalPlan("Sheet1", store);
        Assert.Empty(plan2.DirtyCells); // 无脏格
        var sw2 = Stopwatch.StartNew();
        var res2 = AtomicFileWriter.Write(copy, tmp => ExcelWriteBack.Write(copy, tmp, [plan2]));
        sw2.Stop();
        Assert.True(res2.Succeeded);

        _output.WriteLine(
            $"[RoundTrip] 首次保存(3 脏格) {sw1.ElapsedMilliseconds} ms; "
                + $"二次保存(0 脏格) {sw2.ElapsedMilliseconds} ms"
        );
    }

    private static SheetWritePlan BuildIncrementalPlan(string sheet, ColumnStore store)
    {
        var dirty = store
            .DirtyCells.Select(cell =>
                (cell.Row, cell.Col, (string?)store.GetCell(cell.Row, cell.Col))
            )
            .ToList();
        return new SheetWritePlan(
            sheet,
            Full: store.StructureChanged,
            store.RowCount,
            store.ColumnCount,
            null,
            dirty
        );
    }
}
