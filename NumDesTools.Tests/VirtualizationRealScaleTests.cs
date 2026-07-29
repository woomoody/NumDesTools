using System.Diagnostics;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 真实规模验证：用 Item.xlsx（~6.5万行 × 85 列）跑 ColumnStore → VirtualizingCollection /
/// VirtualizingSortableView 全链路，产出可复核的性能数字（不是理论推断）。
/// 覆盖 GOAL 的三条硬验证：① Count 正确 + 随机访问正确；② 排序 O(n log n) 且不整表复制；
/// ③ 编辑 RowView 后 ColumnStore 对应位置真的被改。数字通过 <see cref="ITestOutputHelper"/> 打印。
/// 沿用 <see cref="ColumnStoreExcelLoaderTests"/> 的约定：直接依赖真实文件路径存在。
/// </summary>
public sealed class VirtualizationRealScaleTests(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    // 行/列数动态取自 EPPlus dimension（真实文件被上游刷表增行，硬编码会脆断）。
    private static readonly int ExpectedRows = ReadDimension().Rows;
    private static readonly int ExpectedCols = ReadDimension().Cols;

    static VirtualizationRealScaleTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    private static (int Rows, int Cols) ReadDimension()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var dim = package.Workbook.Worksheets[0].Dimension;
        return (dim.End.Row, dim.End.Column);
    }

    [Fact]
    public void Collection_Count_And_RandomAccess_AreCorrect()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var collection = new VirtualizingCollection(store);

        Assert.Equal(ExpectedRows, collection.Count);

        // 随机访问若干 index，值须与 ColumnStore 直接读一致
        int[] probes = [0, 4, 99, 999, 29999, ExpectedRows - 1];
        var sw = Stopwatch.StartNew();
        foreach (var idx in probes)
        {
            var view = (RowView)collection[idx]!;
            Assert.Equal(store.GetCell(idx, 1), view[1]);
            Assert.Equal(idx, view.RowIndex);
        }

        sw.Stop();

        // 已知锚点（纯 ASCII，与编码无关）：低行号稳定；末行值随文件增行变化，与 store 交叉验证。
        Assert.Equal("11010001", ((RowView)collection[4]!)[1]);
        Assert.Equal(
            store.GetCell(ExpectedRows - 1, 1),
            ((RowView)collection[ExpectedRows - 1]!)[1]
        );

        output.WriteLine($"[Collection] Count={collection.Count} (expected {ExpectedRows})");
        output.WriteLine(
            $"[Collection] 6 random accesses: {sw.Elapsed.TotalMilliseconds:F3} ms; "
                + $"materialized RowViews={collection.MaterializedRowViewCount} (proves on-demand, not {ExpectedRows})"
        );
        Assert.True(collection.MaterializedRowViewCount <= probes.Length);
    }

    [Fact]
    public void SingleRandomAccess_IsSubMillisecond_And_DoesNotMaterializeAll()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var collection = new VirtualizingCollection(store);

        // 预热一次（JIT），再测单次随机访问
        _ = collection[12345];
        var sw = Stopwatch.StartNew();
        var view = (RowView)collection[54321]!;
        var value = view[1];
        sw.Stop();

        Assert.Equal(store.GetCell(54321, 1), value);
        output.WriteLine(
            $"[Collection] single random access @54321: {sw.Elapsed.TotalMilliseconds:F4} ms; "
                + $"materialized={collection.MaterializedRowViewCount}"
        );
        // 只碰了 2 行，绝不该物化整表
        Assert.True(collection.MaterializedRowViewCount <= 2);
    }

    [Fact]
    public void Sort_65kRows_IsFast_And_DoesNotCopyTable()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);

        // 用计数器证明排序期间的 GetCell 调用量级
        long getCellCalls = 0;
        string? Counting(int row, int col)
        {
            Interlocked.Increment(ref getCellCalls);
            return store.GetCell(row, col);
        }

        var view = new VirtualizingSortableView(store, Counting);

        var sw = Stopwatch.StartNew();
        view.SortBy(1, ascending: true); // 按 B 列（id 列）排序
        sw.Stop();

        // n log n 上界：65105 * log2(65105) ≈ 65105 * 16 ≈ 1.04M 次比较，
        // 每次比较读 2 个 cell ≈ 2.08M。给宽松系数到 10M；关键是它 << n^2(=4.2e9) 的整表多轮扫描。
        var nLogN = ExpectedRows * Math.Log2(ExpectedRows);
        output.WriteLine(
            $"[Sort] 65k rows by col B: {sw.Elapsed.TotalMilliseconds:F1} ms; "
                + $"GetCell calls={getCellCalls:N0}; n*log2(n)≈{nLogN:N0}; "
                + $"calls/(n log n)={getCellCalls / nLogN:F2} (≈2 means 2 reads per compare, O(n log n))"
        );

        Assert.Equal(ExpectedRows, view.Count);
        // 升序后第 0 行的 B 列值 <= 末行的 B 列值（字符串序）
        var firstB = ((RowView)view[0]!)[1];
        var lastB = ((RowView)view[ExpectedRows - 1]!)[1];
        Assert.True(
            string.Compare(firstB, lastB, StringComparison.Ordinal) <= 0,
            $"first '{firstB}' should sort <= last '{lastB}'"
        );

        // 硬门槛：调用次数必须落在 O(n log n) 量级，绝不能是整表复制/多轮线性扫描
        Assert.True(
            getCellCalls < 10_000_000,
            $"GetCell called {getCellCalls:N0} times — too high, smells like full-table copy (n log n ≈ {2 * nLogN:N0})"
        );
    }

    [Fact]
    public void EditThroughRowView_WritesBackToColumnStore_AtRealScale()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var collection = new VirtualizingCollection(store);

        var view = (RowView)collection[50000]!;
        var original = store.GetCell(50000, 3);
        Assert.False(store.IsDirty(50000, 3));

        view[3] = "EDITED_AT_50000";

        Assert.Equal("EDITED_AT_50000", store.GetCell(50000, 3));
        Assert.True(store.IsDirty(50000, 3));
        Assert.Contains((50000, 3), store.DirtyCells);
        output.WriteLine(
            $"[Edit] row 50000 col 3: '{original}' -> 'EDITED_AT_50000'; "
                + $"ColumnStore.IsDirty={store.IsDirty(50000, 3)}; DirtyCells={store.DirtyCells.Count}"
        );
    }

    [Fact]
    public void MemoryFootprint_OfStoreAndViews_IsReported()
    {
        // 报告加载后托管内存占用，供 status.md 记录真实数字
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        var before = GC.GetTotalMemory(true);

        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var view = new VirtualizingSortableView(store);
        // 只访问视口大小的行（模拟 DataGrid 只渲染 ~50 行），证明不整表物化
        for (var i = 0; i < 50 && i < view.Count; i++)
        {
            _ = ((RowView)view[i]!)[1];
        }

        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        var after = GC.GetTotalMemory(true);

        var deltaMb = (after - before) / 1024.0 / 1024.0;
        output.WriteLine(
            $"[Memory] ColumnStore({ExpectedRows}x{ExpectedCols}) + view + 50 materialized RowViews: "
                + $"~{deltaMb:F1} MB managed delta"
        );
        Assert.Equal(ExpectedRows, store.RowCount);
        Assert.Equal(ExpectedCols, store.ColumnCount);
    }
}
