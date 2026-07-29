using System.Diagnostics;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// P5 筛选执行下沉 ColumnStore 的集成测试 + "不整表物化" 证据。
/// <para>
/// 复现 <see cref="ColumnStoreFilterAdapter.OnFilterChanged"/> 的实际数据路径：
/// <c>ColumnFilterPredicate.Build(store, filters)</c> → <c>VirtualizingSortableView.ApplyFilter(predicate)</c>。
/// 这正是 DataGridExtensions 通过 ICustomFilter 把筛选交给我们后走的路（DataGrid/DataGridColumn 需 STA，
/// 无法在单测线程直接构造；此处测其调用的纯数据路径，等价覆盖"谓词下沉 + 不整表物化"）。
/// </para>
/// <para>
/// 关键断言：在真实规模 ColumnStore 上筛选，<see cref="VirtualizingSortableView.MaterializedRowViewCount"/>
/// 保持小量级（不因筛选而物化整表 RowView）。Item.xlsx 相关断言<b>从加载的 store 动态推导期望值</b>
/// （不硬编码行数——真实文件会被外部修改，硬编码基线会脆断）。
/// </para>
/// </summary>
public sealed class ColumnStoreFilterIntegrationTests(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    static ColumnStoreFilterIntegrationTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    [Fact]
    public void ApplyFilter_OverLargeColumnStore_DoesNotMaterializeAllRows()
    {
        // 合成 65000 行 × 3 列（贴近真实大表规模），只在少数行放目标值
        const int rows = 65000;
        var store = ColumnStore.Create(["A", "B", "C"], rows);
        for (var r = 0; r < rows; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, r.ToString());
            store.SetCellQuiet(r, 1, r % 1000 == 0 ? "TARGET" : "other");
            store.SetCellQuiet(r, 2, "x");
        }
        store.ClearDirty();

        var view = new VirtualizingSortableView(store);
        Assert.Equal(rows, view.Count);
        Assert.Equal(0, view.MaterializedRowViewCount); // 构造不物化

        // 走 adapter 内部同一路径：Build 谓词（B 列 contains "TARGET"）→ ApplyFilter
        var predicate = ColumnFilterPredicate.Build(store, [(1, "TARGET", ColumnType.Text)]);
        view.ApplyFilter(predicate);

        Assert.Equal(65, view.Count); // 每 1000 行一个 → 65 行（0,1000,...,64000）

        // ★ 核心证据：筛选后物化的 RowView 数远小于 65000（ApplyFilter 只重排 int[]，不物化任何行）
        Assert.Equal(0, view.MaterializedRowViewCount);
        output.WriteLine(
            $"[no-materialization] rows={rows}, filtered Count={view.Count}, "
                + $"MaterializedRowViewCount 筛选后={view.MaterializedRowViewCount}（应=0，远小于 {rows}）"
        );

        // 只访问视口内前 20 行 → 只物化 20 个（虚拟化正确）
        for (var i = 0; i < Math.Min(20, view.Count); i++)
            _ = view[i];
        Assert.True(
            view.MaterializedRowViewCount <= 20,
            $"访问 20 行后应最多物化 20 个 RowView，实际 {view.MaterializedRowViewCount}"
        );

        // 清筛选恢复全部行，仍不整表物化
        view.ClearFilter();
        Assert.Equal(rows, view.Count);
        Assert.Equal(0, view.MaterializedRowViewCount);
    }

    [Fact]
    public void RealItemXlsx_TextFilter_IsExactAndVirtualized()
    {
        Assert.True(File.Exists(ItemPath), $"缺 {ItemPath}");

        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var total = store.RowCount;
        var view = new VirtualizingSortableView(store);
        Assert.Equal(total, view.Count);

        // 动态推导期望：直接扫 B 列（col 1）统计 contains "11010001" 的行数（不硬编码，文件会变）。
        var expected = 0;
        for (var r = 0; r < total; r++)
            if (
                (store.GetCell(r, 1) ?? string.Empty).Contains(
                    "11010001",
                    StringComparison.OrdinalIgnoreCase
                )
            )
                expected++;

        var predicate = ColumnFilterPredicate.Build(store, [(1, "11010001", ColumnType.Text)]);
        var sw = Stopwatch.StartNew();
        view.ApplyFilter(predicate);
        sw.Stop();

        output.WriteLine(
            $"[Item.xlsx] 总行 {total}；B 列(col1) contains '11010001' → 期望 {expected} 行，实测 {view.Count} 行；"
                + $"筛选耗时 {sw.ElapsedMilliseconds} ms；筛选后物化 RowView={view.MaterializedRowViewCount}"
        );

        Assert.Equal(expected, view.Count); // 谓词命中数 = 独立扫描数（正确性）
        Assert.True(expected >= 1, "B 列应至少含 1 个 '11010001'（id 行）");
        Assert.Equal(0, view.MaterializedRowViewCount); // 大表筛选不整表物化

        // 清筛选恢复全部
        view.ClearFilter();
        Assert.Equal(total, view.Count);
        Assert.Equal(0, view.MaterializedRowViewCount);
    }

    [Fact]
    public void RealItemXlsx_NumericRangeFilter_MatchesIndependentScan()
    {
        Assert.True(File.Exists(ItemPath), $"缺 {ItemPath}");
        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var total = store.RowCount;
        var view = new VirtualizingSortableView(store);

        // 动态推导：id 列（col 1）中能解析为 double 且 >=11010001 的行数
        var expected = 0;
        for (var r = 0; r < total; r++)
        {
            var cell = store.GetCell(r, 1) ?? string.Empty;
            if (
                double.TryParse(
                    cell,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var v
                )
                && v >= 11010001
            )
                expected++;
        }

        var predicate = ColumnFilterPredicate.Build(store, [(1, ">=11010001", ColumnType.Integer)]);
        view.ApplyFilter(predicate);

        output.WriteLine(
            $"[Item.xlsx] 总行 {total}；id 列 >=11010001 → 期望 {expected} 行，实测 {view.Count} 行；"
                + $"物化 RowView={view.MaterializedRowViewCount}"
        );
        Assert.Equal(expected, view.Count);
        Assert.Equal(0, view.MaterializedRowViewCount);
    }
}
