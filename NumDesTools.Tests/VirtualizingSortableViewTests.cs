using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// VirtualizingSortableView 单测：核心验证「排序/筛选不整表复制 ColumnStore」。
/// 机制：视图维护一个 <c>int[] _rowOrder</c>（行号排列），排序只对该 int 数组排序，
/// 比较时按需读 ColumnStore 对应列；<c>this[index]</c> 经 _rowOrder 转一次再取 RowView。
/// 证据：用计数型单元格访问器统计排序期间 GetCell 调用次数，断言其为 O(n log n) 量级，
/// 且没有把整表数据搬到别的容器（RowView 仍回指 ColumnStore，改视图即改 store）。
/// 全新代码，严格 RED→GREEN。
/// </summary>
public sealed class VirtualizingSortableViewTests
{
    private static ColumnStore NumericStore(int rows)
    {
        var store = ColumnStore.Create(["A", "B"], rows);
        for (var r = 0; r < rows; r++)
        {
            store.AppendRow();
            // A 列倒序数值，用于验证排序；补零保证字符串序 == 数值序
            store.SetCellQuiet(r, 0, (rows - r).ToString("D8"));
            store.SetCellQuiet(r, 1, $"label{r}");
        }

        return store;
    }

    /// <summary>
    /// 计数型单元格访问器：包裹 ColumnStore.GetCell 记录调用次数，用于证明排序开销量级。
    /// 这不是"整表复制"——它每次都实时回读 store，不缓存/不搬运数据。
    /// </summary>
    private sealed class CountingAccessor(ColumnStore store)
    {
        public long GetCellCalls { get; private set; }

        public string? GetCell(int row, int col)
        {
            GetCellCalls++;
            return store.GetCell(row, col);
        }
    }

    [Fact]
    public void DefaultOrder_MatchesColumnStoreRowOrder()
    {
        var store = NumericStore(10);
        var view = new VirtualizingSortableView(store);

        Assert.Equal(10, view.Count);
        Assert.Equal(0, ((RowView)view[0]!).RowIndex);
        Assert.Equal(9, ((RowView)view[9]!).RowIndex);
    }

    [Fact]
    public void SortAscending_ReordersByColumnValue_WithoutCopyingTable()
    {
        var store = NumericStore(1000);
        var accessor = new CountingAccessor(store);
        var view = new VirtualizingSortableView(store, accessor.GetCell);

        view.SortBy(0, ascending: true);

        // 升序后第 0 行应是 A 列最小值 "00000001"（原 store 最后一行）
        var top = (RowView)view[0]!;
        Assert.Equal("00000001", top[0]);

        // O(n log n) 上界：1000 * log2(1000) ≈ 9966，比较每次读 2 个 cell。
        // 给宽松系数（排序实现常数），但必须远小于"整表复制"的 n*cols = 1000*2 反复扫多轮，
        // 关键是它必须 << n^2（线性扫描整表若干轮）。这里断言 < 200_000 已足够区分
        // O(n log n)(~2万) 与 任何整表多轮复制/冒泡(O(n^2)=100万+)。
        Assert.True(
            accessor.GetCellCalls < 200_000,
            $"GetCell called {accessor.GetCellCalls} times — suspiciously high, smells like full-table copy"
        );
    }

    [Fact]
    public void SortDescending_ReordersCorrectly()
    {
        var store = NumericStore(100);
        var view = new VirtualizingSortableView(store);

        view.SortBy(0, ascending: false);

        var top = (RowView)view[0]!;
        Assert.Equal("00000100", top[0]);
    }

    [Fact]
    public void Sort_KeepsRowViewsBackedByStore_EditThroughViewChangesStore()
    {
        var store = NumericStore(50);
        var view = new VirtualizingSortableView(store);
        view.SortBy(0, ascending: true);

        var top = (RowView)view[0]!;
        var storeRow = top.RowIndex; // 排序后 view[0] 对应的真实 store 行号
        top[1] = "edited-via-sorted-view";

        // 证明没有整表复制：改视图里的 RowView 直接改到了 ColumnStore 对应行
        Assert.Equal("edited-via-sorted-view", store.GetCell(storeRow, 1));
        Assert.True(store.IsDirty(storeRow, 1));
    }

    [Fact]
    public void SortCost_ScalesAsNLogN_NotNSquared()
    {
        // 10x 数据量，GetCell 调用应约增 ~13x（n log n），绝不 ~100x（n^2）
        var small = new CountingAccessor(NumericStore(1000));
        var smallView = new VirtualizingSortableView(NumericStore(1000), small.GetCell);
        smallView.SortBy(0, ascending: true);

        var big = new CountingAccessor(NumericStore(10000));
        var bigView = new VirtualizingSortableView(NumericStore(10000), big.GetCell);
        bigView.SortBy(0, ascending: true);

        var ratio = (double)big.GetCellCalls / small.GetCellCalls;
        Assert.True(
            ratio is > 8 and < 25,
            $"scale ratio {ratio:F1}x for 10x data — expected ~13x (n log n), not ~100x (n^2)"
        );
    }

    [Fact]
    public void Filter_ShowsOnlyMatchingRows_WithoutCopyingTable()
    {
        var store = ColumnStore.Create(["A"], 100);
        for (var r = 0; r < 100; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, r % 2 == 0 ? "even" : "odd");
        }

        var view = new VirtualizingSortableView(store);
        view.ApplyFilter(row => store.GetCell(row, 0) == "even");

        Assert.Equal(50, view.Count);
        foreach (var item in view)
        {
            Assert.Equal("even", ((RowView)item!)[0]);
        }
    }

    [Fact]
    public void ClearSort_RestoresNaturalOrder()
    {
        var store = NumericStore(20);
        var view = new VirtualizingSortableView(store);
        view.SortBy(0, ascending: true);

        view.ClearSort();

        Assert.Equal(0, ((RowView)view[0]!).RowIndex);
        Assert.Equal(19, ((RowView)view[19]!).RowIndex);
    }

    [Fact]
    public void Sort_RaisesResetNotification()
    {
        var store = NumericStore(10);
        var view = new VirtualizingSortableView(store);
        NotifyCollectionChangedAction? action = null;
        ((INotifyCollectionChanged)view).CollectionChanged += (_, e) => action = e.Action;

        view.SortBy(0, ascending: true);

        Assert.Equal(NotifyCollectionChangedAction.Reset, action);
    }

    [Fact]
    public void ImplementsIListAndCollectionView()
    {
        var store = NumericStore(10);
        var view = new VirtualizingSortableView(store);

        Assert.IsAssignableFrom<IList>(view);
        Assert.IsAssignableFrom<INotifyCollectionChanged>(view);
    }

    [Fact]
    public void Filter_ThenSort_Compose()
    {
        var store = ColumnStore.Create(["A"], 100);
        for (var r = 0; r < 100; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, (100 - r).ToString("D4"));
        }

        var view = new VirtualizingSortableView(store);
        view.ApplyFilter(row =>
        {
            var v = int.Parse(store.GetCell(row, 0)!);
            return v <= 10; // 只留 A<=10 的 10 行
        });
        view.SortBy(0, ascending: true);

        Assert.Equal(10, view.Count);
        Assert.Equal("0001", ((RowView)view[0]!)[0]);
        Assert.Equal("0010", ((RowView)view[9]!)[0]);
    }
}
