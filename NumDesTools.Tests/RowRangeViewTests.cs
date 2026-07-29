using System.Collections;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// P4 WF3 RowRangeView 单测：覆盖 ColumnStore 一个连续行区间，RowView 用真实 store 行号，
/// 编辑经 RowView 写回正确行；按行虚拟化不整表物化。冻结行方案的核心构件。
/// </summary>
public sealed class RowRangeViewTests
{
    private static ColumnStore Make(int rows)
    {
        var store = ColumnStore.Create(["A", "B"], rows);
        for (var r = 0; r < rows; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, $"r{r}");
        }
        store.ClearDirty();
        return store;
    }

    [Fact]
    public void ExposesOnlyRangeRows()
    {
        var store = Make(10);
        var view = new RowRangeView(store, start: 4, count: 3); // rows 4,5,6

        Assert.Equal(3, view.Count);
        Assert.Equal("r4", ((RowView)view[0]!)[0]);
        Assert.Equal("r5", ((RowView)view[1]!)[0]);
        Assert.Equal("r6", ((RowView)view[2]!)[0]);
    }

    [Fact]
    public void RowView_UsesRealStoreRowIndex_ForWriteback()
    {
        var store = Make(10);
        var view = new RowRangeView(store, start: 4, count: 3);

        // 视图局部索引 0 → store 行号 4
        var rv = (RowView)view[0]!;
        Assert.Equal(4, rv.RowIndex);

        rv[1] = "edited"; // 写 store(4,1)
        Assert.Equal("edited", store.GetCell(4, 1));
        Assert.True(store.IsDirty(4, 1));
    }

    [Fact]
    public void TopRange_And_BottomRange_CoverAllRows_NoOverlap()
    {
        var store = Make(10);
        var frozen = new RowRangeView(store, 0, 4); // 冻结前 4 行
        var main = new RowRangeView(store, 4, store.RowCount - 4); // 其余

        Assert.Equal(4, frozen.Count);
        Assert.Equal(6, main.Count);
        Assert.Equal(0, ((RowView)frozen[0]!).RowIndex);
        Assert.Equal(4, ((RowView)main[0]!).RowIndex); // 主区第一行 = store 行 4
        Assert.Equal(9, ((RowView)main[main.Count - 1]!).RowIndex);
    }

    [Fact]
    public void ClampsCountToStoreBounds()
    {
        var store = Make(5);
        var view = new RowRangeView(store, start: 3, count: 100); // 超界

        Assert.Equal(2, view.Count); // 只有 row 3,4
    }

    [Fact]
    public void SetRange_ResetsAndReclamps()
    {
        var store = Make(10);
        var view = new RowRangeView(store, 0, 4);
        var resets = 0;
        view.CollectionChanged += (_, e) =>
        {
            if (e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Reset)
                resets++;
        };

        view.SetRange(4, 6);

        Assert.Equal(1, resets);
        Assert.Equal(6, view.Count);
        Assert.Equal(4, view.Start);
        Assert.Equal("r4", ((RowView)view[0]!)[0]);
    }

    [Fact]
    public void Virtualized_DoesNotMaterializeAllRows()
    {
        var store = Make(1000);
        var view = new RowRangeView(store, 0, 1000);

        _ = view[0];
        _ = view[500];

        Assert.Equal(2, view.MaterializedRowViewCount); // 只物化访问过的 2 行
        Assert.Equal(1000, view.Count);
    }

    [Fact]
    public void IsList_ForDataGridVirtualization()
    {
        var store = Make(3);
        var view = new RowRangeView(store, 0, 3);

        Assert.IsAssignableFrom<IList>(view);
        Assert.True(view.IsFixedSize);
        Assert.True(view.IsReadOnly);
    }
}
