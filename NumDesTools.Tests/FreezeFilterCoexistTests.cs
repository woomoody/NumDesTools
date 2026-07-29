using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// #1 冻结+筛选共存：冻结行时主区只显示"第 N 行之后"的数据行（基础谓词 row&gt;=n），
/// 且列头筛选照常作用于主区数据行（与基础谓词 AND）。冻结区（另一 grid）不参与。
/// 这里锁定 ColumnStoreFilterAdapter 组合谓词后喂给 VirtualizingSortableView.ApplyFilter 的可观察结果
/// （复现主 grid 数据路径；DataGrid/DataGridColumn 需 STA，改测底层组合语义，等价覆盖）。
/// </summary>
public sealed class FreezeFilterCoexistTests
{
    private static ColumnStore MakeStore()
    {
        // 10 行 × 2 列。第 0 列 = 行号字符串；第 1 列 = "kind" 分类
        var store = ColumnStore.Create(["A", "B"], 10);
        for (var r = 0; r < 10; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, r.ToString());
            store.SetCellQuiet(r, 1, r % 2 == 0 ? "even" : "odd");
        }
        store.ClearDirty();
        return store;
    }

    [Fact]
    public void BasePredicateOnly_ShowsOnlyDataRows_FrozenRegionExcluded()
    {
        var store = MakeStore();
        var view = new VirtualizingSortableView(store);

        // 冻结前 3 行（表头区）→ 主区基础谓词 row>=3 → 只剩 7 行数据
        var n = 3;
        view.ApplyFilter(row => row >= n);

        Assert.Equal(7, view.Count);
        Assert.Equal(3, ((RowView)view[0]!).RowIndex); // 主区第一行 = store 行 3
        Assert.Equal(9, ((RowView)view[6]!).RowIndex);
    }

    [Fact]
    public void BasePredicate_AND_ColumnFilter_FiltersDataRowsOnly()
    {
        var store = MakeStore();
        var view = new VirtualizingSortableView(store);
        var n = 3;

        // 组合：row>=3（冻结区排除） AND B 列包含 "even"
        Func<int, bool> basePred = row => row >= n;
        var colPred = ColumnFilterPredicate.Build(store, [(1, "even", ColumnType.Text)]);
        view.ApplyFilter(row => basePred(row) && colPred(row));

        // 数据行区 (3..9) 中 even 的是 4,6,8 → 3 行（行 0,2 虽也是 even 但在冻结区，被 base 排除）
        Assert.Equal(3, view.Count);
        Assert.Equal(4, ((RowView)view[0]!).RowIndex);
        Assert.Equal(6, ((RowView)view[1]!).RowIndex);
        Assert.Equal(8, ((RowView)view[2]!).RowIndex);
        // 不整表物化：只访问了 3 行，物化数应 ≤ 3（远小于 store 的 10 行）
        Assert.True(view.MaterializedRowViewCount <= 3);
    }

    [Fact]
    public void ClearBasePredicate_RestoresAllRows()
    {
        var store = MakeStore();
        var view = new VirtualizingSortableView(store);
        view.ApplyFilter(row => row >= 3); // 冻结态
        Assert.Equal(7, view.Count);

        view.ClearFilter(); // 取消冻结 → 恢复全部
        Assert.Equal(10, view.Count);
        Assert.Equal(0, ((RowView)view[0]!).RowIndex);
    }
}
