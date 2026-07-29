using System.Collections;
using System.Collections.Specialized;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// VirtualizingCollection 单测：验证它作为 <see cref="IList"/> 直接覆盖 ColumnStore，
/// Count 正确、随机访问返回正确行、RowView 按需生成（构造集合本身不物化任何行），
/// 并实现 INotifyCollectionChanged 供 WPF 绑定。
/// 全新代码，严格 RED→GREEN：本文件先于 VirtualizingCollection.cs 存在（编译失败即 RED）。
/// </summary>
public sealed class VirtualizingCollectionTests
{
    private static ColumnStore StoreWith(int rows, int cols = 3)
    {
        var names = new string[cols];
        for (var c = 0; c < cols; c++)
        {
            names[c] = ((char)('A' + c)).ToString();
        }

        var store = ColumnStore.Create(names, rows);
        for (var r = 0; r < rows; r++)
        {
            store.AppendRow();
            store.SetCellQuiet(r, 0, $"r{r}");
        }

        return store;
    }

    [Fact]
    public void Count_MatchesColumnStoreRowCount()
    {
        var store = StoreWith(5000);

        var collection = new VirtualizingCollection(store);

        Assert.Equal(5000, collection.Count);
    }

    [Fact]
    public void Indexer_ReturnsRowView_ForCorrectRow()
    {
        var store = StoreWith(100);

        var collection = new VirtualizingCollection(store);
        var item = collection[42];

        var view = Assert.IsType<RowView>(item);
        Assert.Equal(42, view.RowIndex);
        Assert.Equal("r42", view[0]);
    }

    [Fact]
    public void Construction_DoesNotMaterializeAnyRowView()
    {
        var store = StoreWith(60000);

        var collection = new VirtualizingCollection(store);

        // 构造集合不该预先生成 6 万个 RowView（那就不是虚拟化了）
        Assert.Equal(0, collection.MaterializedRowViewCount);
    }

    [Fact]
    public void RandomAccess_MaterializesOnlyTouchedRows()
    {
        var store = StoreWith(60000);
        var collection = new VirtualizingCollection(store);

        _ = collection[0];
        _ = collection[100];
        _ = collection[59999];

        // 只访问了 3 行，物化的 RowView 数量应远小于总行数
        Assert.True(
            collection.MaterializedRowViewCount <= 3,
            $"materialized {collection.MaterializedRowViewCount} row views for 3 accesses"
        );
    }

    [Fact]
    public void Indexer_SameIndexTwice_ReturnsCachedInstance()
    {
        var store = StoreWith(100);
        var collection = new VirtualizingCollection(store);

        var first = collection[10];
        var second = collection[10];

        Assert.Same(first, second);
        Assert.Equal(1, collection.MaterializedRowViewCount);
    }

    [Fact]
    public void Enumeration_YieldsAllRowsInOrder()
    {
        var store = StoreWith(50);
        var collection = new VirtualizingCollection(store);

        var seen = new List<int>();
        foreach (var item in collection)
        {
            var view = Assert.IsType<RowView>(item);
            seen.Add(view.RowIndex);
        }

        Assert.Equal(Enumerable.Range(0, 50), seen);
    }

    [Fact]
    public void IndexOf_ReturnsRowIndex_ForContainedRowView()
    {
        var store = StoreWith(100);
        var collection = new VirtualizingCollection(store);
        var view = collection[37];

        var index = collection.IndexOf(view);

        Assert.Equal(37, index);
    }

    [Fact]
    public void IsFixedSize_And_ReadOnly_ReportedForWpf()
    {
        var store = StoreWith(10);
        var collection = new VirtualizingCollection(store);

        // DataGrid 通过 IList 探测能力；本集合按行虚拟化，不支持任意 Add/Insert
        Assert.True(collection.IsFixedSize);
        Assert.True(collection.IsReadOnly);
    }

    [Fact]
    public void ImplementsNotifyCollectionChanged()
    {
        var store = StoreWith(10);

        var collection = new VirtualizingCollection(store);

        Assert.IsAssignableFrom<INotifyCollectionChanged>(collection);
    }

    [Fact]
    public void Refresh_RaisesResetNotification()
    {
        var store = StoreWith(10);
        var collection = new VirtualizingCollection(store);
        NotifyCollectionChangedAction? action = null;
        ((INotifyCollectionChanged)collection).CollectionChanged += (_, e) => action = e.Action;

        collection.Refresh();

        Assert.Equal(NotifyCollectionChangedAction.Reset, action);
    }

    [Fact]
    public void Indexer_OutOfRange_Throws()
    {
        var store = StoreWith(5);
        var collection = new VirtualizingCollection(store);

        Assert.Throws<ArgumentOutOfRangeException>(() => _ = collection[5]);
    }
}
