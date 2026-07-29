using System.ComponentModel;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// RowView 轻量代理单测：验证它只持有 ColumnStore 引用 + 行号，读写直接转发到 ColumnStore，
/// 不在内部缓存整行数据；且编辑经过 SetCell 写回并标脏、触发 INotifyPropertyChanged。
/// 全新代码，严格 RED→GREEN：本文件先于 RowView.cs 存在（编译失败即 RED）。
/// </summary>
public sealed class RowViewTests
{
    private static ColumnStore StoreWithRows(int rows)
    {
        var store = ColumnStore.Create(["A", "B", "C"], rows);
        for (var i = 0; i < rows; i++)
        {
            store.AppendRow();
        }

        return store;
    }

    [Fact]
    public void Indexer_ByColumnIndex_ForwardsToColumnStore()
    {
        var store = StoreWithRows(3);
        store.SetCellQuiet(1, 2, "hello");

        var view = new RowView(store, 1);

        Assert.Equal("hello", view[2]);
    }

    [Fact]
    public void Indexer_ByColumnName_ForwardsToColumnStore()
    {
        var store = StoreWithRows(3);
        store.SetCellQuiet(1, 1, "world");

        var view = new RowView(store, 1);

        Assert.Equal("world", view["B"]);
    }

    [Fact]
    public void Indexer_Set_WritesBackToColumnStore_AndMarksDirty()
    {
        var store = StoreWithRows(3);
        var view = new RowView(store, 2);

        view[0] = "edited";

        Assert.Equal("edited", store.GetCell(2, 0));
        Assert.True(store.IsDirty(2, 0));
    }

    [Fact]
    public void Indexer_SetByName_WritesBackToColumnStore_AndMarksDirty()
    {
        var store = StoreWithRows(3);
        var view = new RowView(store, 0);

        view["C"] = "edited-by-name";

        Assert.Equal("edited-by-name", store.GetCell(0, 2));
        Assert.True(store.IsDirty(0, 2));
    }

    [Fact]
    public void Indexer_Set_RaisesPropertyChanged_ForColumnName()
    {
        var store = StoreWithRows(1);
        var view = new RowView(store, 0);
        var raised = new List<string?>();
        ((INotifyPropertyChanged)view).PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        view[1] = "x";

        // WPF DataGrid 绑定 "[B]" 形式，改列后需通知 "Item[]" 让绑定刷新
        Assert.Contains("Item[]", raised);
    }

    [Fact]
    public void RowIndex_IsExposed()
    {
        var store = StoreWithRows(5);

        var view = new RowView(store, 3);

        Assert.Equal(3, view.RowIndex);
    }

    [Fact]
    public void TwoViews_SameRow_ShareUnderlyingStore_NoLocalCopy()
    {
        var store = StoreWithRows(2);
        var a = new RowView(store, 0);
        var b = new RowView(store, 0);

        a[0] = "changed-through-a";

        // b 没有本地缓存，读的是同一份 ColumnStore，立即看到 a 的写入
        Assert.Equal("changed-through-a", b[0]);
    }

    [Fact]
    public void GetCell_NullStaysNull()
    {
        var store = StoreWithRows(1);

        var view = new RowView(store, 0);

        Assert.Null(view[0]);
        Assert.Null(view["A"]);
    }

    [Fact]
    public void UnknownColumnName_Throws()
    {
        var store = StoreWithRows(1);
        var view = new RowView(store, 0);

        Assert.Throws<ArgumentException>(() => _ = view["ZZZ"]);
    }
}
