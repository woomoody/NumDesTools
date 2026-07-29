using System.Collections;
using System.Collections.Specialized;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 直接覆盖 <see cref="ColumnStore"/> 的虚拟化行集合。因为 ColumnStore 已是内存里的全量列式数据，
/// 这里的"虚拟化"含义是<b>不为每行预先创建 .NET 对象树/绑定路径</b>：<see cref="this[int]"/> 在被访问时
/// 才按需生成轻量 <see cref="RowView"/>（并缓存复用），构造集合本身不物化任何行。
/// 实现非泛型 <see cref="IList"/>（WPF DataGrid 靠它正确虚拟化）+ <see cref="INotifyCollectionChanged"/>。
/// 集合大小随 ColumnStore 固定（按行虚拟化，不支持任意 Add/Insert），故 <see cref="IsFixedSize"/> /
/// <see cref="IsReadOnly"/> 均为 true。
/// </summary>
public sealed class VirtualizingCollection : IList, INotifyCollectionChanged
{
    private readonly ColumnStore _store;
    private readonly Dictionary<int, RowView> _cache = [];

    public VirtualizingCollection(ColumnStore store) =>
        _store = store ?? throw new ArgumentNullException(nameof(store));

    public event NotifyCollectionChangedEventHandler? CollectionChanged;

    /// <summary>诊断用：已按需物化的 RowView 数量。虚拟化正确时它远小于 <see cref="Count"/>。</summary>
    public int MaterializedRowViewCount => _cache.Count;

    public int Count => _store.RowCount;

    public bool IsFixedSize => true;

    public bool IsReadOnly => true;

    public bool IsSynchronized => false;

    public object SyncRoot { get; } = new();

    public object? this[int index]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _store.RowCount);
            return GetOrCreate(index);
        }
        set => throw new NotSupportedException("VirtualizingCollection is virtualized by row.");
    }

    /// <summary>丢弃已物化的 RowView 缓存并广播 Reset，让绑定重新拉取（结构性变更后调用）。</summary>
    public void Refresh()
    {
        _cache.Clear();
        CollectionChanged?.Invoke(
            this,
            new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset)
        );
    }

    public int IndexOf(object? value) => value is RowView view ? view.RowIndex : -1;

    public bool Contains(object? value) =>
        value is RowView view && view.RowIndex >= 0 && view.RowIndex < _store.RowCount;

    public IEnumerator GetEnumerator()
    {
        for (var row = 0; row < _store.RowCount; row++)
        {
            yield return GetOrCreate(row);
        }
    }

    public void CopyTo(Array array, int index)
    {
        ArgumentNullException.ThrowIfNull(array);
        for (var row = 0; row < _store.RowCount; row++)
        {
            array.SetValue(GetOrCreate(row), index + row);
        }
    }

    public int Add(object? value) =>
        throw new NotSupportedException("VirtualizingCollection is virtualized by row.");

    public void Insert(int index, object? value) =>
        throw new NotSupportedException("VirtualizingCollection is virtualized by row.");

    public void Remove(object? value) =>
        throw new NotSupportedException("VirtualizingCollection is virtualized by row.");

    public void RemoveAt(int index) =>
        throw new NotSupportedException("VirtualizingCollection is virtualized by row.");

    public void Clear() =>
        throw new NotSupportedException("VirtualizingCollection is virtualized by row.");

    private RowView GetOrCreate(int row)
    {
        if (_cache.TryGetValue(row, out var existing))
        {
            return existing;
        }

        var view = new RowView(_store, row);
        _cache[row] = view;
        return view;
    }
}
