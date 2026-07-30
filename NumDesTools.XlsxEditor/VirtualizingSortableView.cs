using System.Collections;
using System.Collections.Specialized;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 可排序/可筛选的虚拟化视图，覆盖 <see cref="ColumnStore"/>。
/// <para>
/// 核心设计——<b>排序/筛选不整表复制</b>：视图只维护一个 <c>int[] _rowOrder</c>（ColumnStore 行号的排列/子集）。
/// 排序时仅对这个 int 数组排序，比较器按需读 ColumnStore 对应列的值（<see cref="_cellAccessor"/>）；
/// <see cref="this[int]"/> 通过 <c>_rowOrder[index]</c> 转一次再生成回指 store 的 <see cref="RowView"/>。
/// 因此排序开销是 O(n log n) 次单元格读取（就地重排 int），而不是把 n×列 的数据从 ColumnStore 搬到别处。
/// 筛选同理：<c>_rowOrder</c> 变成"可见行号"的子集，被隐藏的行仍原地留在 ColumnStore，不复制不删除。
/// </para>
/// <para>
/// 与 WPF 集成：实现非泛型 <see cref="IList"/> + <see cref="INotifyCollectionChanged"/>，可直接作为
/// DataGrid 的 ItemsSource。相比子类化 <see cref="System.Windows.Data.ListCollectionView"/> 重写
/// RefreshOverride，本视图把排序键的产生完全下沉到 ColumnStore 列级读取，天然不触发 CollectionView 的整表遍历复制。
/// </para>
/// <para>
/// LRU 缓存：RowView 物化后进入 <c>_cache</c>（链表+字典），命中时移到链表头，超 <c>MaxCacheRows</c>
/// 淘汰尾部。滚动 6.5 万行后不会全部驻留，限制在 ~500 个 RowView。
/// </para>
/// </summary>
public sealed class VirtualizingSortableView : IList, INotifyCollectionChanged
{
    /// <summary>LRU 缓存上限。实测虚拟化视口+预生成通常远低于此值，6.5 万行滚动不会全驻留。</summary>
    private const int MaxCacheRows = 500;

    private readonly ColumnStore _store;
    private readonly Func<int, int, string?> _cellAccessor;

    // LRU：链表头=最近使用，尾=最久未用。Dictionary O(1) 查节点。
    private readonly LinkedList<int> _lruKeys = new();
    private readonly Dictionary<int, (LinkedListNode<int> Node, RowView View)> _cache = [];

    private int[] _rowOrder;

    public VirtualizingSortableView(ColumnStore store, Func<int, int, string?>? cellAccessor = null)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _cellAccessor = cellAccessor ?? _store.GetCell;
        _rowOrder = BuildNaturalOrder(_store.RowCount);
    }

    public event NotifyCollectionChangedEventHandler? CollectionChanged;

    /// <summary>诊断用：已按需物化的 RowView 数量。虚拟化正确时它远小于 <see cref="Count"/>（筛选也不应整表物化）。</summary>
    public int MaterializedRowViewCount => _cache.Count;

    public int Count => _rowOrder.Length;

    public bool IsFixedSize => true;

    public bool IsReadOnly => true;

    public bool IsSynchronized => false;

    public object SyncRoot { get; } = new();

    public object? this[int index]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _rowOrder.Length);
            return GetOrCreate(_rowOrder[index]);
        }
        set => throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");
    }

    /// <summary>
    /// 按 <paramref name="col"/> 列排序：只重排 <c>_rowOrder</c> 这个 int 数组，
    /// 比较时按需读 ColumnStore，不复制整表。字符串序（与既有全 string 列模型一致）。
    /// </summary>
    public void SortBy(int col, bool ascending)
    {
        var order = ascending ? 1 : -1;
        Array.Sort(
            _rowOrder,
            (leftRow, rightRow) =>
                order
                * string.Compare(
                    _cellAccessor(leftRow, col),
                    _cellAccessor(rightRow, col),
                    StringComparison.Ordinal
                )
        );
        RaiseReset();
    }

    /// <summary>清除排序，恢复 ColumnStore 的自然行序（仍受当前筛选影响则以当前可见集合的自然序）。</summary>
    public void ClearSort()
    {
        Array.Sort(_rowOrder);
        RaiseReset();
    }

    /// <summary>
    /// 应用筛选：<c>_rowOrder</c> 收缩为满足 <paramref name="predicate"/> 的行号子集。
    /// 谓词按需读 ColumnStore（调用方自行决定读哪列），被过滤掉的行原地留在 store，不复制不移除。
    /// </summary>
    public void ApplyFilter(Func<int, bool> predicate)
    {
        ArgumentNullException.ThrowIfNull(predicate);
        var visible = new List<int>(_store.RowCount);
        for (var row = 0; row < _store.RowCount; row++)
        {
            if (predicate(row))
            {
                visible.Add(row);
            }
        }

        _rowOrder = [.. visible];
        _cache.Clear();
        _lruKeys.Clear();
        RaiseReset();
    }

    /// <summary>清除筛选，恢复全部行（保持当前排序会被重置为自然序，调用方需要时再排一次）。</summary>
    public void ClearFilter()
    {
        _rowOrder = BuildNaturalOrder(_store.RowCount);
        _cache.Clear();
        _lruKeys.Clear();
        RaiseReset();
    }

    public int IndexOf(object? value)
    {
        if (value is not RowView view)
        {
            return -1;
        }

        return Array.IndexOf(_rowOrder, view.RowIndex);
    }

    /// <summary>视图行号 → ColumnStore 真实行号。O(1)。</summary>
    public int GetStoreRowIndex(int viewIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(viewIndex);
        ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(viewIndex, _rowOrder.Length);
        return _rowOrder[viewIndex];
    }

    /// <summary>ColumnStore 真实行号 → 视图行号。线性查找 O(n)，仅小批量调用（如粘贴刷新）。</summary>
    public int GetViewIndex(int storeRowIndex)
    {
        return Array.IndexOf(_rowOrder, storeRowIndex);
    }

    public bool Contains(object? value) => IndexOf(value) >= 0;

    public IEnumerator GetEnumerator()
    {
        foreach (var row in _rowOrder)
        {
            yield return GetOrCreate(row);
        }
    }

    public void CopyTo(Array array, int index)
    {
        ArgumentNullException.ThrowIfNull(array);
        for (var i = 0; i < _rowOrder.Length; i++)
        {
            array.SetValue(GetOrCreate(_rowOrder[i]), index + i);
        }
    }

    public int Add(object? value) =>
        throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");

    public void Insert(int index, object? value) =>
        throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");

    public void Remove(object? value) =>
        throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");

    public void RemoveAt(int index) =>
        throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");

    public void Clear() =>
        throw new NotSupportedException("VirtualizingSortableView is virtualized by row.");

    private static int[] BuildNaturalOrder(int rowCount)
    {
        var order = new int[rowCount];
        for (var i = 0; i < rowCount; i++)
        {
            order[i] = i;
        }

        return order;
    }

    private RowView GetOrCreate(int row)
    {
        // LRU 命中：移到链表头（O(1)），返回已物化的 view。
        if (_cache.TryGetValue(row, out var entry))
        {
            _lruKeys.Remove(entry.Node);
            _lruKeys.AddFirst(entry.Node);
            return entry.View;
        }

        // 超限淘汰：链表尾=最久未用，从字典和链表同时移除。
        while (_cache.Count >= MaxCacheRows)
        {
            var evictRow = _lruKeys.Last!.Value;
            _cache.Remove(evictRow);
            _lruKeys.RemoveLast();
        }

        var node = _lruKeys.AddFirst(row);
        var view = new RowView(_store, row);
        _cache[row] = (node, view);
        return view;
    }

    private void RaiseReset() =>
        CollectionChanged?.Invoke(
            this,
            new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset)
        );
}
