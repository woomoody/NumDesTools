using System.Collections;
using System.Collections.Specialized;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 覆盖 <see cref="ColumnStore"/> 一个连续行区间 <c>[Start, Start+Count)</c> 的虚拟化视图（P4 冻结行用）。
/// 冻结行方案：上下两个 DataGrid 共享同一 ColumnStore——冻结 grid 绑 <c>RowRangeView(store, 0, N)</c>，
/// 主 grid 绑 <c>RowRangeView(store, N, RowCount-N)</c>。<see cref="this[int]"/> 生成的 <see cref="RowView"/>
/// 用<b>真实 store 行号</b>（Start + 局部索引），故两区编辑都经 RowView→ColumnStore 写回正确行、正确标脏。
/// <para>
/// 不排序、不筛选（冻结行是固定表头，参与排序语义模糊——见 status.md 里的语义选择）。
/// 与 <see cref="VirtualizingSortableView"/> 一样按行虚拟化：构造不物化任何行，RowView 按需生成并缓存。
/// </para>
/// </summary>
public sealed class RowRangeView : IList, INotifyCollectionChanged
{
    private readonly ColumnStore _store;
    private readonly Dictionary<int, RowView> _cache = [];
    private int _start;
    private int _count;

    /// <summary>
    /// 覆盖 <paramref name="store"/> 的 <c>[start, start+count)</c> 行区间。
    /// count 传负或超界会被夹到 <c>[0, RowCount-start]</c>。
    /// </summary>
    public RowRangeView(ColumnStore store, int start, int count)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        SetRange(start, count);
    }

    public event NotifyCollectionChangedEventHandler? CollectionChanged;

    /// <summary>诊断用：已按需物化的 RowView 数量（虚拟化正确时远小于 <see cref="Count"/>）。</summary>
    public int MaterializedRowViewCount => _cache.Count;

    /// <summary>区间起始 store 行号（0-based）。</summary>
    public int Start => _start;

    public int Count => _count;

    public bool IsFixedSize => true;

    public bool IsReadOnly => true;

    public bool IsSynchronized => false;

    public object SyncRoot { get; } = new();

    public object? this[int index]
    {
        get
        {
            ArgumentOutOfRangeException.ThrowIfNegative(index);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(index, _count);
            return GetOrCreate(_start + index);
        }
        set => throw new NotSupportedException("RowRangeView is virtualized by row.");
    }

    /// <summary>重设区间并广播 Reset（结构变更后调，让两个 grid 重新拉取正确行）。</summary>
    public void SetRange(int start, int count)
    {
        var rowCount = _store.RowCount;
        _start = Math.Clamp(start, 0, rowCount);
        _count = Math.Clamp(count, 0, rowCount - _start);
        _cache.Clear();
        RaiseReset();
    }

    public int IndexOf(object? value) =>
        value is RowView view && view.RowIndex >= _start && view.RowIndex < _start + _count
            ? view.RowIndex - _start
            : -1;

    public bool Contains(object? value) => IndexOf(value) >= 0;

    public IEnumerator GetEnumerator()
    {
        for (var i = 0; i < _count; i++)
        {
            yield return GetOrCreate(_start + i);
        }
    }

    public void CopyTo(Array array, int index)
    {
        ArgumentNullException.ThrowIfNull(array);
        for (var i = 0; i < _count; i++)
        {
            array.SetValue(GetOrCreate(_start + i), index + i);
        }
    }

    public int Add(object? value) =>
        throw new NotSupportedException("RowRangeView is virtualized by row.");

    public void Insert(int index, object? value) =>
        throw new NotSupportedException("RowRangeView is virtualized by row.");

    public void Remove(object? value) =>
        throw new NotSupportedException("RowRangeView is virtualized by row.");

    public void RemoveAt(int index) =>
        throw new NotSupportedException("RowRangeView is virtualized by row.");

    public void Clear() => throw new NotSupportedException("RowRangeView is virtualized by row.");

    private RowView GetOrCreate(int storeRow)
    {
        if (_cache.TryGetValue(storeRow, out var existing))
        {
            return existing;
        }

        var view = new RowView(_store, storeRow);
        _cache[storeRow] = view;
        return view;
    }

    private void RaiseReset() =>
        CollectionChanged?.Invoke(
            this,
            new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset)
        );
}
