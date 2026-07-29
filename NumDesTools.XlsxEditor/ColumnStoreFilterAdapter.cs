using System.ComponentModel;
using System.Windows.Controls;
using DataGridExtensions;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// P5：DataGridExtensions 的 <see cref="ICustomFilter"/> 适配器——把列头筛选路由到
/// <see cref="VirtualizingSortableView.ApplyFilter"/>（int 行号级），<b>绕开默认的 CollectionView.Filter
/// 整表枚举</b>。设为 DataGrid 的 <c>DataContext</c> 即生效。
/// <para>
/// 关键机制（见 DataGridFilterHost.EvaluateFilter 源码）：当 <c>DataGrid.DataContext is ICustomFilter</c> 时，
/// DataGridExtensions 执行 <c>collectionView.Filter = null;</c>（<b>不设逐项谓词、不枚举 65105 个 RowView</b>）
/// 然后调用本类的 <see cref="OnFilterChanged"/>，把带筛选值的列交给我们。我们据此读每列的筛选文本 +
/// 列类型，用 <see cref="ColumnFilterPredicate.Build"/> 编译成按行读 ColumnStore 的谓词，喂给 ApplyFilter。
/// </para>
/// <para>
/// #1 冻结+筛选共存：冻结行时主区只显示"第 N 行之后"的数据行，通过 <see cref="BasePredicate"/>（<c>row &gt;= n</c>）
/// 与列头筛选谓词 AND 组合——这样主区数据行照常可筛，冻结区（另一个 grid，固定表头）不受影响。
/// </para>
/// <para>
/// 列 → ColumnStore 列号：<see cref="DataGridColumn.SortMemberPath"/> 携带 store 列号（BuildDataColumns 里写入）。
/// 列类型：构造时注入的 <c>typeResolver</c>（复用 ColumnTypeDetector 采样结果）。
/// </para>
/// </summary>
public sealed class ColumnStoreFilterAdapter : ICustomFilter
{
    private readonly ColumnStore _store;
    private readonly VirtualizingSortableView _view;
    private readonly Func<int, ColumnType> _typeResolver;
    private readonly Action<int>? _onFilteredCountChanged;

    public ColumnStoreFilterAdapter(
        ColumnStore store,
        VirtualizingSortableView view,
        Func<int, ColumnType> typeResolver,
        Action<int>? onFilteredCountChanged = null
    )
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _view = view ?? throw new ArgumentNullException(nameof(view));
        _typeResolver = typeResolver ?? throw new ArgumentNullException(nameof(typeResolver));
        _onFilteredCountChanged = onFilteredCountChanged;
    }

    /// <summary>
    /// #1：基础行谓词，与列头筛选 AND 组合。冻结行模式下设为 <c>row =&gt; row &gt;= n</c>（只显示数据行区）；
    /// 非冻结设为 null（全表）。设值后立即重算一次筛选。
    /// </summary>
    public Func<int, bool>? BasePredicate { get; set; }

    /// <summary>本编辑器不支持多列排序（列头排序未启用，DataGridExtensions 排序钩子不触发）。</summary>
    public bool DisableMultipleColumnSorting => true;

    /// <summary>排序未启用（DataGrid.CanUserSortColumns=false），此钩子不会被有效触发；空实现。</summary>
    public void OnSortChanged(
        DataGrid dataGrid,
        IReadOnlyCollection<SortDescription> sortDescriptions
    )
    {
        // no-op: 列头排序未启用；排序由 VirtualizingSortableView.SortBy 单独驱动（如需要）。
    }

    /// <summary>
    /// DataGridExtensions 在筛选变化时调用（此前已把 CollectionView.Filter 置 null，无整表枚举）。
    /// 把带筛选值的列编译成 ColumnStore 行谓词，与 <see cref="BasePredicate"/> AND 组合后 ApplyFilter。
    /// </summary>
    public void OnFilterChanged(
        DataGrid dataGrid,
        IReadOnlyCollection<DataGridColumn> dataGridColumns
    )
    {
        var filters = new List<(int Col, string Value, ColumnType Type)>(dataGridColumns.Count);
        foreach (var column in dataGridColumns)
        {
            var text = ExtractFilterText(column.GetFilter());
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            var col = ResolveStoreColumn(column);
            if (col < 0 || col >= _store.ColumnCount)
            {
                continue;
            }

            filters.Add((col, text, _typeResolver(col)));
        }

        ApplyCombined(filters);
    }

    /// <summary>
    /// 重新应用当前列筛选 + <see cref="BasePredicate"/>（供设置 BasePredicate 后手动触发，
    /// 因为 DataGridExtensions 只在筛选框变化时回调 OnFilterChanged）。
    /// dataGridColumns 传当前 grid 的带筛选列。
    /// </summary>
    public void Reapply(IReadOnlyCollection<DataGridColumn> dataGridColumns)
    {
        OnFilterChanged(null!, dataGridColumns);
    }

    private void ApplyCombined(IReadOnlyList<(int Col, string Value, ColumnType Type)> filters)
    {
        var basePredicate = BasePredicate;

        if (filters.Count == 0)
        {
            if (basePredicate is null)
            {
                _view.ClearFilter();
            }
            else
            {
                _view.ApplyFilter(basePredicate);
            }
        }
        else
        {
            var columnPredicate = ColumnFilterPredicate.Build(_store, filters);
            _view.ApplyFilter(
                basePredicate is null
                    ? columnPredicate
                    : row => basePredicate(row) && columnPredicate(row)
            );
        }

        _onFilteredCountChanged?.Invoke(_view.Count);
    }

    /// <summary>从 DataGridExtensions 的 filter 对象取筛选文本（默认 SimpleContentFilter 的内容即字符串）。</summary>
    private static string ExtractFilterText(object? filter) =>
        filter switch
        {
            null => string.Empty,
            string s => s,
            _ => filter.ToString() ?? string.Empty,
        };

    /// <summary>列 → ColumnStore 列号：优先读 SortMemberPath（BuildDataColumns 写入的 store 列号）。</summary>
    private static int ResolveStoreColumn(DataGridColumn column) =>
        int.TryParse(column.SortMemberPath, out var idx) ? idx : column.DisplayIndex;
}
