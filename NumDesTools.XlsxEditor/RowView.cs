using System.ComponentModel;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 单行的轻量代理：只持有 <see cref="ColumnStore"/> 引用 + 行号，读写直接转发到 ColumnStore，
/// <b>不在内部缓存整行数据</b>。WPF DataGrid 通过索引器绑定（<c>Binding "[列名]"</c>），
/// 编辑经 <see cref="ColumnStore.SetCell"/> 写回并标脏，触发 <see cref="INotifyPropertyChanged"/>
/// 让绑定刷新。这是数据虚拟化的关键：6.5 万行只在视口需要时按需生成 RowView，不预先物化整表对象树。
/// </summary>
public sealed class RowView(ColumnStore store, int rowIndex) : INotifyPropertyChanged
{
    private static readonly PropertyChangedEventArgs IndexerChangedArgs = new("Item[]");
    private static readonly PropertyChangedEventArgs DirtyStateChangedArgs = new(
        nameof(DirtyState)
    );

    private readonly ColumnStore _store = store ?? throw new ArgumentNullException(nameof(store));
    private int _dirtyState;

    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>该视图对应 ColumnStore 中的真实行号（0-based）。</summary>
    public int RowIndex { get; } = rowIndex;

    /// <summary>
    /// 脏状态版本。单元格样式绑定此值，以便某列编辑或保存清脏后重新查询
    /// <see cref="IsColumnDirty"/>；值本身没有业务含义。
    /// </summary>
    public int DirtyState => _dirtyState;

    /// <summary>按列索引读写。写入走 <see cref="ColumnStore.SetCell"/>（驻留 + 标脏）。</summary>
    public string? this[int col]
    {
        get => _store.GetCell(RowIndex, col);
        set
        {
            _store.SetCell(RowIndex, col, value);
            RaiseIndexerChanged();
            RefreshDirtyState();
        }
    }

    /// <summary>按 Excel 列名（A/B/.../CF）读写，内部解析为列索引后转发。</summary>
    public string? this[string columnName]
    {
        get => this[ResolveColumn(columnName)];
        set => this[ResolveColumn(columnName)] = value;
    }

    private int ResolveColumn(string columnName)
    {
        ArgumentNullException.ThrowIfNull(columnName);
        var col = _store.IndexOfColumn(columnName);
        if (col < 0)
        {
            throw new ArgumentException($"Column '{columnName}' not found", nameof(columnName));
        }

        return col;
    }

    public bool IsColumnDirty(int col) => _store.IsDirty(RowIndex, col);

    public void RefreshDirtyState()
    {
        _dirtyState++;
        PropertyChanged?.Invoke(this, DirtyStateChangedArgs);
    }

    private void RaiseIndexerChanged() => PropertyChanged?.Invoke(this, IndexerChangedArgs);
}
