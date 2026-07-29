namespace NumDesTools.XlsxEditor;

/// <summary>
/// #6：统一的"可撤销操作"抽象。一次用户动作（单格编辑、多格粘贴、增删行、增删列）压入一个
/// <see cref="IUndoableAction"/> 作为<b>一个撤销单元</b>——Ctrl+Z 调 <see cref="Undo"/> 整体撤销，
/// Ctrl+Y 调 <see cref="Redo"/> 整体重做。所有实现只操作纯数据层 <see cref="ColumnStore"/>，不碰 WPF，
/// 故可单测。
/// <para>
/// 结构性操作（增删行/列）与 <see cref="ColumnStore"/> 的脏跟踪 remap 是<b>两套机制</b>：脏跟踪服务
/// 增量写回（数据层），撤销栈服务用户的 Ctrl+Z（交互层）。撤销一次增删行必须把行内容和位置都还原正确
/// （见 <see cref="DeleteRowsAction"/> 保存被删行内容）。
/// </para>
/// </summary>
public interface IUndoableAction
{
    /// <summary>撤销：把 <paramref name="store"/> 还原到本操作发生<b>前</b>的状态。</summary>
    void Undo(ColumnStore store);

    /// <summary>重做：把 <paramref name="store"/> 重新应用本操作（还原到发生<b>后</b>的状态）。</summary>
    void Redo(ColumnStore store);

    /// <summary>本操作是否改变了行/列结构（增删行列）。UI 据此决定撤销/重做后是否重建视图行序/列。</summary>
    bool IsStructural { get; }
}

/// <summary>
/// 单格编辑 / 多格粘贴：一批 <see cref="CellEditRecord"/> 作为一个撤销单元。
/// 单格编辑 = 1 元素批次；粘贴 = N 元素批次。撤销整批恢复 OldValue，重做整批写 NewValue。
/// 越界行（罕见，理论上结构撤销后可能出现）静默跳过，不抛。
/// </summary>
public sealed class CellBatchAction(IReadOnlyList<CellEditRecord> edits) : IUndoableAction
{
    private readonly IReadOnlyList<CellEditRecord> _edits =
        edits ?? throw new ArgumentNullException(nameof(edits));

    public bool IsStructural => false;

    public void Undo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        foreach (var record in _edits)
        {
            if (IsValid(store, record.Row, record.Col))
            {
                store.SetCell(record.Row, record.Col, record.OldValue?.ToString());
            }
        }
    }

    public void Redo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        foreach (var record in _edits)
        {
            if (IsValid(store, record.Row, record.Col))
            {
                store.SetCell(record.Row, record.Col, record.NewValue);
            }
        }
    }

    private static bool IsValid(ColumnStore store, int row, int col) =>
        row >= 0 && row < store.RowCount && col >= 0 && col < store.ColumnCount;
}

/// <summary>
/// 在 <see cref="At"/> 处插入了一个空行。撤销 = 删除该行；重做 = 再插入一个空行。
/// （新插入的行内容为空，故撤销无需保存内容。）
/// </summary>
public sealed class InsertRowAction(int at) : IUndoableAction
{
    public int At { get; } = at;

    public bool IsStructural => true;

    public void Undo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        if (At >= 0 && At < store.RowCount)
        {
            store.DeleteRow(At);
        }
    }

    public void Redo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        if (At >= 0 && At <= store.RowCount)
        {
            store.InsertRow(At);
        }
    }
}

/// <summary>
/// 删除了一批行（支持多选）。撤销时按<b>升序</b>重新插入每行并还原其完整内容（删除前已快照），
/// 重做时按<b>降序</b>再次删除。行内容 + 位置都精确还原。
/// </summary>
public sealed class DeleteRowsAction : IUndoableAction
{
    // 每个被删行：删除前的绝对行号（升序）+ 该行所有列的值快照
    private readonly List<(int Row, string?[] Values)> _deleted;

    /// <param name="deletedRows">删除前抓取的 (行号, 整行值快照)，行号须为删除前的绝对行号。</param>
    public DeleteRowsAction(IReadOnlyList<(int Row, string?[] Values)> deletedRows)
    {
        ArgumentNullException.ThrowIfNull(deletedRows);
        _deleted = deletedRows.OrderBy(d => d.Row).ToList();
    }

    public bool IsStructural => true;

    /// <summary>撤销：按行号升序重新插入并回填每行内容（升序保证后插入的行号不被前面的插入顶偏）。</summary>
    public void Undo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        foreach (var (row, values) in _deleted)
        {
            var at = Math.Clamp(row, 0, store.RowCount);
            store.InsertRow(at);
            for (var col = 0; col < values.Length && col < store.ColumnCount; col++)
            {
                store.SetCell(at, col, values[col]);
            }
        }
    }

    /// <summary>重做：按行号降序再次删除（降序避免删除后行号移位）。</summary>
    public void Redo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        foreach (var (row, _) in _deleted.OrderByDescending(d => d.Row))
        {
            if (row >= 0 && row < store.RowCount)
            {
                store.DeleteRow(row);
            }
        }
    }
}

/// <summary>
/// 在末尾追加了一列（<see cref="ColumnStore.EnsureColumnCount"/> 只能加最右）。撤销 = 删除最后一列；
/// 重做 = 再追加一列（列名由 <paramref name="nameFactory"/> 生成，与首次一致）。
/// </summary>
public sealed class InsertColumnAction(Func<int, string> nameFactory) : IUndoableAction
{
    private readonly Func<int, string> _nameFactory =
        nameFactory ?? throw new ArgumentNullException(nameof(nameFactory));

    public bool IsStructural => true;

    public void Undo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        store.RemoveLastColumn();
    }

    public void Redo(ColumnStore store)
    {
        ArgumentNullException.ThrowIfNull(store);
        store.EnsureColumnCount(store.ColumnCount + 1, _nameFactory);
    }
}

/// <summary>
/// #6：<see cref="IUndoableAction"/> 撤销/重做栈的重放逻辑（纯静态、不碰 WPF、可单测）。
/// 一次 <see cref="Undo"/> 弹出栈顶 action 调其 <see cref="IUndoableAction.Undo"/> 并压入 redo；
/// action 自身对称（自带 Undo/Redo），重做时再 <see cref="IUndoableAction.Redo"/>。空栈无操作。
/// 覆盖"几乎所有可变动作"的统一撤销路径（单格/粘贴/增删行/增删列）。
/// </summary>
public static class UndoableStack
{
    public static void Undo(
        ColumnStore store,
        Stack<IUndoableAction> undo,
        Stack<IUndoableAction> redo
    )
    {
        ArgumentNullException.ThrowIfNull(store);
        ArgumentNullException.ThrowIfNull(undo);
        ArgumentNullException.ThrowIfNull(redo);
        if (undo.Count == 0)
        {
            return;
        }

        var action = undo.Pop();
        action.Undo(store);
        redo.Push(action);
    }

    public static void Redo(
        ColumnStore store,
        Stack<IUndoableAction> undo,
        Stack<IUndoableAction> redo
    )
    {
        ArgumentNullException.ThrowIfNull(store);
        ArgumentNullException.ThrowIfNull(undo);
        ArgumentNullException.ThrowIfNull(redo);
        if (redo.Count == 0)
        {
            return;
        }

        var action = redo.Pop();
        action.Redo(store);
        undo.Push(action);
    }
}
