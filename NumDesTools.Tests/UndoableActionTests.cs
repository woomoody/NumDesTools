using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// #6 结构性撤销/重做：增删行、增删列作为一个撤销单元，Ctrl+Z 整体撤销（内容 + 位置精确还原）、
/// Ctrl+Y 整体重做。<see cref="IUndoableAction"/> 只操作 <see cref="ColumnStore"/>（纯数据，可单测）。
/// 覆盖每条入口的"做操作 → Undo → 状态还原 → Redo → 状态复原"闭环。
/// </summary>
public sealed class UndoableActionTests
{
    private static ColumnStore Make(int rows, int cols = 3)
    {
        var names = Enumerable.Range(0, cols).Select(c => ((char)('A' + c)).ToString()).ToArray();
        var store = ColumnStore.Create(names, rows);
        for (var r = 0; r < rows; r++)
        {
            store.AppendRow();
            for (var c = 0; c < cols; c++)
            {
                store.SetCellQuiet(r, c, $"r{r}c{c}");
            }
        }
        store.ClearDirty();
        return store;
    }

    // ── 单格 / 批量编辑（CellBatchAction）─────────────────────────────

    [Fact]
    public void CellBatch_Undo_RestoresOldValues_Redo_RestoresNew()
    {
        var store = Make(3);
        var batch = new List<CellEditRecord>();
        (int r, int c, string v)[] edits = [(0, 0, "P00"), (1, 1, "P11"), (2, 2, "P22")];
        foreach (var (r, c, v) in edits)
        {
            batch.Add(new CellEditRecord(r, c, store.GetCell(r, c), v));
            store.SetCell(r, c, v);
        }
        var action = new CellBatchAction(batch);

        action.Undo(store);
        Assert.Equal("r0c0", store.GetCell(0, 0));
        Assert.Equal("r1c1", store.GetCell(1, 1));
        Assert.Equal("r2c2", store.GetCell(2, 2));

        action.Redo(store);
        Assert.Equal("P00", store.GetCell(0, 0));
        Assert.Equal("P11", store.GetCell(1, 1));
        Assert.Equal("P22", store.GetCell(2, 2));

        Assert.False(action.IsStructural);
    }

    // ── 插入行（InsertRowAction）──────────────────────────────────────

    [Fact]
    public void InsertRow_Undo_RemovesRow_Redo_ReAddsEmptyRow()
    {
        var store = Make(3); // rows 0,1,2
        // 在 row1 下方插入（at=1，原 row1 及以后下移）
        store.InsertRow(1);
        var action = new InsertRowAction(1);
        Assert.Equal(4, store.RowCount);
        Assert.Null(store.GetCell(1, 0)); // 新插入空行
        Assert.Equal("r1c0", store.GetCell(2, 0)); // 原 row1 下移到 row2

        action.Undo(store);
        Assert.Equal(3, store.RowCount);
        Assert.Equal("r1c0", store.GetCell(1, 0)); // 原 row1 回位

        action.Redo(store);
        Assert.Equal(4, store.RowCount);
        Assert.Null(store.GetCell(1, 0));
        Assert.Equal("r1c0", store.GetCell(2, 0));

        Assert.True(action.IsStructural);
    }

    [Fact]
    public void AppendRow_AsInsertAtEnd_UndoRedo()
    {
        var store = Make(2); // rows 0,1
        store.InsertRow(store.RowCount); // append at end (at=2)
        var action = new InsertRowAction(2);
        Assert.Equal(3, store.RowCount);

        action.Undo(store);
        Assert.Equal(2, store.RowCount);
        Assert.Equal("r1c0", store.GetCell(1, 0));

        action.Redo(store);
        Assert.Equal(3, store.RowCount);
    }

    // ── 删除行（DeleteRowsAction）────────────────────────────────────

    [Fact]
    public void DeleteRows_Undo_RestoresContentAndPosition_Redo_DeletesAgain()
    {
        var store = Make(5); // rows 0..4
        // 删除 row1 和 row3（多选）；删除前抓取内容快照
        var snap = new List<(int, string?[])>
        {
            (1, [store.GetCell(1, 0), store.GetCell(1, 1), store.GetCell(1, 2)]),
            (3, [store.GetCell(3, 0), store.GetCell(3, 1), store.GetCell(3, 2)]),
        };
        // 降序删除（生产代码即降序，避免行号移位）
        store.DeleteRow(3);
        store.DeleteRow(1);
        var action = new DeleteRowsAction(snap);
        Assert.Equal(3, store.RowCount);
        Assert.Equal("r0c0", store.GetCell(0, 0));
        Assert.Equal("r2c0", store.GetCell(1, 0)); // 原 row2 上移到 row1
        Assert.Equal("r4c0", store.GetCell(2, 0)); // 原 row4 上移到 row2

        // 撤销：两行内容 + 位置精确还原
        action.Undo(store);
        Assert.Equal(5, store.RowCount);
        Assert.Equal("r0c0", store.GetCell(0, 0));
        Assert.Equal("r1c1", store.GetCell(1, 1)); // row1 内容还原
        Assert.Equal("r2c0", store.GetCell(2, 0));
        Assert.Equal("r3c2", store.GetCell(3, 2)); // row3 内容还原
        Assert.Equal("r4c0", store.GetCell(4, 0));

        // 重做：再次删除
        action.Redo(store);
        Assert.Equal(3, store.RowCount);
        Assert.Equal("r2c0", store.GetCell(1, 0));
        Assert.Equal("r4c0", store.GetCell(2, 0));

        Assert.True(action.IsStructural);
    }

    [Fact]
    public void DeleteSingleRow_UndoRedo()
    {
        var store = Make(3);
        var snap = new List<(int, string?[])>
        {
            (0, [store.GetCell(0, 0), store.GetCell(0, 1), store.GetCell(0, 2)]),
        };
        store.DeleteRow(0);
        var action = new DeleteRowsAction(snap);
        Assert.Equal("r1c0", store.GetCell(0, 0));

        action.Undo(store);
        Assert.Equal(3, store.RowCount);
        Assert.Equal("r0c0", store.GetCell(0, 0));

        action.Redo(store);
        Assert.Equal(2, store.RowCount);
        Assert.Equal("r1c0", store.GetCell(0, 0));
    }

    // ── 插入列（InsertColumnAction）──────────────────────────────────

    [Fact]
    public void InsertColumn_Undo_RemovesLastColumn_Redo_ReAdds()
    {
        var store = Make(3, cols: 3); // A B C
        store.EnsureColumnCount(4, _ => "D");
        var action = new InsertColumnAction(_ => "D");
        Assert.Equal(4, store.ColumnCount);
        Assert.Equal("D", store.ColumnNames[3]);

        action.Undo(store);
        Assert.Equal(3, store.ColumnCount);
        Assert.Equal("C", store.ColumnNames[^1]);

        action.Redo(store);
        Assert.Equal(4, store.ColumnCount);
        Assert.Equal("D", store.ColumnNames[3]);

        Assert.True(action.IsStructural);
    }

    [Fact]
    public void RemoveLastColumn_PrunesDirtyForThatColumn()
    {
        var store = Make(2, cols: 2); // A B
        store.EnsureColumnCount(3, _ => "C");
        store.SetCell(0, 2, "dirtyC"); // 脏格在末列
        store.SetCell(0, 0, "dirtyA");
        Assert.True(store.IsDirty(0, 2));

        store.RemoveLastColumn();
        Assert.Equal(2, store.ColumnCount);
        Assert.True(store.IsDirty(0, 0)); // A 列脏保留
        // C 列的脏条目已随列删除被剪掉（列号 2 已不存在）
        Assert.DoesNotContain(store.DirtyCells, cell => cell.Col == 2);
    }

    // ── 混合序列：模拟真实操作流的撤销/重做栈 ─────────────────────────

    [Fact]
    public void MixedSequence_UndoRedoStack_RestoresStateStepByStep()
    {
        var store = Make(3);
        var undo = new Stack<IUndoableAction>();
        var redo = new Stack<IUndoableAction>();

        // 1) 编辑 (0,0)
        var e1 = new CellBatchAction([new CellEditRecord(0, 0, store.GetCell(0, 0), "X")]);
        store.SetCell(0, 0, "X");
        undo.Push(e1);
        redo.Clear();

        // 2) 插入行 at=1
        store.InsertRow(1);
        undo.Push(new InsertRowAction(1));
        redo.Clear();
        Assert.Equal(4, store.RowCount);

        // 撤销插入行
        undo.Pop().Undo(store);
        Assert.Equal(3, store.RowCount);
        Assert.Equal("X", store.GetCell(0, 0)); // 编辑仍在

        // 撤销编辑
        undo.Pop().Undo(store);
        Assert.Equal("r0c0", store.GetCell(0, 0));
        Assert.Empty(undo);
    }

    // ── UndoableStack 统一重放（生产 OnUndoClick/OnRedoClick 走的路径）────────

    [Fact]
    public void UndoableStack_MixedOps_UndoRedoInReverseChronologicalOrder()
    {
        var store = Make(4); // rows 0..3, cols A B C
        var undo = new Stack<IUndoableAction>();
        var redo = new Stack<IUndoableAction>();

        // op1: 编辑 (0,0)=r0c0 → E1
        store.SetCell(0, 0, "E1");
        undo.Push(new CellBatchAction([new CellEditRecord(0, 0, "r0c0", "E1")]));
        redo.Clear();

        // op2: 插入行 at=2
        store.InsertRow(2);
        undo.Push(new InsertRowAction(2));
        redo.Clear();
        Assert.Equal(5, store.RowCount);

        // op3: 删除行 0（先快照）
        var snap = new List<(int, string?[])>
        {
            (0, [store.GetCell(0, 0), store.GetCell(0, 1), store.GetCell(0, 2)]),
        };
        store.DeleteRow(0);
        undo.Push(new DeleteRowsAction(snap));
        redo.Clear();
        Assert.Equal(4, store.RowCount);
        Assert.Equal("r1c0", store.GetCell(0, 0)); // 原 row1 上移

        // 撤销 op3（恢复删除的行 0）
        UndoableStack.Undo(store, undo, redo);
        Assert.Equal(5, store.RowCount);
        Assert.Equal("E1", store.GetCell(0, 0)); // 被删行内容还原（含 op1 编辑）

        // 撤销 op2（删掉插入的空行）
        UndoableStack.Undo(store, undo, redo);
        Assert.Equal(4, store.RowCount);

        // 撤销 op1（编辑回退）
        UndoableStack.Undo(store, undo, redo);
        Assert.Equal("r0c0", store.GetCell(0, 0));
        Assert.Empty(undo);
        Assert.Equal(3, redo.Count);

        // 全部重做，回到 op3 后状态
        UndoableStack.Redo(store, undo, redo); // op1
        Assert.Equal("E1", store.GetCell(0, 0));
        UndoableStack.Redo(store, undo, redo); // op2
        Assert.Equal(5, store.RowCount);
        UndoableStack.Redo(store, undo, redo); // op3
        Assert.Equal(4, store.RowCount);
        Assert.Equal("r1c0", store.GetCell(0, 0));
        Assert.Empty(redo);
    }

    [Fact]
    public void UndoableStack_EmptyStacks_NoThrow()
    {
        var store = Make(2);
        var undo = new Stack<IUndoableAction>();
        var redo = new Stack<IUndoableAction>();

        UndoableStack.Undo(store, undo, redo);
        UndoableStack.Redo(store, undo, redo);
        Assert.Empty(undo);
        Assert.Empty(redo);
    }

    [Fact]
    public void UndoableStack_ColumnInsert_UndoRedo()
    {
        var store = Make(2, cols: 2); // A B
        var undo = new Stack<IUndoableAction>();
        var redo = new Stack<IUndoableAction>();

        store.EnsureColumnCount(3, _ => "C");
        undo.Push(new InsertColumnAction(_ => "C"));
        redo.Clear();
        Assert.Equal(3, store.ColumnCount);

        UndoableStack.Undo(store, undo, redo);
        Assert.Equal(2, store.ColumnCount);

        UndoableStack.Redo(store, undo, redo);
        Assert.Equal(3, store.ColumnCount);
        Assert.Equal("C", store.ColumnNames[2]);
    }
}
