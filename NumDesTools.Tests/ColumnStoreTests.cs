using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// ColumnStore 列式内存存储单测：验证形状、读写、脏跟踪、增删行、字符串驻留。
/// 全新代码，严格 RED→GREEN：本文件先于 ColumnStore.cs 存在（编译失败即 RED）。
/// </summary>
public sealed class ColumnStoreTests
{
    private static ColumnStore NewStore(int rows = 0) => ColumnStore.Create(["A", "B", "C"], rows);

    [Fact]
    public void Create_ExposesColumnShape()
    {
        var store = ColumnStore.Create(["A", "B", "C"], initialRowCapacity: 10);

        Assert.Equal(0, store.RowCount);
        Assert.Equal(3, store.ColumnCount);
        Assert.Equal(new[] { "A", "B", "C" }, store.ColumnNames);
    }

    [Fact]
    public void AppendRow_GrowsRowCount()
    {
        var store = NewStore();

        var first = store.AppendRow();
        var second = store.AppendRow();

        Assert.Equal(0, first);
        Assert.Equal(1, second);
        Assert.Equal(2, store.RowCount);
    }

    [Fact]
    public void SetCell_ThenGetCell_RoundTrips()
    {
        var store = NewStore();
        store.AppendRow();

        store.SetCell(0, 1, "hello");

        Assert.Equal("hello", store.GetCell(0, 1));
    }

    [Fact]
    public void GetCell_DefaultsToNull()
    {
        var store = NewStore();
        store.AppendRow();

        Assert.Null(store.GetCell(0, 0));
    }

    [Fact]
    public void SetCell_MarksDirty_And_TracksCoordinates()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();

        store.SetCell(1, 2, "x");

        Assert.True(store.IsDirty(1, 2));
        Assert.False(store.IsDirty(0, 0));
        Assert.Contains((1, 2), store.DirtyCells);
        Assert.Single(store.DirtyCells);
    }

    [Fact]
    public void SetCell_SameCellTwice_DirtyRecordedOnce()
    {
        var store = NewStore();
        store.AppendRow();

        store.SetCell(0, 0, "a");
        store.SetCell(0, 0, "b");

        Assert.Equal("b", store.GetCell(0, 0));
        Assert.Single(store.DirtyCells);
    }

    [Fact]
    public void SetCell_InternsEqualStrings_ToSameReference()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();
        store.AppendRow();

        // 构造两个内容相同但引用不同的字符串，防止 JIT 常量池干扰
        var a = new string("dup".ToCharArray());
        var b = new string("dup".ToCharArray());
        Assert.False(ReferenceEquals(a, b));

        store.SetCell(0, 0, a);
        store.SetCell(1, 0, b);
        store.SetCell(2, 1, b);

        var stored0 = store.GetCell(0, 0);
        var stored1 = store.GetCell(1, 0);
        var stored2 = store.GetCell(2, 1);

        Assert.Equal("dup", stored0);
        // 驻留生效：三处存的是同一个引用，不重复分配
        Assert.Same(stored0, stored1);
        Assert.Same(stored0, stored2);
    }

    [Fact]
    public void SetCell_NullValue_StaysNull_NotInterned()
    {
        var store = NewStore();
        store.AppendRow();

        store.SetCell(0, 0, null);

        Assert.Null(store.GetCell(0, 0));
        Assert.True(store.IsDirty(0, 0));
    }

    [Fact]
    public void InsertRow_ShiftsExistingRowsDown()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();
        store.SetCell(0, 0, "row0");
        store.SetCell(1, 0, "row1");

        store.InsertRow(1);

        Assert.Equal(3, store.RowCount);
        Assert.Equal("row0", store.GetCell(0, 0));
        Assert.Null(store.GetCell(1, 0)); // 新插入的空行
        Assert.Equal("row1", store.GetCell(2, 0));
    }

    [Fact]
    public void InsertRow_AtEnd_AppendsBlankRow()
    {
        var store = NewStore();
        store.AppendRow();
        store.SetCell(0, 0, "only");

        store.InsertRow(1);

        Assert.Equal(2, store.RowCount);
        Assert.Equal("only", store.GetCell(0, 0));
        Assert.Null(store.GetCell(1, 0));
    }

    [Fact]
    public void DeleteRow_ShiftsRemainingRowsUp()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();
        store.AppendRow();
        store.SetCell(0, 0, "row0");
        store.SetCell(1, 0, "row1");
        store.SetCell(2, 0, "row2");

        store.DeleteRow(1);

        Assert.Equal(2, store.RowCount);
        Assert.Equal("row0", store.GetCell(0, 0));
        Assert.Equal("row2", store.GetCell(1, 0));
    }

    [Fact]
    public void InsertRow_OutOfRange_Throws()
    {
        var store = NewStore();
        store.AppendRow();

        Assert.Throws<ArgumentOutOfRangeException>(() => store.InsertRow(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => store.InsertRow(2));
    }

    [Fact]
    public void DeleteRow_OutOfRange_Throws()
    {
        var store = NewStore();
        store.AppendRow();

        Assert.Throws<ArgumentOutOfRangeException>(() => store.DeleteRow(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => store.DeleteRow(1));
    }

    [Fact]
    public void GetCell_OutOfRange_Throws()
    {
        var store = NewStore();
        store.AppendRow();

        Assert.Throws<ArgumentOutOfRangeException>(() => store.GetCell(1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => store.GetCell(0, 3));
    }

    // ─────────────────────────────────────────────────────────────────
    //  P4 WF1: 结构操作脏跟踪 remap（取代 P2 的 _dirty.Clear()）+ StructureChanged + ClearDirty
    //  行为变更授权见 .remember\...\status.md（P4 任务书 工作流1）。
    //  动机：P2 里 InsertRow/DeleteRow 直接清脏，导致"编辑几格→再增删行→那几格脏标记丢失"，
    //  P4 增量写回会漏写这些格。改为 remap 行号，脏标记随行移动保留。
    // ─────────────────────────────────────────────────────────────────

    [Fact]
    public void InsertRow_RemapsDirtyRows_AtOrAfterInsertionPoint_ShiftDown()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();
        store.AppendRow();
        store.SetCell(0, 0, "r0"); // dirty (0,0) — 在插入点之前，不动
        store.SetCell(1, 1, "r1"); // dirty (1,1) — 在插入点，下移到 (2,1)
        store.SetCell(2, 2, "r2"); // dirty (2,2) — 在插入点之后，下移到 (3,2)

        store.InsertRow(1);

        Assert.Equal(4, store.RowCount);
        // 插入点之前的脏格不变
        Assert.True(store.IsDirty(0, 0));
        // 插入点及之后的脏格行号 +1
        Assert.False(store.IsDirty(1, 1));
        Assert.True(store.IsDirty(2, 1));
        Assert.False(store.IsDirty(2, 2));
        Assert.True(store.IsDirty(3, 2));
        // 脏集合大小不变（remap 而非清空）
        Assert.Equal(3, store.DirtyCells.Count);
    }

    [Fact]
    public void DeleteRow_DropsDirtyAtDeletedRow_AndShiftsLaterRowsUp()
    {
        var store = NewStore();
        store.AppendRow();
        store.AppendRow();
        store.AppendRow();
        store.SetCell(0, 0, "r0"); // dirty (0,0) — 删除点之前，不动
        store.SetCell(1, 1, "r1"); // dirty (1,1) — 删除点，丢弃
        store.SetCell(2, 2, "r2"); // dirty (2,2) — 删除点之后，上移到 (1,2)

        store.DeleteRow(1);

        Assert.Equal(2, store.RowCount);
        Assert.True(store.IsDirty(0, 0)); // 之前的脏格不变
        Assert.False(store.IsDirty(1, 1)); // 被删行的脏格已丢弃
        Assert.True(store.IsDirty(1, 2)); // 之后的脏格行号 -1
        Assert.Equal(2, store.DirtyCells.Count); // 删掉 1 个（被删行），remap 1 个
    }

    [Fact]
    public void InsertRow_SetsStructureChanged()
    {
        var store = NewStore();
        store.AppendRow();
        Assert.False(store.StructureChanged);

        store.InsertRow(0);

        Assert.True(store.StructureChanged);
    }

    [Fact]
    public void DeleteRow_SetsStructureChanged()
    {
        var store = NewStore();
        store.AppendRow();
        Assert.False(store.StructureChanged);

        store.DeleteRow(0);

        Assert.True(store.StructureChanged);
    }

    [Fact]
    public void EnsureColumnCount_Grow_SetsStructureChanged()
    {
        var store = NewStore();
        store.AppendRow();
        Assert.False(store.StructureChanged);

        store.EnsureColumnCount(5, col => $"col{col}");

        Assert.True(store.StructureChanged);
    }

    [Fact]
    public void EnsureColumnCount_NoGrow_DoesNotSetStructureChanged()
    {
        var store = NewStore(); // 3 列
        store.AppendRow();

        store.EnsureColumnCount(2, col => $"col{col}"); // 不扩（已有 3 列）

        Assert.False(store.StructureChanged);
    }

    [Fact]
    public void AppendRow_DoesNotSetStructureChanged()
    {
        // AppendRow 是加载快路径，不算结构性改动（不影响已有行号，脏跟踪无需 remap）
        var store = NewStore();

        store.AppendRow();

        Assert.False(store.StructureChanged);
    }

    [Fact]
    public void ClearDirty_ClearsDirtyCells_AndResetsStructureChanged()
    {
        var store = NewStore();
        store.AppendRow();
        store.SetCell(0, 0, "x");
        store.InsertRow(0);
        Assert.NotEmpty(store.DirtyCells);
        Assert.True(store.StructureChanged);

        store.ClearDirty();

        Assert.Empty(store.DirtyCells);
        Assert.False(store.StructureChanged);
    }

    [Fact]
    public void EditAfterClearDirty_MarksDirtyAgain()
    {
        var store = NewStore();
        store.AppendRow();
        store.SetCell(0, 0, "x");
        store.ClearDirty();
        Assert.Empty(store.DirtyCells);

        store.SetCell(0, 1, "y");

        Assert.True(store.IsDirty(0, 1));
        Assert.Single(store.DirtyCells);
    }
}
