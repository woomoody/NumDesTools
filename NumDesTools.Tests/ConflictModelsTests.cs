using NumDesTools.ConflictResolver;

namespace NumDesTools.Tests;

public class ConflictModelsTests
{
    // H4: Differ 在对象初始化器里设 RowChoice 默认值，不应触发"用户已选"标志
    // 构造方式与 ExcelConflictDiffer.DiffSheets 完全一致
    [Fact]
    public void OnlyOursRow_IsResolved_FalseWhenBuiltByDiffer()
    {
        // 复现 Differ 的构造：用 DefaultRowChoice 设默认值，不触发"用户已明确处理"
        var row = new RowConflict
        {
            DiffType = RowDiffType.OnlyOurs,
            DefaultRowChoice = ConflictChoice.Ours,
        };
        Assert.False(row.IsResolved);
    }

    [Fact]
    public void OnlyTheirsRow_IsResolved_FalseWhenBuiltByDiffer()
    {
        var row = new RowConflict
        {
            DiffType = RowDiffType.OnlyTheirs,
            DefaultRowChoice = ConflictChoice.Theirs,
        };
        Assert.False(row.IsResolved);
    }

    [Fact]
    public void OnlyOursRow_IsResolved_TrueAfterRowChoiceSet()
    {
        var row = new RowConflict { DiffType = RowDiffType.OnlyOurs };
        row.RowChoice = ConflictChoice.Ours;
        Assert.True(row.IsResolved);
    }

    [Fact]
    public void OnlyTheirsRow_IsResolved_TrueAfterRowChoiceSet()
    {
        var row = new RowConflict { DiffType = RowDiffType.OnlyTheirs };
        row.RowChoice = ConflictChoice.Theirs;
        Assert.True(row.IsResolved);
    }

    // Modified 行：所有 cell 都 explicit 才算解决
    [Fact]
    public void ModifiedRow_IsResolved_FalseWhenCellNotExplicit()
    {
        var row = new RowConflict { DiffType = RowDiffType.Modified };
        row.Cells.Add(new CellConflict { ColName = "col1" }); // IsExplicit=false
        Assert.False(row.IsResolved);
    }

    [Fact]
    public void ModifiedRow_IsResolved_TrueWhenAllCellsExplicit()
    {
        var row = new RowConflict { DiffType = RowDiffType.Modified };
        var cell = new CellConflict { ColName = "col1" };
        cell.IsExplicit = true;
        row.Cells.Add(cell);
        Assert.True(row.IsResolved);
    }
}
