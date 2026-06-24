using NumDesTools.ConflictResolver;

namespace NumDesTools.Tests;

public class ConflictModelsTests
{
    // DefaultRowChoice（Differ 构造时调用）把 OnlyOurs/OnlyTheirs 行标记为"系统已自动选好"
    // = IsResolved=true，让 BatchAutoResolve 可以直接 apply，无需用户点击。
    [Fact]
    public void OnlyOursRow_IsResolved_TrueWhenBuiltByDiffer()
    {
        var row = new RowConflict
        {
            DiffType = RowDiffType.OnlyOurs,
            DefaultRowChoice = ConflictChoice.Ours,
        };
        Assert.True(row.IsResolved); // Differ 设置的默认即视为"已解决"
    }

    [Fact]
    public void OnlyTheirsRow_IsResolved_TrueWhenBuiltByDiffer()
    {
        var row = new RowConflict
        {
            DiffType = RowDiffType.OnlyTheirs,
            DefaultRowChoice = ConflictChoice.Theirs,
        };
        Assert.True(row.IsResolved);
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
