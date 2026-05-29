namespace NumDesTools.Tests;

public class CleanRepeatValueTests
{
    private static object[,] Make(object[][] rows)
    {
        var r = rows.Length;
        var c = rows[0].Length;
        var arr = new object[r, c];
        for (var i = 0; i < r; i++)
        for (var j = 0; j < c; j++)
            arr[i, j] = rows[i][j];
        return arr;
    }

    private static object[,] Make1Based(object[][] data)
    {
        var rows = data.Length;
        var cols = data[0].Length;
        var arr = (object[,])Array.CreateInstance(typeof(object), [rows, cols], [1, 1]);
        for (var i = 0; i < rows; i++)
        for (var j = 0; j < cols; j++)
            arr[i + 1, j + 1] = data[i][j];
        return arr;
    }

    [Fact]
    public void IsRowTrue_RemovesDuplicateCols_ByKeyRow()
    {
        var array = Make([
            ["A", "B", "A"],
            [1, 2, 3],
            [4, 5, 6],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 0, isRow: true, baseIndex: 0);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(3, result.GetLength(1));
        Assert.Equal("A", result[0, 0]);
        Assert.Equal("B", result[1, 0]);
    }

    [Fact]
    public void IsRowFalse_RemovesDuplicateRows_ByKeyCol()
    {
        var array = Make([
            ["X", 1],
            ["Y", 2],
            ["X", 3],
            ["Z", 4],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 0, isRow: false, baseIndex: 0);
        Assert.Equal(3, result.GetLength(0));
        Assert.Equal("X", result[0, 0]);
        Assert.Equal("Y", result[1, 0]);
        Assert.Equal("Z", result[2, 0]);
    }

    [Fact]
    public void EmptyFilterTrue_SkipsNullKeys()
    {
        var array = Make([
            ["A", 1],
            [null!, 2],
            ["B", 3],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(
            array,
            index: 0,
            isRow: false,
            baseIndex: 0,
            emptyFilter: true
        );
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal("A", result[0, 0]);
        Assert.Equal("B", result[1, 0]);
    }

    [Fact]
    public void EmptyFilterTrue_SkipsWhitespaceStringKeys()
    {
        var array = Make([
            ["A", 1],
            ["   ", 2],
            ["B", 3],
            ["", 4],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(
            array,
            index: 0,
            isRow: false,
            baseIndex: 0,
            emptyFilter: true
        );
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal("A", result[0, 0]);
        Assert.Equal("B", result[1, 0]);
    }

    [Fact]
    public void EmptyFilterFalse_KeepsNullKeyRows()
    {
        var array = Make([
            ["A", 1],
            [null!, 2],
            ["B", 3],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(
            array,
            index: 0,
            isRow: false,
            baseIndex: 0,
            emptyFilter: false
        );
        Assert.Equal(3, result.GetLength(0));
        Assert.Null(result[1, 0]);
    }

    [Fact]
    public void EmptyFilterFalse_DuplicateNullKeys_OnlyFirstKept()
    {
        var array = Make([
            [null!, 1],
            [null!, 2],
            ["A", 3],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(
            array,
            index: 0,
            isRow: false,
            baseIndex: 0,
            emptyFilter: false
        );
        Assert.Equal(2, result.GetLength(0));
        Assert.Null(result[0, 0]);
        Assert.Equal(1, result[0, 1]);
        Assert.Equal("A", result[1, 0]);
    }

    [Fact]
    public void BaseIndex1_IsRowFalse_DedupRows_On1BasedArray()
    {
        var array = Make1Based([
            ["X", 10],
            ["X", 20],
            ["Y", 30],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 1, isRow: false, baseIndex: 1);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal("X", result[0, 0]);
        Assert.Equal("Y", result[1, 0]);
    }

    [Fact]
    public void BaseIndex1_IsRowTrue_DedupCols_On1BasedArray()
    {
        var array = Make1Based([
            ["P", "Q", "P"],
            [10, 20, 30],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 1, isRow: true, baseIndex: 1);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal("P", result[0, 0]);
        Assert.Equal("Q", result[1, 0]);
    }

    [Fact]
    public void SingleRow_IsRowFalse_ReturnsSameRow()
    {
        var array = Make([
            ["alpha", "beta", "gamma"],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 0, isRow: false, baseIndex: 0);
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal(3, result.GetLength(1));
        Assert.Equal("alpha", result[0, 0]);
    }

    [Fact]
    public void AllRowsSameKey_OnlyFirstRowKept()
    {
        var array = Make([
            ["dup", 1],
            ["dup", 2],
            ["dup", 3],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 0, isRow: false, baseIndex: 0);
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal("dup", result[0, 0]);
        Assert.Equal(1, result[0, 1]);
    }

    [Fact]
    public void AllColumnsSameKey_OnlyFirstColumnKept()
    {
        var array = Make([
            ["same", "same", "same"],
            [10, 20, 30],
        ]);
        var result = PubMetToExcel.CleanRepeatValue(array, index: 0, isRow: true, baseIndex: 0);
        Assert.Equal(1, result.GetLength(0));
        Assert.Equal("same", result[0, 0]);
        Assert.Equal(10, result[0, 1]);
    }
}
