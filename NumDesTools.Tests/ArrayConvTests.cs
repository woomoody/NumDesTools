namespace NumDesTools.Tests;

public class ArrayConvTests
{
    // ── ConvertListToArray ──────────────────────────────────────────────────

    [Fact]
    public void ConvertListToArray_ObjectList_ReturnsCorrect2D()
    {
        var input = new List<List<object>>
        {
            new() { "a", 1 },
            new() { "b", 2 },
        };
        var result = PubMetToExcel.ConvertListToArray(input);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(2, result.GetLength(1));
        Assert.Equal("a", result[0, 0]);
        Assert.Equal(2, result[1, 1]);
    }

    [Fact]
    public void ConvertListToArray_StringList_ReturnsCorrect2D()
    {
        var input = new List<List<string>>
        {
            new() { "x", "y" },
            new() { "1", "2" },
        };
        var result = PubMetToExcel.ConvertListToArray(input);
        Assert.Equal("x", result[0, 0]);
        Assert.Equal("2", result[1, 1]);
    }

    [Fact]
    public void ConvertListToArray_EmptyList_Returns0x0()
    {
        var result = PubMetToExcel.ConvertListToArray(new List<List<object>>());
        Assert.Equal(0, result.GetLength(0));
        Assert.Equal(0, result.GetLength(1));
    }

    [Fact]
    public void ConvertListToArray_ObjectList1D_ReturnsRow()
    {
        var input = new List<object> { 1, "two", 3.0 };
        var result = PubMetToExcel.ConvertListToArray(input);
        Assert.Equal(3, result.Length);
        Assert.Equal("two", result[1]);
    }

    // ── TwoDArrayToDictionary ───────────────────────────────────────────────

    [Fact]
    public void TwoDArrayToDictionary_KeysAre1Based()
    {
        // keys are 1-based (i+1) by design
        object[,] arr =
        {
            { "a", "b" },
            { "c", "d" },
        };
        var dict = PubMetToExcel.TwoDArrayToDictionary(arr);
        Assert.Equal(2, dict.Count);
        Assert.True(dict.ContainsKey(1));
        Assert.True(dict.ContainsKey(2));
        Assert.Equal(["a", "b"], dict[1].Select(x => x.ToString()));
        Assert.Equal(["c", "d"], dict[2].Select(x => x.ToString()));
    }

    [Fact]
    public void TwoDArrayToDictionaryFirstKey_IncludesKeyColInValue()
    {
        // value list includes the key column itself (j starts at 0)
        object[,] arr =
        {
            { "k1", "v1" },
            { "k2", "v2" },
        };
        var dict = PubMetToExcel.TwoDArrayToDictionaryFirstKey(arr);
        Assert.True(dict.ContainsKey("k1"));
        Assert.Equal(new[] { "k1", "v1" }, dict["k1"]);
    }

    // ── TwoDArrayToDicFirstKeyStr ───────────────────────────────────────────

    [Fact]
    public void TwoDArrayToDicFirstKeyStr_ValueExcludesKeyCol()
    {
        // 首列是 key，value 只包含剩余列（不含 key 列自身）
        // 此函数专用于 _通配符 表：key=名称，value=函数描述符（如 "Var#字段名"）
        object[,] arr =
        {
            { "k1", "a", "b", "c" },
            { "k2", "x", "y", "z" },
        };
        var dict = PubMetToExcel.TwoDArrayToDicFirstKeyStr(arr);
        Assert.Equal("a#b#c", dict["k1"]);
        Assert.Equal("x#y#z", dict["k2"]);
    }

    [Fact]
    public void TwoDArrayToDicFirstKeyStr_TwoColTable_ReturnsDescriptorAsIs()
    {
        // _通配符 表典型格式：key 列 + 函数描述符列，value 应直接返回描述符不含 key
        object[,] arr =
        {
            { "寻找编号2", "Var#寻找ID2" },
            { "物品编号", "Left#物品编号#5" },
        };
        var dict = PubMetToExcel.TwoDArrayToDicFirstKeyStr(arr);
        Assert.Equal("Var#寻找ID2", dict["寻找编号2"]);
        Assert.Equal("Left#物品编号#5", dict["物品编号"]);
    }

    // ── IsValidArray ────────────────────────────────────────────────────────

    [Fact]
    public void IsValidArray_ValidJson_ReturnsTrueAndArray()
    {
        var ok = PubMetToExcel.IsValidArray("[1,2,3]", out object[] arr);
        Assert.True(ok);
        Assert.Equal(3, arr.Length);
    }

    [Fact]
    public void IsValidArray_InvalidJson_ReturnsFalse()
    {
        var ok = PubMetToExcel.IsValidArray("not_json", out object[] _);
        Assert.False(ok);
    }

    [Fact]
    public void IsValidArray_Jagged_ReturnsTrueAndJagged()
    {
        var ok = PubMetToExcel.IsValidArray("[[1,2],[3,4]]", out object[][] arr);
        Assert.True(ok);
        Assert.Equal(2, arr.Length);
    }

    // ── Merge2DArrays0 (列合并，行数必须相同) ─────────────────────────────

    [Fact]
    public void Merge2DArrays0_MergesCols_SameRowCount()
    {
        // Merge2DArrays0 appends cols (requires equal row count)
        object[,] a =
        {
            { "a1" },
            { "a2" },
        };
        object[,] b =
        {
            { "b1" },
            { "b2" },
        };
        var result = PubMetToExcel.Merge2DArrays0(a, b);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(2, result.GetLength(1));
        Assert.Equal("b1", result[0, 1]);
        Assert.Equal("b2", result[1, 1]);
    }

    [Fact]
    public void Merge2DArrays0_MismatchedRows_Throws()
    {
        object[,] a =
        {
            { "a" },
        };
        object[,] b =
        {
            { "b" },
            { "c" },
        };
        Assert.Throws<InvalidOperationException>(() => PubMetToExcel.Merge2DArrays0(a, b));
    }

    // ── ConvertToCommaSeparatedArray ────────────────────────────────────────

    [Fact]
    public void ConvertToCommaSeparatedArray_JoinsRows()
    {
        object[,] arr =
        {
            { "a", "b", "c" },
            { "1", "2", "3" },
        };
        var result = PubMetToExcel.ConvertToCommaSeparatedArray(arr);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(1, result.GetLength(1));
        Assert.Equal("a,b,c", result[0, 0]?.ToString());
    }

    // ── IsArrayOfType ───────────────────────────────────────────────────────

    [Fact]
    public void IsArrayOfType_AllInts_ReturnsTrue()
    {
        var arr = new object[] { 1, 2, 3 };
        Assert.True(PubMetToExcel.IsArrayOfType(arr, typeof(int)));
    }

    [Fact]
    public void IsArrayOfType_AllStrings_ReturnsTrue()
    {
        var arr = new object[] { "a", "b" };
        Assert.True(PubMetToExcel.IsArrayOfType(arr, typeof(string)));
    }
}
