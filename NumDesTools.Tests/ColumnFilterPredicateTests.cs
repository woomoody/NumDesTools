using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// P5 列头筛选谓词逻辑单测（类型感知 + 多列 AND + ColumnStore 列级访问）。
/// 这批逻辑从 P3.2 的 <c>MainWindow.MatchesColumnFilter</c>（private，不可测）提取到 public 静态
/// <see cref="ColumnFilterPredicate"/>，供 DataGridExtensions 的 ICustomFilter 适配器复用并单测。
/// 严格 RED→GREEN：本文件先于 ColumnFilterPredicate.cs 存在（编译失败即 RED）。
/// </summary>
public sealed class ColumnFilterPredicateTests
{
    // ── 单格匹配：文本 contains（大小写不敏感）──
    [Theory]
    [InlineData("Apple", "app", true)]
    [InlineData("Apple", "APP", true)]
    [InlineData("Apple", "ple", true)]
    [InlineData("Apple", "xyz", false)]
    [InlineData("", "x", false)]
    [InlineData("anything", "", true)] // 空筛选值 = 匹配（不过滤）
    public void Matches_Text_ContainsOrdinalIgnoreCase(string cell, string filter, bool expected)
    {
        Assert.Equal(expected, ColumnFilterPredicate.Matches(cell, filter, ColumnType.Text));
    }

    // ── 单格匹配：数字范围前缀 ──
    [Theory]
    [InlineData("100", ">=100", true)]
    [InlineData("99", ">=100", false)]
    [InlineData("100", "<=100", true)]
    [InlineData("101", "<=100", false)]
    [InlineData("101", ">100", true)]
    [InlineData("100", ">100", false)]
    [InlineData("99", "<100", true)]
    [InlineData("100", "<100", false)]
    [InlineData("100", "100", true)] // 无前缀 = 等于
    [InlineData("100", "=100", true)] // 显式 =
    [InlineData("100.5", ">=100", true)] // Float
    [InlineData("abc", ">=100", false)] // 非数字格 → 不匹配
    public void Matches_Numeric_RangePrefix(string cell, string filter, bool expected)
    {
        Assert.Equal(expected, ColumnFilterPredicate.Matches(cell, filter, ColumnType.Integer));
        // Float 走同一分支
        Assert.Equal(expected, ColumnFilterPredicate.Matches(cell, filter, ColumnType.Float));
    }

    // ── 单格匹配：Enum 精确 ──
    [Theory]
    [InlineData("A", "A", true)]
    [InlineData("A", "a", false)] // Enum 精确、大小写敏感
    [InlineData("AB", "A", false)] // 精确非包含
    public void Matches_Enum_Exact(string cell, string filter, bool expected)
    {
        Assert.Equal(expected, ColumnFilterPredicate.Matches(cell, filter, ColumnType.Enum));
    }

    // ── 单格匹配：Date 走文本包含（P3.2 既有行为）──
    [Theory]
    [InlineData("2026-07-28", "2026", true)]
    [InlineData("2026-07-28", "08", false)]
    public void Matches_Date_TextContains(string cell, string filter, bool expected)
    {
        Assert.Equal(expected, ColumnFilterPredicate.Matches(cell, filter, ColumnType.Date));
    }

    // ── Build：多列 AND 谓词，走 ColumnStore 列级访问 ──
    [Fact]
    public void Build_MultiColumn_AndSemantics_OverColumnStore()
    {
        var store = ColumnStore.Create(["A", "B", "C"], 3);
        for (var r = 0; r < 3; r++)
            store.AppendRow();
        // row0: apple / 100 / x
        store.SetCellQuiet(0, 0, "apple");
        store.SetCellQuiet(0, 1, "100");
        store.SetCellQuiet(0, 2, "x");
        // row1: apple / 50 / y
        store.SetCellQuiet(1, 0, "apple");
        store.SetCellQuiet(1, 1, "50");
        store.SetCellQuiet(1, 2, "y");
        // row2: banana / 100 / z
        store.SetCellQuiet(2, 0, "banana");
        store.SetCellQuiet(2, 1, "100");
        store.SetCellQuiet(2, 2, "z");

        // 筛选：A contains "app" AND B >= 100
        var filters = new List<(int Col, string Value, ColumnType Type)>
        {
            (0, "app", ColumnType.Text),
            (1, ">=100", ColumnType.Integer),
        };
        var predicate = ColumnFilterPredicate.Build(store, filters);

        Assert.True(predicate(0)); // apple + 100 ✓
        Assert.False(predicate(1)); // apple + 50 ✗ (B<100)
        Assert.False(predicate(2)); // banana ✗ (A)
    }

    [Fact]
    public void Build_EmptyFilters_MatchesAll()
    {
        var store = ColumnStore.Create(["A"], 2);
        store.AppendRow();
        store.AppendRow();
        store.SetCellQuiet(0, 0, "x");
        store.SetCellQuiet(1, 0, "y");

        var predicate = ColumnFilterPredicate.Build(store, new List<(int, string, ColumnType)>());

        Assert.True(predicate(0));
        Assert.True(predicate(1));
    }

    [Fact]
    public void Build_NullCell_TreatedAsEmpty()
    {
        var store = ColumnStore.Create(["A"], 1);
        store.AppendRow(); // A0 = null (未 set)

        var predicate = ColumnFilterPredicate.Build(store, [(0, "x", ColumnType.Text)]);

        Assert.False(predicate(0)); // null 格不含 "x"
    }
}
