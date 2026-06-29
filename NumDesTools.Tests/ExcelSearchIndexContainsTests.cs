using NumDesTools.ExcelIndex;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// TDD：验证 SearchByContains 正确性，驱动 Span+SortedKeys 优化实现。
/// </summary>
public class ExcelSearchIndexContainsTests
{
    // ── 辅助：快速构建带若干 key 的索引 ─────────────────────────────────────

    private static ExcelSearchIndex MakeIndex(params string[] keys)
    {
        var idx = new ExcelSearchIndex();
        idx.Files.Add("a.xlsx");
        idx.Sheets.Add("Sheet1");
        idx.AllSheets.Add((0, 0));
        idx.RebuildLookups();

        for (int i = 0; i < keys.Length; i++)
            idx.Exact[keys[i]] = [new CellHit(0, 0, i + 5, 2)];

        return idx;
    }

    // ── RED 1：基础包含匹配 ────────────────────────────────────────────────

    [Fact]
    public void Contains_KeywordInMiddle_ReturnsHit()
    {
        var idx = MakeIndex("10001_item", "20002_weapon", "30003_armor");

        var hits = idx.SearchByContains("item", StringComparison.OrdinalIgnoreCase);

        Assert.Single(hits);
    }

    // ── RED 2：大小写不敏感 ────────────────────────────────────────────────

    [Fact]
    public void Contains_CaseInsensitive_MatchesUpperLower()
    {
        var idx = MakeIndex("ItemData", "ITEMDATA2", "other");

        var hits = idx.SearchByContains("itemdata", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(2, hits.Count);
    }

    // ── RED 3：无命中返回空 ────────────────────────────────────────────────

    [Fact]
    public void Contains_NoMatch_ReturnsEmpty()
    {
        var idx = MakeIndex("abc", "def", "ghi");

        var hits = idx.SearchByContains("xyz", StringComparison.OrdinalIgnoreCase);

        Assert.Empty(hits);
    }

    // ── RED 4：cap 限制生效 ────────────────────────────────────────────────

    [Fact]
    public void Contains_ExceedsCap_ReturnsCappedResults()
    {
        // 10 个 key 都含 "x"，每 key 1 条 hit，cap=3
        var keys = Enumerable.Range(1, 10).Select(i => $"x{i:D4}").ToArray();
        var idx = MakeIndex(keys);

        var hits = idx.SearchByContains("x", StringComparison.Ordinal, maxCap: 3);

        Assert.True(hits.Count <= 3);
    }

    // ── RED 5：多 hit per key（一个 key 对应多行命中）────────────────────

    [Fact]
    public void Contains_MultiHitPerKey_AllHitsReturned()
    {
        var idx = new ExcelSearchIndex();
        idx.Files.Add("a.xlsx");
        idx.Sheets.Add("Sheet1");
        idx.AllSheets.Add((0, 0));
        idx.RebuildLookups();
        // key "abc" 出现在第 5、6、7 行
        idx.Exact["abc"] =
        [
            new CellHit(0, 0, 5, 2),
            new CellHit(0, 0, 6, 2),
            new CellHit(0, 0, 7, 2),
        ];
        idx.Exact["xyz"] = [new CellHit(0, 0, 8, 2)];

        var hits = idx.SearchByContains("abc", StringComparison.Ordinal);

        Assert.Equal(3, hits.Count);
    }

    // ── RED 6：BuildSortedKeys 之后 Contains 结果与之前一致 ───────────────

    [Fact]
    public void Contains_AfterBuildSortedKeys_SameResultAsBefore()
    {
        var idx = MakeIndex("apple_01", "pineapple_02", "banana_03", "grape_04");

        var beforeSort = idx.SearchByContains("apple", StringComparison.OrdinalIgnoreCase);
        idx.BuildSortedKeys();
        var afterSort = idx.SearchByContains("apple", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(beforeSort.Count, afterSort.Count);
    }

    // ── RED 7：数字 key（游戏配置 ID 场景）────────────────────────────────

    [Fact]
    public void Contains_NumericKeys_PrefixFragment_ReturnsMatches()
    {
        var idx = MakeIndex("7632010101", "7632010102", "7632020001", "9999001");
        idx.BuildSortedKeys();

        var hits = idx.SearchByContains("763201", StringComparison.Ordinal);

        Assert.Equal(2, hits.Count);
    }

    // ── RED 8：空索引不报错 ────────────────────────────────────────────────

    [Fact]
    public void Contains_EmptyIndex_ReturnsEmpty()
    {
        var idx = new ExcelSearchIndex();
        idx.BuildSortedKeys();

        var hits = idx.SearchByContains("anything", StringComparison.Ordinal);

        Assert.Empty(hits);
    }
}
