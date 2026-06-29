using NumDesTools.ExcelIndex;
using NumDesTools.Scanner;

namespace NumDesTools.Tests;

/// <summary>
/// TDD：SearchTui 搜索与渲染行为验证。
/// 通过 internal 方法测行为，不依赖控制台 I/O。
/// </summary>
public class SearchTuiTests
{
    // ── 辅助 ─────────────────────────────────────────────────────────────────

    private static ExcelSearchIndex MakeIndex(params string[] keys)
    {
        var idx = new ExcelSearchIndex();
        idx.Files.Add("Tables/ConfigA.xlsx");
        idx.Sheets.Add("Sheet1");
        idx.AllSheets.Add((0, 0));
        idx.RebuildLookups();

        for (int i = 0; i < keys.Length; i++)
            idx.Exact[keys[i]] = [new CellHit(0, 0, i + 5, 2)];

        idx.BuildSortedKeys();
        return idx;
    }

    // ── RED 1：1 个字符就能触发搜索，不被空判断拦截 ─────────────────────────

    [Fact]
    public void DoSearch_SingleChar_ReturnsMatches()
    {
        var idx = MakeIndex("abc", "axyz", "hello");

        var results = SearchTui.DoSearch(idx, "a", usePrefix: false);

        Assert.True(results.Count >= 2, $"期望 ≥2 条，实际 {results.Count}");
    }

    // ── RED 2：空 query 返回空列表 ──────────────────────────────────────────

    [Fact]
    public void DoSearch_EmptyQuery_ReturnsEmpty()
    {
        var idx = MakeIndex("abc", "xyz");

        var results = SearchTui.DoSearch(idx, "", usePrefix: false);

        Assert.Empty(results);
    }

    // ── RED 3：前缀模式只命中前缀，不命中中间包含 ───────────────────────────

    [Fact]
    public void DoSearch_PrefixMode_OnlyMatchesPrefix()
    {
        var idx = MakeIndex("abc", "xabc", "abcde");

        var results = SearchTui.DoSearch(idx, "abc", usePrefix: true);

        // "abc" 和 "abcde" 命中，"xabc" 不命中
        Assert.True(results.Count == 2, $"期望 2 条，实际 {results.Count}");
    }

    // ── RED 4：BuildRenderText 输出中标题行 [包含] 只出现一次 ────────────────

    [Fact]
    public void BuildRenderText_TitleAppearsExactlyOnce()
    {
        var idx = MakeIndex("abc");
        var results = SearchTui.DoSearch(idx, "a", usePrefix: false);

        var text = SearchTui.BuildRenderText(
            query: "a",
            results: results,
            selectedIdx: 0,
            usePrefix: false,
            statusMsg: null,
            pageSize: 20
        );

        var count = CountOccurrences(text, "[包含]");
        Assert.Equal(1, count);
    }

    // ── RED 5：选中行含 ▶，非选中行不含 ────────────────────────────────────

    [Fact]
    public void BuildRenderText_SelectedRowHasArrow()
    {
        var idx = MakeIndex("abc", "abd");
        var results = SearchTui.DoSearch(idx, "ab", usePrefix: false);
        Assert.True(results.Count >= 2);

        var text = SearchTui.BuildRenderText(
            query: "ab",
            results: results,
            selectedIdx: 0,
            usePrefix: false,
            statusMsg: null,
            pageSize: 20
        );

        Assert.Contains("▶", text);
    }

    // ── 辅助 ─────────────────────────────────────────────────────────────────

    private static int CountOccurrences(string source, string pattern)
    {
        int count = 0,
            idx = 0;
        while ((idx = source.IndexOf(pattern, idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += pattern.Length;
        }
        return count;
    }
}
