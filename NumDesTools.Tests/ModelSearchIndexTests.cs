using NumDesTools;
using NumDesTools.ExcelIndex;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// 测试 SearchModelKeyMiniExcelMulti 的索引聚合逻辑（纯内存，无需打开文件）
/// </summary>
public class ModelSearchIndexTests
{
    // ── 辅助：构建最小可用的 ExcelSearchIndex ────────────────────────────────

    private static ExcelSearchIndex BuildIndex(
        string excelsRoot,
        IEnumerable<(string relPath, string sheet, string value, int row, int col)> entries)
    {
        var idx = new ExcelSearchIndex();
        idx.RebuildLookups();
        var knownPairs = new System.Collections.Generic.HashSet<(int, int)>();

        foreach (var (relPath, sheet, value, row, col) in entries)
        {
            if (!idx.FileIds.TryGetValue(relPath, out var fid))
            {
                fid = idx.Files.Count;
                idx.Files.Add(relPath);
                idx.FileIds[relPath] = fid;
            }
            if (!idx.SheetIds.TryGetValue(sheet, out var sid))
            {
                sid = idx.Sheets.Count;
                idx.Sheets.Add(sheet);
                idx.SheetIds[sheet] = sid;
            }
            if (knownPairs.Add((fid, sid)))
                idx.AllSheets.Add((fid, sid));

            var hit = new CellHit(fid, sid, row, col);
            if (!idx.Exact.TryGetValue(value, out var list))
            {
                list = new System.Collections.Generic.List<CellHit>();
                idx.Exact[value] = list;
            }
            list.Add(hit);
        }
        return idx;
    }

    // ── P2 Test 1：SheetName 精确搜索 → 返回对应 (file, sheet) ───────────────

    [Fact]
    public void SearchSheetName_ExactMatch_ReturnsFileAndSheet()
    {
        var root = @"C:\Project\Excels";
        var idx = BuildIndex(root, new[]
        {
            ("Tables/ItemData.xlsx", "c_item",  "1", 5, 2),
            ("Tables/DropData.xlsx", "c_drop",  "2", 5, 2),
            ("Tables/ItemData.xlsx", "c_merge", "3", 6, 2),
        });

        var result = ExcelIndex.ExcelIndexManager.SearchSheetNameFromIndex(
            "c_item", isContains: false, idx, root);

        Assert.Single(result);
        Assert.Contains("ItemData.xlsx", result[0].file);
        Assert.Equal("c_item", result[0].sheet);
    }

    // ── P2 Test 2：SheetName 包含搜索 → 多个匹配 ─────────────────────────────

    [Fact]
    public void SearchSheetName_ContainsMatch_ReturnsMultiple()
    {
        var root = @"C:\Project\Excels";
        var idx = BuildIndex(root, new[]
        {
            ("Tables/ItemData.xlsx", "c_item",  "1", 5, 2),
            ("Tables/DropData.xlsx", "c_item_drop", "2", 5, 2),
        });

        var result = ExcelIndex.ExcelIndexManager.SearchSheetNameFromIndex(
            "c_item", isContains: true, idx, root);

        Assert.Equal(2, result.Count);
    }

    // ── P3 Test：前缀搜索 → 返回所有以 prefix 开头的命中 ────────────────────

    [Fact]
    public void PrefixSearch_ReturnAllMatchingKeys()
    {
        var root = @"C:\Project\Excels";
        var idx = BuildIndex(root, new[]
        {
            ("Tables/a.xlsx", "Sheet1", "7632010101", 5, 2),
            ("Tables/a.xlsx", "Sheet1", "7632010102", 6, 2),
            ("Tables/b.xlsx", "Sheet1", "7632020001", 7, 2),
            ("Tables/a.xlsx", "Sheet1", "99990001",   8, 2),
        });
        idx.BuildSortedKeys(); // 构建前缀查询用的有序数组

        var result = ExcelIndex.ExcelIndexManager.SearchByPrefix("763201", idx, root);

        // 763201 开头的：7632010101, 7632010102
        Assert.Equal(2, result.Count);
        Assert.All(result, r => Assert.True(
            r.file.Contains("a.xlsx") || r.file.Contains("b.xlsx")));
    }

    // ── Test 3：无命中 → 返回空 dict ─────────────────────────────────────────

    [Fact]
    public void NoHits_ReturnsEmptyDict()
    {
        var idx = BuildIndex(@"C:\Root", System.Array.Empty<(string,string,string,int,int)>());
        var result = PubMetToExcelFunc.BuildModelResultFromIndex(new[] { "99999" }, idx, @"C:\Root");
        Assert.Empty(result);
    }

    // ── Test 4：含 * 的 ID 不经此方法（调用方 fallback）────────────────────

    [Fact]
    public void PrefixId_IsNotInIndex_ReturnsEmpty()
    {
        // 索引里存的是原始值 "10001"，带 * 的 "*10001" 不会命中
        var idx = BuildIndex(@"C:\Root", new[] { ("Tables/a.xlsx", "Sheet1", "10001", 5, 2) });
        var result = PubMetToExcelFunc.BuildModelResultFromIndex(new[] { "*10001" }, idx, @"C:\Root");
        Assert.Empty(result);  // 索引里没有 "*10001" 这个 key
    }

    // ── Test 2：含 $ 文件名 → key 用 filename#sheetName ─────────────────────

    [Fact]
    public void DollarSignFileName_KeyIncludesSheetName()
    {
        var root = @"C:\Project\Excels";
        var idx = BuildIndex(root, new[]
        {
            ("Tables/Multi$Sheet.xlsx", "c_item", "20001", 5, 2),
        });

        var result = PubMetToExcelFunc.BuildModelResultFromIndex(new[] { "20001" }, idx, root);

        Assert.Contains("Multi$Sheet.xlsx#c_item", result.Keys);
        Assert.DoesNotContain("Multi$Sheet.xlsx", result.Keys);
    }

    // ── Test 1：精确 ID 命中 → 结果按 filename 聚合 ──────────────────────────

    [Fact]
    public void ExactIds_HitsGroupedByFileName()
    {
        var root = @"C:\Project\Excels";
        var idx = BuildIndex(root, new[]
        {
            ("Tables/ItemData.xlsx", "Sheet1", "10001", 5, 2),
            ("Tables/ItemData.xlsx", "Sheet1", "10002", 6, 2),
            ("Tables/DropData.xlsx", "Sheet1", "10001", 8, 2),
        });

        var ids = new[] { "10001", "10002" };

        var result = PubMetToExcelFunc.BuildModelResultFromIndex(ids, idx, root);

        // 10001 出现在两个文件里
        Assert.Contains("ItemData.xlsx", result.Keys);
        Assert.Contains("DropData.xlsx", result.Keys);
        Assert.Contains("10001", result["ItemData.xlsx"]);
        Assert.Contains("10002", result["ItemData.xlsx"]);
        Assert.Contains("10001", result["DropData.xlsx"]);
    }
}
