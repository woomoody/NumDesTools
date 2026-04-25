using NumDesTools.Scanner;

namespace NumDesTools.Tests;

public class ScannerTests
{
    // ── TableMatcher.MatchTypesFromText ───────────────────────────────────────

    [Fact]
    public void MatchTypes_ExplicitTypeTag_ReturnsCorrectType()
    {
        var types = TableMatcher.MatchTypesFromText("type=2 LTE活动第1期", []);
        Assert.Contains("2", types);
        Assert.Equal("2", types[0]); // 显式 type 优先插到头部
    }

    [Fact]
    public void MatchTypes_NoteSubstringInText_Matches()
    {
        var index = new List<(string, int)> { ("家装lte", 2), ("对对碰", 73) };
        var types = TableMatcher.MatchTypesFromText("家装玩法-LTE-第9期春日记忆", index);
        Assert.Contains("2", types);
    }

    [Fact]
    public void MatchTypes_UnrelatedText_ReturnsEmpty()
    {
        var index = new List<(string, int)> { ("家装lte", 2) };
        var types = TableMatcher.MatchTypesFromText("UI布局调整需求", index);
        Assert.Empty(types);
    }

    [Fact]
    public void MatchTypes_DeprecatedNote_Skipped()
    {
        var index = new List<(string, int)> { ("家装lte废弃", 2) };
        var types = TableMatcher.MatchTypesFromText("家装lte废弃活动", index);
        Assert.Empty(types);
    }

    // ── FeishuWorkItemFetcher.ExtractActivityIds ──────────────────────────────

    [Fact]
    public void ExtractActivityIds_PureNumber_Found()
    {
        var ids = FeishuWorkItemFetcher.ExtractActivityIds("【800909】不完成任务导致卡死");
        Assert.Single(ids);
        Assert.Equal("800909", ids[0]);
    }

    [Fact]
    public void ExtractActivityIds_PrefixText_Found()
    {
        var ids = FeishuWorkItemFetcher.ExtractActivityIds("【闯关LTE763401】进入报错");
        Assert.Single(ids);
        Assert.Equal("763401", ids[0]);
    }

    [Fact]
    public void ExtractActivityIds_NoId_ReturnsEmpty()
    {
        var ids = FeishuWorkItemFetcher.ExtractActivityIds("UI顶部穿帮问题");
        Assert.Empty(ids);
    }

    [Fact]
    public void ExtractActivityIds_TooShort_NotMatched()
    {
        // 少于4位不应匹配
        var ids = FeishuWorkItemFetcher.ExtractActivityIds("【123】三位数");
        Assert.Empty(ids);
    }

    // ── CommentBuilder.BuildStoryComment ─────────────────────────────────────

    [Fact]
    public void BuildStoryComment_ContainsMarkerAndTableName()
    {
        var story  = new WorkItem("1", "测试需求", "", "");
        var tables = new List<TableMatch>
        {
            new("ActivityClientData.xlsx", "活动主表", ["id", "type"], "id", null),
        };
        var comment = CommentBuilder.BuildStoryComment(story, tables, 1);
        Assert.Contains(CommentBuilder.StoryMarker, comment);
        Assert.Contains("ActivityClientData.xlsx", comment);
    }

    [Fact]
    public void BuildStoryComment_Phase2_ShowsFollowupHeader()
    {
        var story  = new WorkItem("1", "续期需求第2期", "", "");
        var tables = new List<TableMatch> { new("ActivityClientData.xlsx", "主表", [], "id", null) };
        var comment = CommentBuilder.BuildStoryComment(story, tables, 2);
        Assert.Contains("续期需求", comment);
        Assert.DoesNotContain("必填字段", comment);
    }

    [Fact]
    public void BuildStoryComment_WithHistoryNotes_IncludesNotes()
    {
        var story  = new WorkItem("1", "需求", "", "");
        var tables = new List<TableMatch> { new("Foo.xlsx", "", [], "id", "2") };
        var notes  = new List<string> { "多语言需求已添加" };
        var comment = CommentBuilder.BuildStoryComment(story, tables, 1, notes);
        Assert.Contains("多语言需求已添加", comment);
        Assert.Contains("历史经验", comment);
    }

    // ── CommentBuilder.BuildIssueComment ─────────────────────────────────────

    [Fact]
    public void BuildIssueComment_ContainsIssueMarkerAndActivityId()
    {
        var issue  = new WorkItem("9", "【800909】兑换点icon不同", "", "待处理");
        var tables = new List<TableMatch> { new("LteExchangeData.xlsx", "", [], "activityID", "2") };
        var comment = CommentBuilder.BuildIssueComment(issue, tables, "800909", 2, "家装LTE");
        Assert.Contains(CommentBuilder.IssueMarker, comment);
        Assert.Contains("800909", comment);
        Assert.Contains("type=2", comment);
    }

    [Fact]
    public void InferBugHints_FindKeyword_Matched()
    {
        var (excels, fix, matched) = CommentBuilder.InferBugHints("【763401】寻找不到目标");
        Assert.True(matched);
        Assert.Contains("FindTargetTemplateData.xlsx", excels);
    }

    [Fact]
    public void InferBugHints_UnknownKeyword_NotMatched()
    {
        var (_, _, matched) = CommentBuilder.InferBugHints("【763301】图九去掉三合一标识");
        Assert.False(matched);
    }

    // ── CommentBuilder.BuildIssueComment with findChain ──────────────────────

    [Fact]
    public void BuildIssueComment_WithFindChain_IncludesChainSection()
    {
        var issue   = new WorkItem("9", "【800909】寻找不到目标", "", "待处理");
        var tables  = new List<TableMatch> { new("FindTargetTemplateData.xlsx", "", [], "id", "2") };
        var comment = CommentBuilder.BuildIssueComment(issue, tables, "800909", 2, "家装LTE",
            findChainAnalysis: "  物品/目标 ID=42011604\n    └ FindTargetTemplateData[42011604].findTargets = ...");
        Assert.Contains("寻找链分析", comment);
        Assert.Contains("42011604", comment);
    }

    [Fact]
    public void BuildIssueComment_NoFindChain_NoChainSection()
    {
        var issue   = new WorkItem("9", "【800909】icon不同", "", "待处理");
        var tables  = new List<TableMatch> { new("LteExchangeData.xlsx", "", [], "activityID", "2") };
        var comment = CommentBuilder.BuildIssueComment(issue, tables, "800909", 2, "家装LTE",
            findChainAnalysis: null);
        Assert.DoesNotContain("寻找链分析", comment);
    }

    // ── KnowledgeReviewer.QueryKnowledge ─────────────────────────────────────

    [Fact]
    public void QueryKnowledge_ReturnsMatchingNotes()
    {
        var kb = new KnowledgeBase();
        kb.SetEntries("type_2", [new KnowledgeEntry
        {
            SourceId   = "999",
            SourceName = "测试活动",
            HumanNotes = ["多语言需求已添加", "节点点一下完成"],
        }]);

        var reviewer = new KnowledgeReviewer(@"C:\nonexistent");
        var notes    = reviewer.QueryKnowledge(["2"], kb);
        Assert.Equal(2, notes.Count);
        Assert.Contains(notes, n => n.Contains("多语言需求已添加"));
    }

    [Fact]
    public void QueryKnowledge_UnknownType_ReturnsEmpty()
    {
        var kb    = new KnowledgeBase();
        var rev   = new KnowledgeReviewer(@"C:\nonexistent");
        var notes = rev.QueryKnowledge(["999"], kb);
        Assert.Empty(notes);
    }
}
