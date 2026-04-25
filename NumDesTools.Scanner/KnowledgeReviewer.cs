using Newtonsoft.Json;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

/// <summary>
/// 知识库持久化 + 回读闭环。
/// 对应 Python 版本的 load_knowledge / save_knowledge / review_human_comments / run_review_pass。
/// </summary>
public class KnowledgeReviewer
{
    private readonly string _knowledgePath;
    private readonly string _reviewedPath;
    private static readonly Regex RxImage = new(@"\[图片\]\s*", RegexOptions.Compiled);
    private static readonly Regex RxTypeFromComment = new(@"type=(\d+)", RegexOptions.Compiled);

    public KnowledgeReviewer(string configDir)
    {
        _knowledgePath = Path.Combine(configDir, "knowledge_base.json");
        _reviewedPath  = Path.Combine(configDir, "reviewed_state.json");
    }

    // ── 持久化 ───────────────────────────────────────────────────────────────

    public KnowledgeBase LoadKnowledge()
    {
        if (!File.Exists(_knowledgePath)) return new KnowledgeBase();
        try { return JsonConvert.DeserializeObject<KnowledgeBase>(File.ReadAllText(_knowledgePath)) ?? new(); }
        catch { return new KnowledgeBase(); }
    }

    public void SaveKnowledge(KnowledgeBase kb)
        => File.WriteAllText(_knowledgePath, JsonConvert.SerializeObject(kb, Formatting.Indented));

    public ReviewedState LoadReviewed()
    {
        if (!File.Exists(_reviewedPath)) return new ReviewedState();
        try { return JsonConvert.DeserializeObject<ReviewedState>(File.ReadAllText(_reviewedPath)) ?? new(); }
        catch { return new ReviewedState(); }
    }

    public void SaveReviewed(ReviewedState reviewed)
        => File.WriteAllText(_reviewedPath, JsonConvert.SerializeObject(reviewed, Formatting.Indented));

    // ── 单条回读 ─────────────────────────────────────────────────────────────

    public async Task<int> ReviewItemAsync(
        WorkItem item, List<string> typeNums,
        KnowledgeBase kb, ReviewedState reviewed)
    {
        // AI分析写在正文，人工备注在评论里回读
        bool hasAi = item.Desc.Contains(CommentBuilder.StoryMarker)
                  || item.Desc.Contains(CommentBuilder.IssueMarker);
        if (!hasAi) return 0;

        List<(string Id, string Content)> comments;
        try { comments = await FeishuWorkItemFetcher.GetCommentsAsync(item.Id); }
        catch { return 0; }

        reviewed.TryGetValue(item.Id, out var seenList);
        var seen = (seenList ?? []).ToHashSet();
        var newNotes = new List<string>();

        foreach (var (cid, content) in comments)
        {
            if (seen.Contains(cid)) continue;

            var clean = RxImage.Replace(content, "").Trim();
            if (string.IsNullOrEmpty(clean)) continue;

            newNotes.Add(clean);
            seen.Add(cid);
        }

        if (newNotes.Count == 0) return 0;

        var itemId = item.Id;
        var itemName = item.Name;

        foreach (var tNum in typeNums.DefaultIfEmpty("unknown"))
        {
            var key     = $"type_{tNum}";
            var entries = kb.GetEntries(key);
            var existing = entries.FirstOrDefault(e => e.SourceId == itemId);
            if (existing != null)
            {
                foreach (var note in newNotes)
                    if (!existing.HumanNotes.Contains(note))
                        existing.HumanNotes.Add(note);
                existing.UpdatedAt = DateTime.Today.ToString("yyyy-MM-dd");
            }
            else
            {
                entries.Add(new KnowledgeEntry
                {
                    SourceId   = itemId,
                    SourceName = itemName,
                    HumanNotes = newNotes,
                    LearnedAt  = DateTime.Today.ToString("yyyy-MM-dd"),
                    UpdatedAt  = DateTime.Today.ToString("yyyy-MM-dd"),
                });
            }
            kb.SetEntries(key, entries);
        }

        reviewed[itemId] = seen.ToList();
        return newNotes.Count;
    }

    // ── 全量回读 ─────────────────────────────────────────────────────────────

    public async Task RunReviewPassAsync(
        ActivityTableRules rules, ActivityTypeIndex typeIndex)
    {
        Console.WriteLine($"\n[{Now()}] 开始知识回读（已有AI分析的条目）...");
        var kb       = LoadKnowledge();
        var reviewed = LoadReviewed();
        int total    = 0;

        var stories = await FeishuWorkItemFetcher.FetchStoriesAsync(fetchAll: true);
        // AI分析已写在正文，直接用 Desc 判断，无需额外网络请求
        var candidates = stories
            .Where(s => FeishuWorkItemFetcher.HasExistingAiDescription(s, CommentBuilder.StoryMarker))
            .ToList();

        Console.WriteLine($"  找到 {candidates.Count} 条已有AI分析的工作项，开始回读...");

        foreach (var s in candidates)
        {
            var text     = s.Name + " " + s.Desc;
            var typeNums = TableMatcher.MatchTypesFromText(text, typeIndex.TypeIndex);

            // 若 typeIndex 不可用，从正文中反解 type
            if (typeNums.Count == 0)
                typeNums = RxTypeFromComment.Matches(s.Desc)
                    .Select(m => m.Groups[1].Value).Distinct().ToList();

            int n = await ReviewItemAsync(s, typeNums, kb, reviewed);
            if (n > 0)
            {
                Console.WriteLine($"  [学到 {n} 条] {s.Name[..Math.Min(40, s.Name.Length)]}");
                total += n;
            }
        }

        SaveKnowledge(kb);
        SaveReviewed(reviewed);
        Console.WriteLine($"[{Now()}] 知识回读完成，本次新学 {total} 条。");
    }

    // ── 知识库查询 ────────────────────────────────────────────────────────────

    public List<string> QueryKnowledge(List<string> typeNums, KnowledgeBase kb)
    {
        var seen   = new HashSet<string>();
        var result = new List<string>();

        foreach (var tNum in typeNums)
        {
            var key = $"type_{tNum}";
            foreach (var entry in kb.GetEntries(key))
            {
                foreach (var note in entry.HumanNotes)
                {
                    var item = $"[{entry.SourceName[..Math.Min(20, entry.SourceName.Length)]}] {note}";
                    if (seen.Add(item)) result.Add(item);
                    if (result.Count >= 5) return result;
                }
            }
        }
        return result;
    }

    private static string Now() => DateTime.Now.ToString("HH:mm:ss");
}
