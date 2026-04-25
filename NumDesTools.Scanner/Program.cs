using Newtonsoft.Json;
using System.Text;

namespace NumDesTools.Scanner;

/// <summary>
/// 入口。命令行参数与 Python 版本一致：
///   --all      全量扫描需求，分析完即退出
///   --bugs     缺陷分析模式
///   --review   仅执行知识回读
///   --release  发布模式（可写入飞书）
///   --confirm  自动确认写入（跳过交互）
///   --item X,Y 指定工作项ID
/// </summary>
internal class Program
{
    // ── 配置路径 ─────────────────────────────────────────────────────────────
    private const string ConfigDir       = @"C:\Users\cent\Documents\NumDesTools\Config";
    private const string ActivityXlsx    = @"C:\M1Work\public\Excels\Tables\ActivityClientData.xlsx";
    private const string RulesPath       = ConfigDir + @"\ActivityTableRules.json";

    private const string McpToken        = "m-002580b6-b3db-405a-aab8-007f927ef4eb";
    private const string ProjectKey      = "t89j73";
    private const string WrittenItemsPath = ConfigDir + @"\written_items.json";

    // 本地已写入记录：workItemId → 写入日期，避免依赖飞书 API 一致性延迟
    private static HashSet<string> LoadWrittenItems()
    {
        if (!File.Exists(WrittenItemsPath)) return [];
        try
        {
            var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                File.ReadAllText(WrittenItemsPath)) ?? [];
            return dict.Keys.ToHashSet();
        }
        catch { return []; }
    }

    private static void SaveWrittenItem(string itemId)
    {
        Dictionary<string, string> dict = [];
        if (File.Exists(WrittenItemsPath))
            try { dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(WrittenItemsPath)) ?? []; }
            catch { }
        dict[itemId] = DateTime.Today.ToString("yyyy-MM-dd");
        File.WriteAllText(WrittenItemsPath, JsonConvert.SerializeObject(dict, Formatting.Indented));
    }

    private static readonly int ConfirmTimeoutSec = 30 * 60;

    static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        // 解析参数
        bool releaseMode = args.Contains("--release");
        bool confirmMode = args.Contains("--confirm");
        bool fetchAll    = args.Contains("--all");
        bool bugsMode    = args.Contains("--bugs");
        bool reviewOnly  = args.Contains("--review");
        bool forceMode   = args.Contains("--force"); // 跳过已有AI评论检测，强制重新分析
        List<string>? itemIds = null;
        int itemIdx = Array.IndexOf(args, "--item");
        if (itemIdx >= 0 && itemIdx + 1 < args.Length)
            itemIds = args[itemIdx + 1].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

        // 初始化飞书客户端
        FeishuMcpClient.McpToken   = McpToken;
        FeishuMcpClient.ProjectKey = ProjectKey;

        // 加载规则 & 索引
        var rules = LoadRules();
        var typeIndex = new ActivityTypeIndex();
        typeIndex.Load(ActivityXlsx);

        var reviewer = new KnowledgeReviewer(ConfigDir);

        // ── 仅回读模式 ───────────────────────────────────────────────────────
        if (reviewOnly)
        {
            await reviewer.RunReviewPassAsync(rules, typeIndex);
            return 0;
        }

        var writtenItems = LoadWrittenItems();

        // ── 缺陷分析模式 ─────────────────────────────────────────────────────
        if (bugsMode)
        {
            Console.WriteLine($"\n[{Now()}] {(itemIds != null ? $"缺陷分析 指定ID {string.Join(",", itemIds)}" : "缺陷全量分析")}...");
            var kb      = reviewer.LoadKnowledge();
            var pending = await ScanIssuesAsync(rules, typeIndex, reviewer, kb, itemIds, fetchAll: true, force: forceMode, writtenItems: writtenItems);
            if (pending.Count == 0) { Console.WriteLine("[INFO] 无待分析缺陷。"); return 0; }

            var docPath = WriteBugDoc(pending);
            Console.WriteLine($"\n[INFO] 缺陷分析文档：{docPath}");

            if (!releaseMode)
            {
                Console.WriteLine("[INFO] 测试模式，以上为分析结果，不执行写入。");
                Console.WriteLine("[INFO] 使用 --release 参数启动可开启写入飞书缺陷评论。");
                return 0;
            }
            await WriteToFeishu(pending, confirmMode, "[缺陷评论]");
            return 0;
        }

        // ── 需求扫描模式（--all 或 --item）──────────────────────────────────
        if (itemIds != null || fetchAll)
        {
            var label = itemIds != null ? $"指定ID {string.Join(",", itemIds)}" : "全量扫描";
            Console.WriteLine($"\n[{Now()}] {label}...");
            var kb      = reviewer.LoadKnowledge();
            var seen    = new HashSet<string>();
            var pending = await ScanStoriesAsync(rules, typeIndex, reviewer, kb, seen, itemIds, fetchAll, forceMode, writtenItems);
            if (pending.Count > 0)
                await PromptAndWrite(pending, releaseMode, confirmMode);
            else
                Console.WriteLine("[INFO] 无待写入项。");
            return 0;
        }

        Console.WriteLine("[INFO] 请指定模式：--all | --bugs | --review | --item <id>");
        return 1;
    }

    // ── 需求扫描 ─────────────────────────────────────────────────────────────

    private static async Task<List<PendingItem>> ScanStoriesAsync(
        ActivityTableRules rules, ActivityTypeIndex typeIndex,
        KnowledgeReviewer reviewer, KnowledgeBase kb,
        HashSet<string> seenIds, List<string>? itemIds, bool fetchAll, bool force = false,
        HashSet<string>? writtenItems = null)
    {
        var stories = await FeishuWorkItemFetcher.FetchStoriesAsync(itemIds, fetchAll);
        Console.WriteLine($"  开始分析 {stories.Count} 条需求...\n");
        var pending = new List<PendingItem>();

        foreach (var s in stories)
        {
            if (itemIds == null && !fetchAll && !seenIds.Add(s.Id)) continue;
            seenIds.Add(s.Id);

            if (!TableMatcher.IsPlannerRelated(s)) continue;

            var (tables, phase) = TableMatcher.IdentifyTables(s, rules, typeIndex.TypeIndex);
            if (tables.Count == 0) { Console.WriteLine($"  未识别到表格：{Truncate(s.Name, 40)}"); continue; }

            if (!force)
            {
                if (writtenItems?.Contains(s.Id) == true)
                { Console.WriteLine($"  已有AI分析，跳过：{Truncate(s.Name, 40)}"); continue; }
                if (await FeishuWorkItemFetcher.HasExistingAiCommentAsync(s, CommentBuilder.StoryMarker))
                { Console.WriteLine($"  已有AI分析，跳过：{Truncate(s.Name, 40)}"); continue; }
            }

            var typeNums     = tables.Where(t => t.TypeNum != null).Select(t => t.TypeNum!).Distinct().ToList();
            var historyNotes = reviewer.QueryKnowledge(typeNums, kb);
            var phaseTag     = phase > 1 ? $"第{phase}期续期" : "首期/全新";

            Console.WriteLine($"  新增待写入[{phaseTag}]：{Truncate(s.Name, 40)}");
            Console.WriteLine($"    涉及表格：{string.Join(", ", tables.Select(t => t.Excel))}");
            if (historyNotes.Count > 0) Console.WriteLine($"    引用历史经验：{historyNotes.Count} 条");

            pending.Add(new PendingItem
            {
                Id           = s.Id,
                Name         = s.Name,
                Tables       = tables.Select(t => t.Excel).ToList(),
                Comment      = CommentBuilder.BuildStoryComment(s, tables, phase, historyNotes),
                ItemType     = "story",
                OriginalDesc = s.Desc,
            });
        }
        return pending;
    }

    // ── 缺陷扫描 ─────────────────────────────────────────────────────────────

    private static async Task<List<PendingItem>> ScanIssuesAsync(
        ActivityTableRules rules, ActivityTypeIndex typeIndex,
        KnowledgeReviewer reviewer, KnowledgeBase kb,
        List<string>? itemIds, bool fetchAll, bool force = false,
        HashSet<string>? writtenItems = null)
    {
        var issues = await FeishuWorkItemFetcher.FetchIssuesAsync(itemIds, fetchAll);
        var findAnalyzer = new FindChainAnalyzer();
        Console.WriteLine($"  开始分析 {issues.Count} 条缺陷...\n");
        var pending = new List<PendingItem>();

        foreach (var issue in issues)
        {
            var actIds = FeishuWorkItemFetcher.ExtractActivityIds(issue.Name);
            if (actIds.Count == 0) actIds = FeishuWorkItemFetcher.ExtractActivityIds(issue.Desc);
            if (actIds.Count == 0) { Console.WriteLine($"  未提取到活动ID，跳过：{Truncate(issue.Name, 40)}"); continue; }

            var actId = actIds[0];
            var (typeNum, typeNote) = typeIndex.LookupByActivityId(actId);

            var fakeItem = new WorkItem(issue.Id, issue.Name,
                (typeNum.HasValue ? $"type={typeNum} " : "") + issue.Desc, issue.Status);
            var (tables, _) = TableMatcher.IdentifyTables(fakeItem, rules, typeIndex.TypeIndex);

            if (tables.Count == 0 && typeNum.HasValue)
            {
                var mainDef = rules.Tables.FirstOrDefault(t => t.Name == "ActivityClientData");
                if (mainDef != null)
                    tables = [new TableMatch(mainDef.ExcelFile, mainDef.Desc, [], mainDef.KeyField, null)];
            }
            if (tables.Count == 0) { Console.WriteLine($"  未识别到相关表格，跳过：{Truncate(issue.Name, 40)}"); continue; }

            if (!force)
            {
                if (writtenItems?.Contains(issue.Id) == true)
                { Console.WriteLine($"  已有AI缺陷分析，跳过：{Truncate(issue.Name, 40)}"); continue; }
                if (await FeishuWorkItemFetcher.HasExistingAiCommentAsync(issue, CommentBuilder.IssueMarker))
                { Console.WriteLine($"  已有AI缺陷分析，跳过：{Truncate(issue.Name, 40)}"); continue; }
            }

            var typeNums     = typeNum.HasValue ? [typeNum.Value.ToString()] : new List<string>();
            var historyNotes = reviewer.QueryKnowledge(typeNums, kb);
            var findChain    = findAnalyzer.Analyze(issue.Name, issue.Desc);
            var (_, _, hintMatched) = CommentBuilder.InferBugHints(issue.Name);

            if (hintMatched)
                Console.WriteLine($"  新增缺陷待写入：{Truncate(issue.Name, 40)}");
            else
                Console.WriteLine($"  新增缺陷（仅报告，关键词未命中）：{Truncate(issue.Name, 40)}");
            Console.WriteLine($"    activityID={actId}, type={typeNum?.ToString() ?? "?"}");
            if (historyNotes.Count > 0) Console.WriteLine($"    引用历史经验：{historyNotes.Count} 条");
            if (findChain != null) Console.WriteLine($"    寻找链分析：已提取到ID追溯");

            pending.Add(new PendingItem
            {
                Id           = issue.Id,
                Name         = issue.Name,
                Tables       = tables.Select(t => t.Excel).ToList(),
                Comment      = CommentBuilder.BuildIssueComment(issue, tables, actId, typeNum, typeNote, historyNotes, findChain),
                ItemType     = "issue",
                SkipComment  = !hintMatched,
                SkipReason   = hintMatched ? "" : "标题关键词未匹配到已知配置规则",
                OriginalDesc = issue.Desc,
            });
        }
        return pending;
    }

    // ── 写入 ─────────────────────────────────────────────────────────────────

    private static async Task PromptAndWrite(List<PendingItem> pending, bool releaseMode, bool confirmMode)
    {
        Console.WriteLine();
        Console.WriteLine(new string('=', 60));
        Console.WriteLine(releaseMode
            ? "【发布模式】待写入评论摘要（输入 y 执行写入）："
            : "【测试模式】分析结果预览（不会写入飞书）：");
        Console.WriteLine(new string('=', 60));
        for (int i = 0; i < pending.Count; i++)
        {
            var item = pending[i];
            Console.WriteLine($"  {i + 1}. [{item.Id}] {Truncate(item.Name, 50)}");
            Console.WriteLine($"     涉及表格：{string.Join(", ", item.Tables)}");
        }
        Console.WriteLine(new string('=', 60));

        var today    = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath  = Path.Combine(ConfigDir, $"pending_comments_{today}.md");
        WriteReviewDoc(pending, docPath);
        Console.WriteLine($"\n[INFO] 审阅文档：{docPath}");

        if (!releaseMode)
        {
            Console.WriteLine("[INFO] 测试模式，以上为分析结果，不执行写入。");
            Console.WriteLine("[INFO] 使用 --release 参数启动可开启写入。");
            return;
        }
        await WriteToFeishu(pending, confirmMode, "[需求评论]");
    }

    private static async Task WriteToFeishu(List<PendingItem> pending, bool confirmMode, string label)
    {
        string answer;
        if (confirmMode)
        {
            Console.WriteLine("\n[INFO] --confirm 参数已传入，自动确认写入。");
            answer = "y";
        }
        else
        {
            Console.Write($"\n是否写入飞书{label}？(y/N)  [30分钟内无回复自动跳过]: ");
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(ConfirmTimeoutSec));
            string? input = null;
            try
            {
                var readTask = Task.Run(() => Console.ReadLine(), cts.Token);
                input = await readTask;
            }
            catch (OperationCanceledException) { Console.WriteLine("\n[INFO] 超时，本次不写入。"); return; }
            answer = (input ?? "").Trim().ToLower();
        }

        if (answer != "y") { Console.WriteLine("[INFO] 已取消。"); return; }

        int skipped = pending.Count(p => p.SkipComment);
        if (skipped > 0)
            Console.WriteLine($"[INFO] 跳过 {skipped} 条关键词未命中的缺陷（仅报告不写飞书）");

        Console.WriteLine("\n[INFO] 开始写入飞书正文...");
        int ok = 0, fail = 0;

        foreach (var item in pending)
        {
            if (item.SkipComment) continue;
            var typeKey = item.ItemType == "story" ? "story" : "issue";
            try
            {
                var fieldKey = typeKey == "story"
                    ? FeishuMcpClient.StoryContentFieldKey
                    : FeishuMcpClient.IssueDescFieldKey;

                // 实时拉当前正文，剥离旧 AI 块后追加，避免重复叠加
                var current   = await FeishuMcpClient.GetCurrentFieldValueAsync(item.Id, fieldKey);
                // 飞书可能将换行规范化，用宽松正则找分割线
                var sepMatch  = System.Text.RegularExpressions.Regex.Match(
                    current, @"\n+---+\n+");
                var humanPart = sepMatch.Success
                    ? current[..sepMatch.Index].Trim()
                    : current.Trim();
                // 若人工原文本身就是 AI 内容（旧单子无人工原文），则置空
                var aiMarkers = new[] { CommentBuilder.StoryMarker, CommentBuilder.IssueMarker };
                if (aiMarkers.Any(m => humanPart.StartsWith(m))) humanPart = "";

                var fullDesc = string.IsNullOrEmpty(humanPart)
                    ? item.Comment
                    : $"{humanPart}\n\n---\n\n{item.Comment}";

                await FeishuMcpClient.UpdateTextField(item.Id, typeKey, fieldKey, fullDesc);
                SaveWrittenItem(item.Id);
                Console.WriteLine($"  [OK] {Truncate(item.Name, 50)}");
                ok++;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  [FAIL] {Truncate(item.Name, 50)} -- {ex.Message}");
                fail++;
            }
        }
        Console.WriteLine($"\n[INFO] 写入完成：成功 {ok}，失败 {fail}");
    }

    // ── 文档输出 ─────────────────────────────────────────────────────────────

    private static string WriteBugDoc(List<PendingItem> pending)
    {
        var today    = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath  = Path.Combine(ConfigDir, $"bug_analysis_{today}.md");
        var sb       = new StringBuilder();
        var writable = pending.Where(p => !p.SkipComment).ToList();
        var skipped  = pending.Where(p => p.SkipComment).ToList();

        sb.AppendLine("# 飞书缺陷分析报告");
        sb.AppendLine($"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"共 {pending.Count} 条 | 可写入飞书：{writable.Count} 条 | 关键词未命中（仅报告）：{skipped.Count} 条\n\n---\n");

        if (writable.Count > 0)
        {
            sb.AppendLine("## 可写入飞书的缺陷分析\n");
            foreach (var item in writable)
            {
                sb.AppendLine($"### [{item.Id}] {item.Name}\n");
                sb.AppendLine("分析内容：\n\n```");
                sb.AppendLine(item.Comment);
                sb.AppendLine("```\n\n---\n");
            }
        }

        if (skipped.Count > 0)
        {
            sb.AppendLine("## 关键词未命中（仅记录，不写飞书）\n");
            foreach (var item in skipped)
                sb.AppendLine($"- [{item.Id}] {item.Name}（{item.SkipReason}）");
        }

        File.WriteAllText(docPath, sb.ToString(), Encoding.UTF8);
        return docPath;
    }

    private static void WriteReviewDoc(List<PendingItem> pending, string docPath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# 飞书评论待写入审阅\n");
        sb.AppendLine($"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"共 {pending.Count} 条待写入\n\n---\n");
        foreach (var item in pending)
        {
            sb.AppendLine($"## [{item.Id}] {item.Name}\n");
            sb.AppendLine($"涉及表格：{string.Join(", ", item.Tables)}\n");
            sb.AppendLine("分析内容：\n\n```");
            sb.AppendLine(item.Comment);
            sb.AppendLine("```\n\n---\n");
        }
        File.WriteAllText(docPath, sb.ToString(), Encoding.UTF8);
    }

    // ── 工具 ─────────────────────────────────────────────────────────────────

    private static ActivityTableRules LoadRules()
    {
        if (!File.Exists(RulesPath))
        {
            Console.WriteLine($"[WARN] 规则文件不存在：{RulesPath}");
            return new ActivityTableRules();
        }
        return JsonConvert.DeserializeObject<ActivityTableRules>(File.ReadAllText(RulesPath))
               ?? new ActivityTableRules();
    }

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "…";

    private static string Now() => DateTime.Now.ToString("HH:mm:ss");
}
