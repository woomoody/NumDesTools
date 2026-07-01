using System.Text;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace NumDesTools.Scanner;

internal static class ScanOrchestrator
{
    internal static async Task<List<PendingItem>> ScanStoriesAsync(
        ActivityTableRules rules,
        ActivityTypeIndex typeIndex,
        KnowledgeReviewer reviewer,
        KnowledgeBase kb,
        HashSet<string> seenIds,
        List<string>? itemIds,
        bool fetchAll,
        bool force = false,
        HashSet<string>? writtenItems = null,
        HashSet<string>? noPermItems = null
    )
    {
        var stories = await FeishuWorkItemFetcher.FetchStoriesAsync(itemIds, fetchAll);
        Console.WriteLine($"  开始分析 {stories.Count} 条需求...\n");
        var pending = new List<PendingItem>();

        foreach (var s in stories)
        {
            if (itemIds == null && !fetchAll && !seenIds.Add(s.Id))
                continue;
            seenIds.Add(s.Id);

            if (!TableMatcher.IsPlannerRelated(s))
                continue;

            var (tables, phase) = TableMatcher.IdentifyTables(s, rules, typeIndex.TypeIndex);
            if (tables.Count == 0)
            {
                Console.WriteLine($"  未识别到表格：{Helpers.Truncate(s.Name, 40)}");
                continue;
            }

            if (!force)
            {
                if (writtenItems?.Contains(s.Id) == true)
                {
                    Console.WriteLine($"  已有AI分析，跳过：{Helpers.Truncate(s.Name, 40)}");
                    continue;
                }
                if (
                    await FeishuWorkItemFetcher.HasExistingAiCommentAsync(
                        s,
                        CommentBuilder.StoryMarker
                    )
                )
                {
                    Console.WriteLine($"  已有AI分析，跳过：{Helpers.Truncate(s.Name, 40)}");
                    continue;
                }
            }

            var typeNums = tables
                .Where(t => t.TypeNum != null)
                .Select(t => t.TypeNum!)
                .Distinct()
                .ToList();
            var historyNotes = reviewer.QueryKnowledge(typeNums, kb);
            var phaseTag = phase > 1 ? $"第{phase}期续期" : "首期/全新";

            List<string> humanComments = [];
            try
            {
                humanComments = (await FeishuWorkItemFetcher.GetCommentsAsync(s.Id))
                    .Select(c => c.Content)
                    .ToList();
            }
            catch { }
            var gitAnalysis = GitCommitAnalyzer.Analyze(s.Name, s.Desc, humanComments);

            var noPerm = noPermItems?.Contains(s.Id) == true;
            Console.WriteLine(
                noPerm
                    ? $"  新增待分析[{phaseTag}]（无写入权限）：{Helpers.Truncate(s.Name, 40)}"
                    : $"  新增待写入[{phaseTag}]：{Helpers.Truncate(s.Name, 40)}"
            );
            Console.WriteLine($"    涉及表格：{string.Join(", ", tables.Select(t => t.Excel))}");
            if (historyNotes.Count > 0)
                Console.WriteLine($"    引用历史经验：{historyNotes.Count} 条");
            if (gitAnalysis != null)
                Console.WriteLine($"    关联 Git 提交：已提取");

            pending.Add(
                new PendingItem
                {
                    Id = s.Id,
                    Name = s.Name,
                    Tables = tables.Select(t => t.Excel).ToList(),
                    Comment = CommentBuilder.BuildStoryComment(
                        s,
                        tables,
                        phase,
                        historyNotes,
                        gitAnalysis
                    ),
                    ItemType = "story",
                    OriginalDesc = s.Desc,
                    SkipComment = noPerm,
                    SkipReason = noPerm ? "无编辑权限" : "",
                }
            );
        }
        return pending;
    }

    internal static async Task<List<PendingItem>> ScanIssuesAsync(
        ActivityTableRules rules,
        ActivityTypeIndex typeIndex,
        KnowledgeReviewer reviewer,
        KnowledgeBase kb,
        List<string>? itemIds,
        bool fetchAll,
        bool force = false,
        HashSet<string>? writtenItems = null,
        HashSet<string>? noPermItems = null
    )
    {
        var issues = await FeishuWorkItemFetcher.FetchIssuesAsync(itemIds, fetchAll);
        var findAnalyzer = new FindChainAnalyzer();
        Console.WriteLine($"  开始分析 {issues.Count} 条缺陷...\n");
        var pending = new List<PendingItem>();

        foreach (var issue in issues)
        {
            var actIds = FeishuWorkItemFetcher.ExtractActivityIds(issue.Name);
            if (actIds.Count == 0)
                actIds = FeishuWorkItemFetcher.ExtractActivityIds(issue.Desc);
            if (actIds.Count == 0)
            {
                Console.WriteLine($"  未提取到活动ID，跳过：{Helpers.Truncate(issue.Name, 40)}");
                continue;
            }

            var actId = actIds[0];
            var (typeNum, typeNote) = typeIndex.LookupByActivityId(actId);

            var fakeItem = new WorkItem(
                issue.Id,
                issue.Name,
                (typeNum.HasValue ? $"type={typeNum} " : "") + issue.Desc,
                issue.Status
            );
            var (tables, _) = TableMatcher.IdentifyTables(fakeItem, rules, typeIndex.TypeIndex);

            if (tables.Count == 0 && typeNum.HasValue)
            {
                var mainDef = rules.Tables.FirstOrDefault(t => t.Name == "ActivityClientData");
                if (mainDef != null)
                    tables =
                    [
                        new TableMatch(mainDef.ExcelFile, mainDef.Desc, [], mainDef.KeyField, null),
                    ];
            }
            if (tables.Count == 0)
            {
                Console.WriteLine($"  未识别到相关表格，跳过：{Helpers.Truncate(issue.Name, 40)}");
                continue;
            }

            if (!force)
            {
                if (writtenItems?.Contains(issue.Id) == true)
                {
                    Console.WriteLine($"  已有AI缺陷分析，跳过：{Helpers.Truncate(issue.Name, 40)}");
                    continue;
                }
                if (
                    await FeishuWorkItemFetcher.HasExistingAiCommentAsync(
                        issue,
                        CommentBuilder.IssueMarker
                    )
                )
                {
                    Console.WriteLine($"  已有AI缺陷分析，跳过：{Helpers.Truncate(issue.Name, 40)}");
                    continue;
                }
            }

            var typeNums = typeNum.HasValue ? [typeNum.Value.ToString()] : new List<string>();
            var historyNotes = reviewer.QueryKnowledge(typeNums, kb);
            var findChain = findAnalyzer.Analyze(issue.Name, issue.Desc);
            var (_, _, hintMatched) = CommentBuilder.InferBugHints(issue.Name);

            List<string> humanComments = [];
            try
            {
                humanComments = (await FeishuWorkItemFetcher.GetCommentsAsync(issue.Id))
                    .Select(c => c.Content)
                    .ToList();
            }
            catch { }
            var gitAnalysis = GitCommitAnalyzer.Analyze(issue.Name, issue.Desc, humanComments);

            var noPerm = noPermItems?.Contains(issue.Id) == true;
            var skipWrite = noPerm || !hintMatched;
            var skipReason = noPerm
                ? "无编辑权限"
                : (hintMatched ? "" : "标题关键词未匹配到已知配置规则");

            if (noPerm)
                Console.WriteLine($"  新增缺陷（无写入权限）：{Helpers.Truncate(issue.Name, 40)}");
            else if (hintMatched)
                Console.WriteLine($"  新增缺陷待写入：{Helpers.Truncate(issue.Name, 40)}");
            else
                Console.WriteLine(
                    $"  新增缺陷（仅报告，关键词未命中）：{Helpers.Truncate(issue.Name, 40)}"
                );
            Console.WriteLine($"    activityID={actId}, type={typeNum?.ToString() ?? "?"}");
            if (historyNotes.Count > 0)
                Console.WriteLine($"    引用历史经验：{historyNotes.Count} 条");
            if (findChain != null)
                Console.WriteLine($"    寻找链分析：已提取到ID追溯");
            if (gitAnalysis != null)
                Console.WriteLine($"    关联 Git 提交：已提取");

            pending.Add(
                new PendingItem
                {
                    Id = issue.Id,
                    Name = issue.Name,
                    Tables = tables.Select(t => t.Excel).ToList(),
                    Comment = CommentBuilder.BuildIssueComment(
                        issue,
                        tables,
                        actId,
                        typeNum,
                        typeNote,
                        historyNotes,
                        findChain,
                        gitAnalysis
                    ),
                    ItemType = "issue",
                    SkipComment = skipWrite,
                    SkipReason = skipReason,
                    OriginalDesc = issue.Desc,
                }
            );
        }
        return pending;
    }

    internal static async Task PromptAndWrite(
        List<PendingItem> pending,
        bool releaseMode,
        bool confirmMode
    )
    {
        Console.WriteLine();
        Console.WriteLine(new string('=', 60));
        Console.WriteLine(
            releaseMode
                ? "【发布模式】待写入评论摘要（输入 y 执行写入）："
                : "【测试模式】分析结果预览（不会写入飞书）："
        );
        Console.WriteLine(new string('=', 60));
        for (int i = 0; i < pending.Count; i++)
        {
            var item = pending[i];
            Console.WriteLine($"  {i + 1}. [{item.Id}] {Helpers.Truncate(item.Name, 50)}");
            Console.WriteLine($"     涉及表格：{string.Join(", ", item.Tables)}");
        }
        Console.WriteLine(new string('=', 60));

        var today = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath = Path.Combine(Helpers.ConfigDir, $"pending_comments_{today}.md");
        ReportGenerator.WriteReviewDoc(pending, docPath);
        Console.WriteLine($"\n[INFO] 审阅文档：{docPath}");

        if (!releaseMode)
        {
            Console.WriteLine("[INFO] 测试模式，以上为分析结果，不执行写入。");
            Console.WriteLine("[INFO] 使用 --release 参数启动可开启写入。");
            return;
        }
        await WriteToFeishu(pending, confirmMode, "[需求评论]");
    }

    internal static async Task WriteToFeishu(
        List<PendingItem> pending,
        bool confirmMode,
        string label
    )
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
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(Helpers.ConfirmTimeoutSec));
            string? input = null;
            try
            {
                var readTask = Task.Run(() => Console.ReadLine(), cts.Token);
                input = await readTask;
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("\n[INFO] 超时，本次不写入。");
                return;
            }
            answer = (input ?? "").Trim().ToLower();
        }

        if (answer != "y")
        {
            Console.WriteLine("[INFO] 已取消。");
            return;
        }

        var noPermItems = Helpers.LoadSet(Helpers.NoPermItemsPath);

        int skipped = pending.Count(p => p.SkipComment);
        int noPerm = pending.Count(p => !p.SkipComment && noPermItems.Contains(p.Id));
        if (skipped > 0)
            Console.WriteLine($"[INFO] 跳过 {skipped} 条关键词未命中的缺陷（仅报告不写飞书）");
        if (noPerm > 0)
            Console.WriteLine(
                $"[INFO] 跳过 {noPerm} 条无编辑权限的工作项（已记录在 no_permission_items.json）"
            );

        Console.WriteLine("\n[INFO] 开始写入飞书正文...");
        int ok = 0,
            fail = 0;

        foreach (var item in pending)
        {
            if (item.SkipComment)
                continue;
            if (noPermItems.Contains(item.Id))
            {
                Console.WriteLine($"  [SKIP/无权限] {Helpers.Truncate(item.Name, 50)}");
                continue;
            }

            var typeKey = item.ItemType == "story" ? "story" : "issue";
            try
            {
                var fieldKey =
                    typeKey == "story"
                        ? FeishuMcpClient.StoryContentFieldKey
                        : FeishuMcpClient.IssueDescFieldKey;

                var current = await FeishuMcpClient.GetCurrentFieldValueAsync(item.Id, fieldKey);
                var sepMatch = System.Text.RegularExpressions.Regex.Match(current, @"\n+---+\n+");
                var humanPart = sepMatch.Success
                    ? current[..sepMatch.Index].Trim()
                    : current.Trim();
                var aiMarkers = new[] { CommentBuilder.StoryMarker, CommentBuilder.IssueMarker };
                if (aiMarkers.Any(m => humanPart.StartsWith(m)))
                    humanPart = "";

                var fullDesc = string.IsNullOrEmpty(humanPart)
                    ? item.Comment
                    : $"{humanPart}\n\n---\n\n{item.Comment}";

                await FeishuMcpClient.UpdateTextField(item.Id, typeKey, fieldKey, fullDesc);
                Helpers.SaveSet(Helpers.WrittenItemsPath, item.Id);
                Console.WriteLine($"  [OK] {Helpers.Truncate(item.Name, 50)}");
                ok++;
            }
            catch (Exception ex)
            {
                if (Helpers.IsPermissionError(ex))
                {
                    Helpers.SaveSet(Helpers.NoPermItemsPath, item.Id, overwriteExisting: false);
                    noPermItems.Add(item.Id);
                    Console.WriteLine(
                        $"  [无权限] {Helpers.Truncate(item.Name, 50)}（已记录，后续自动跳过写入）"
                    );
                }
                else
                {
                    Console.WriteLine($"  [FAIL] {Helpers.Truncate(item.Name, 50)} -- {ex.Message}");
                }
                fail++;
            }
        }
        Console.WriteLine($"\n[INFO] 写入完成：成功 {ok}，失败 {fail}");
    }

    internal static void WarnIfConfigDirty(ActivityTableRules rules)
    {
        int errorCount = 0;
        foreach (var def in rules.Tables.Where(t => !string.IsNullOrEmpty(t.ExcelFile)))
        {
            var path = Path.Combine(Helpers.TablesDir, def.ExcelFile);
            var r1 = L1_TableValidator.Validate(path, def);
            var r2 = L2_RefValidator.Validate(Helpers.TablesDir, def, rules);
            errorCount += r1.Issues.Count(i => i.Level == Severity.Error);
            errorCount += r2.Issues.Count(i => i.Level == Severity.Error);
        }
        if (errorCount <= 0)
            return;
        var prev = Console.ForegroundColor;
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(
            $"[WARN] 配置表有 {errorCount} 处引用断裂，AI 评论可能基于脏数据（使用 --validate 查看详情）"
        );
        Console.ForegroundColor = prev;
    }
}
