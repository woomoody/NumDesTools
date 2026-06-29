using System.Text;
using Newtonsoft.Json;

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
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesTools",
        "Config"
    );
    private const string ActivityXlsx = @"C:\M1Work\public\Excels\Tables\ActivityClientData.xlsx";
    private static readonly string RulesPath = Path.Combine(ConfigDir, "ActivityTableRules.json");

    private const string McpToken = "m-002580b6-b3db-405a-aab8-007f927ef4eb";
    private const string ProjectKey = "t89j73";
    private static readonly string WrittenItemsPath = Path.Combine(ConfigDir, "written_items.json");
    private static readonly string NoPermItemsPath = Path.Combine(
        ConfigDir,
        "no_permission_items.json"
    );

    // ── 已写入记录 ───────────────────────────────────────────────────────────
    private static HashSet<string> LoadWrittenItems()
    {
        if (!File.Exists(WrittenItemsPath))
            return [];
        try
        {
            var dict =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(
                    File.ReadAllText(WrittenItemsPath)
                ) ?? [];
            return dict.Keys.ToHashSet();
        }
        catch
        {
            return [];
        }
    }

    private static void SaveWrittenItem(string itemId)
    {
        Dictionary<string, string> dict = [];
        if (File.Exists(WrittenItemsPath))
            try
            {
                dict =
                    JsonConvert.DeserializeObject<Dictionary<string, string>>(
                        File.ReadAllText(WrittenItemsPath)
                    ) ?? [];
            }
            catch { }
        dict[itemId] = DateTime.Today.ToString("yyyy-MM-dd");
        File.WriteAllText(WrittenItemsPath, JsonConvert.SerializeObject(dict, Formatting.Indented));
    }

    // ── 无权限记录（workItemId → 首次发现日期）───────────────────────────────
    private static HashSet<string> LoadNoPermItems()
    {
        if (!File.Exists(NoPermItemsPath))
            return [];
        try
        {
            var dict =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(
                    File.ReadAllText(NoPermItemsPath)
                ) ?? [];
            return dict.Keys.ToHashSet();
        }
        catch
        {
            return [];
        }
    }

    private static void SaveNoPermItem(string itemId)
    {
        Dictionary<string, string> dict = [];
        if (File.Exists(NoPermItemsPath))
            try
            {
                dict =
                    JsonConvert.DeserializeObject<Dictionary<string, string>>(
                        File.ReadAllText(NoPermItemsPath)
                    ) ?? [];
            }
            catch { }
        if (!dict.ContainsKey(itemId))
        {
            dict[itemId] = DateTime.Today.ToString("yyyy-MM-dd");
            File.WriteAllText(
                NoPermItemsPath,
                JsonConvert.SerializeObject(dict, Formatting.Indented)
            );
        }
    }

    private static bool IsPermissionError(Exception ex) =>
        ex.Message.Contains("1000052092") || ex.Message.Contains("无权编辑");

    private static readonly int ConfirmTimeoutSec = 30 * 60;

    static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;

        // ── xlsx 冲突解决 WPF GUI（独立，可被 lazygit/SmartGit/Fork 调用）──────
        // 用法：--conflict-gui <ours.xlsx> <theirs.xlsx>
        if (args.Contains("--conflict-gui"))
            return ConflictGui.Run(args);

        // ── 全功能冲突管理器（自动发现 git 仓库中所有 xlsx 冲突）─────────────
        // 用法：--conflict-manager（在 git 仓库根目录下运行，或 lazygit global context）
        if (args.Contains("--conflict-manager"))
            return ConflictManager.Run(args);

        // ── xlsx 冲突解决 TUI（独立，不依赖飞书）────────────────────────────────
        // 用法：--conflict <ours.xlsx> <theirs.xlsx> [base.xlsx]
        if (args.Contains("--conflict"))
            return ConflictTui.Run(args);

        // ── Excel 索引搜索 TUI（独立）───────────────────────────────────────────
        // 用法：--search [--index <path.json.gz>]
        if (args.Contains("--search"))
            return SearchTui.Run(args);

        // ── Excel 索引构建（独立，不依赖插件进程）──────────────────────────────
        // 用法：--build-index <excels-root-dir>
        if (args.Contains("--build-index"))
            return IndexBuilder.Run(args);

        // ── 配置自检模式（独立，不依赖飞书）────────────────────────────────────
        if (args.Contains("--validate"))
            return ConfigValidator.Run(args);

        // ── LTE地图写入模式（独立，不依赖飞书）─────────────────────────────────
        // 用法：--write-map <xlsx> [<图号1> <图号2> ...]
        //   不指定图号 → 写全部1~10
        //   指定图号   → 只写指定图（如：--write-map lte.xlsx 2 5 10）
        if (args.Contains("--write-map"))
        {
            int idx = Array.IndexOf(args, "--write-map");
            if (idx < 0 || idx + 1 >= args.Length)
            {
                Console.WriteLine("[ERROR] 用法：--write-map <lte_template.xlsx> [图号...]");
                return 1;
            }
            var xlsxPath = args[idx + 1];
            var mapNums = args.Skip(idx + 2)
                .Select(s => int.TryParse(s, out var n) ? n : -1)
                .Where(n => n >= 1 && n <= 10)
                .ToList();
            if (mapNums.Count == 0)
                LteMapWriter.RunAll(xlsxPath);
            else
                foreach (var n in mapNums)
                    LteMapWriter.Run(xlsxPath, n);
            return 0;
        }

        // ── 果酱节地图可视化（独立）─────────────────────────────────────────────
        // 用法：--write-jam-map <output.xlsx> [数据json路径]
        if (args.Contains("--write-jam-map"))
        {
            int idx = Array.IndexOf(args, "--write-jam-map");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            var dataPath =
                idx + 2 < args.Length && !args[idx + 2].StartsWith("--") ? args[idx + 2] : null;
            if (dataPath != null)
                JamMapWriter.Run(outDir, dataPath);
            else
                JamMapWriter.Run(outDir);
            return 0;
        }

        // ── Gossip Harbor 竞品分析（独立）─────────────────────────────────────
        // 用法：--write-gossip [输出目录]
        if (args.Contains("--write-gossip"))
        {
            int idx = Array.IndexOf(args, "--write-gossip");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            GossipHarborWriter.Run(outDir);
            return 0;
        }

        // ── Travel Town 竞品分析（独立）────────────────────────────────────────
        // 用法：--write-traveltown [输出目录]
        if (args.Contains("--write-traveltown"))
        {
            int idx = Array.IndexOf(args, "--write-traveltown");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            TravelTownWriter.Run(outDir);
            return 0;
        }

        // ── Merge Cooking 竞品分析（独立）──────────────────────────────────────
        // 用法：--write-mergecooking [输出目录]
        if (args.Contains("--write-mergecooking"))
        {
            int idx = Array.IndexOf(args, "--write-mergecooking");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            MergeCookingWriter.Run(outDir);
            return 0;
        }

        // ── Tasty Travels 竞品分析（独立）──────────────────────────────────────
        // 用法：--write-tasty [输出目录]
        if (args.Contains("--write-tasty"))
        {
            int idx = Array.IndexOf(args, "--write-tasty");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            TastyTravelsWriter.Run(outDir);
            return 0;
        }

        // ── 云雾地图可视化（海岛/邮差/主线，独立）──────────────────────────────
        // 用法：--write-cloud-maps [输出目录] [数据目录]   → 所有已知活动
        //       --write-ocean-map  [输出目录] [数据目录]   → 仅海岛
        if (args.Contains("--write-cloud-maps"))
        {
            int idx = Array.IndexOf(args, "--write-cloud-maps");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            var dataDir =
                idx + 2 < args.Length && !args[idx + 2].StartsWith("--") ? args[idx + 2] : null;
            CloudMapWriter.RunAll(outDir, dataDir);
            return 0;
        }
        if (args.Contains("--write-ocean-map"))
        {
            int idx = Array.IndexOf(args, "--write-ocean-map");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            var dataDir =
                idx + 2 < args.Length && !args[idx + 2].StartsWith("--") ? args[idx + 2] : null;
            CloudMapWriter.Run(CloudMapWriter.KnownActivities[0], outDir, dataDir);
            return 0;
        }

        // ── 活动写入模式（独立，不依赖飞书）────────────────────────────────────
        if (args.Contains("--write-activity"))
        {
            int idx = Array.IndexOf(args, "--write-activity");
            if (idx < 0 || idx + 1 >= args.Length)
            {
                Console.WriteLine("[ERROR] 用法：--write-activity <write_plan.json>");
                return 1;
            }
            ActivityWriter.RunFromFile(args[idx + 1]);
            return 0;
        }

        // 解析参数
        bool releaseMode = args.Contains("--release");
        bool confirmMode = args.Contains("--confirm");
        bool fetchAll = args.Contains("--all");
        bool bugsMode = args.Contains("--bugs");
        bool reviewOnly = args.Contains("--review");
        bool forceMode = args.Contains("--force"); // 跳过已有AI评论检测，强制重新分析
        List<string>? itemIds = null;
        int itemIdx = Array.IndexOf(args, "--item");
        if (itemIdx >= 0 && itemIdx + 1 < args.Length)
            itemIds = args[itemIdx + 1]
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();

        // 初始化飞书客户端
        FeishuMcpClient.McpToken = McpToken;
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
        var noPermItems = LoadNoPermItems();

        // ── 缺陷分析模式 ─────────────────────────────────────────────────────
        if (bugsMode)
        {
            Console.WriteLine(
                $"\n[{Now()}] {(itemIds != null ? $"缺陷分析 指定ID {string.Join(",", itemIds)}" : "缺陷全量分析")}..."
            );
            var kb = reviewer.LoadKnowledge();
            var pending = await ScanIssuesAsync(
                rules,
                typeIndex,
                reviewer,
                kb,
                itemIds,
                fetchAll: true,
                force: forceMode,
                writtenItems: writtenItems,
                noPermItems: noPermItems
            );
            if (pending.Count == 0)
            {
                Console.WriteLine("[INFO] 无待分析缺陷。");
                return 0;
            }

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
            var kb = reviewer.LoadKnowledge();
            var seen = new HashSet<string>();
            var pending = await ScanStoriesAsync(
                rules,
                typeIndex,
                reviewer,
                kb,
                seen,
                itemIds,
                fetchAll,
                forceMode,
                writtenItems,
                noPermItems
            );
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
                Console.WriteLine($"  未识别到表格：{Truncate(s.Name, 40)}");
                continue;
            }

            if (!force)
            {
                if (writtenItems?.Contains(s.Id) == true)
                {
                    Console.WriteLine($"  已有AI分析，跳过：{Truncate(s.Name, 40)}");
                    continue;
                }
                if (
                    await FeishuWorkItemFetcher.HasExistingAiCommentAsync(
                        s,
                        CommentBuilder.StoryMarker
                    )
                )
                {
                    Console.WriteLine($"  已有AI分析，跳过：{Truncate(s.Name, 40)}");
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

            // 从人工评论中提取 git commit SHA 并分析
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
                    ? $"  新增待分析[{phaseTag}]（无写入权限）：{Truncate(s.Name, 40)}"
                    : $"  新增待写入[{phaseTag}]：{Truncate(s.Name, 40)}"
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

    // ── 缺陷扫描 ─────────────────────────────────────────────────────────────

    private static async Task<List<PendingItem>> ScanIssuesAsync(
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
                Console.WriteLine($"  未提取到活动ID，跳过：{Truncate(issue.Name, 40)}");
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
                        new TableMatch(mainDef.ExcelFile, mainDef.Desc, [], mainDef.KeyField, null)
                    ];
            }
            if (tables.Count == 0)
            {
                Console.WriteLine($"  未识别到相关表格，跳过：{Truncate(issue.Name, 40)}");
                continue;
            }

            if (!force)
            {
                if (writtenItems?.Contains(issue.Id) == true)
                {
                    Console.WriteLine($"  已有AI缺陷分析，跳过：{Truncate(issue.Name, 40)}");
                    continue;
                }
                if (
                    await FeishuWorkItemFetcher.HasExistingAiCommentAsync(
                        issue,
                        CommentBuilder.IssueMarker
                    )
                )
                {
                    Console.WriteLine($"  已有AI缺陷分析，跳过：{Truncate(issue.Name, 40)}");
                    continue;
                }
            }

            var typeNums = typeNum.HasValue ? [typeNum.Value.ToString()] : new List<string>();
            var historyNotes = reviewer.QueryKnowledge(typeNums, kb);
            var findChain = findAnalyzer.Analyze(issue.Name, issue.Desc);
            var (_, _, hintMatched) = CommentBuilder.InferBugHints(issue.Name);

            // 从人工评论中提取 git commit SHA 并分析
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
            // SkipComment = 无权限 OR 关键词未命中（两者都不写飞书，但都参与分析）
            var skipWrite = noPerm || !hintMatched;
            var skipReason = noPerm ? "无编辑权限" : (hintMatched ? "" : "标题关键词未匹配到已知配置规则");

            if (noPerm)
                Console.WriteLine($"  新增缺陷（无写入权限）：{Truncate(issue.Name, 40)}");
            else if (hintMatched)
                Console.WriteLine($"  新增缺陷待写入：{Truncate(issue.Name, 40)}");
            else
                Console.WriteLine($"  新增缺陷（仅报告，关键词未命中）：{Truncate(issue.Name, 40)}");
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

    // ── 写入 ─────────────────────────────────────────────────────────────────

    private static async Task PromptAndWrite(
        List<PendingItem> pending,
        bool releaseMode,
        bool confirmMode
    )
    {
        Console.WriteLine();
        Console.WriteLine(new string('=', 60));
        Console.WriteLine(releaseMode ? "【发布模式】待写入评论摘要（输入 y 执行写入）：" : "【测试模式】分析结果预览（不会写入飞书）：");
        Console.WriteLine(new string('=', 60));
        for (int i = 0; i < pending.Count; i++)
        {
            var item = pending[i];
            Console.WriteLine($"  {i + 1}. [{item.Id}] {Truncate(item.Name, 50)}");
            Console.WriteLine($"     涉及表格：{string.Join(", ", item.Tables)}");
        }
        Console.WriteLine(new string('=', 60));

        var today = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath = Path.Combine(ConfigDir, $"pending_comments_{today}.md");
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

    private static async Task WriteToFeishu(
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
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(ConfirmTimeoutSec));
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

        var noPermItems = LoadNoPermItems();

        int skipped = pending.Count(p => p.SkipComment);
        int noPerm = pending.Count(p => !p.SkipComment && noPermItems.Contains(p.Id));
        if (skipped > 0)
            Console.WriteLine($"[INFO] 跳过 {skipped} 条关键词未命中的缺陷（仅报告不写飞书）");
        if (noPerm > 0)
            Console.WriteLine($"[INFO] 跳过 {noPerm} 条无编辑权限的工作项（已记录在 no_permission_items.json）");

        Console.WriteLine("\n[INFO] 开始写入飞书正文...");
        int ok = 0,
            fail = 0;

        foreach (var item in pending)
        {
            if (item.SkipComment)
                continue;
            if (noPermItems.Contains(item.Id))
            {
                Console.WriteLine($"  [SKIP/无权限] {Truncate(item.Name, 50)}");
                continue;
            }

            var typeKey = item.ItemType == "story" ? "story" : "issue";
            try
            {
                var fieldKey =
                    typeKey == "story"
                        ? FeishuMcpClient.StoryContentFieldKey
                        : FeishuMcpClient.IssueDescFieldKey;

                // 实时拉当前正文，剥离旧 AI 块后追加，避免重复叠加
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
                SaveWrittenItem(item.Id);
                Console.WriteLine($"  [OK] {Truncate(item.Name, 50)}");
                ok++;
            }
            catch (Exception ex)
            {
                if (IsPermissionError(ex))
                {
                    SaveNoPermItem(item.Id);
                    noPermItems.Add(item.Id);
                    Console.WriteLine($"  [无权限] {Truncate(item.Name, 50)}（已记录，后续自动跳过写入）");
                }
                else
                {
                    Console.WriteLine($"  [FAIL] {Truncate(item.Name, 50)} -- {ex.Message}");
                }
                fail++;
            }
        }
        Console.WriteLine($"\n[INFO] 写入完成：成功 {ok}，失败 {fail}");
    }

    // ── 文档输出 ─────────────────────────────────────────────────────────────

    private static string WriteBugDoc(List<PendingItem> pending)
    {
        var today = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath = Path.Combine(ConfigDir, $"bug_analysis_{today}.md");
        var sb = new StringBuilder();
        var writable = pending.Where(p => !p.SkipComment).ToList();
        var skipped = pending.Where(p => p.SkipComment).ToList();

        sb.AppendLine("# 飞书缺陷分析报告");
        sb.AppendLine($"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine(
            $"共 {pending.Count} 条 | 可写入飞书：{writable.Count} 条 | 关键词未命中（仅报告）：{skipped.Count} 条"
        );
        sb.Append(BuildBugStatsSummary(pending));
        sb.AppendLine("\n---\n");

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
        sb.AppendLine($"共 {pending.Count} 条待写入\n");
        sb.Append(BuildStoryStatsSummary(pending));
        sb.AppendLine("\n---\n");
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

    private static string BuildStoryStatsSummary(List<PendingItem> pending)
    {
        if (pending.Count == 0)
            return "";

        var sb = new StringBuilder();
        sb.AppendLine("\n## 宏观统计\n");

        // 活动类型分类（按 Name 关键词）
        var typeGroups = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["LTE 限时地图"] = [],
            ["BattlePass / 季节BP"] = [],
            ["合并玩法 (二合/砸冰)"] = [],
            ["礼包 / 商业化"] = [],
            ["BumperHarvest / 联盟"] = [],
            ["Bingo / 小游戏"] = [],
            ["其他需求"] = [],
        };

        foreach (var item in pending)
        {
            var n = item.Name;
            if (n.Contains("LTE") || n.Contains("限时地图") || n.Contains("寻宝玩法"))
                typeGroups["LTE 限时地图"].Add(item.Name);
            else if (
                n.Contains("BP")
                || n.Contains("BattlePass")
                || n.Contains("通行证")
                || n.Contains("季节")
            )
                typeGroups["BattlePass / 季节BP"].Add(item.Name);
            else if (n.Contains("二合") || n.Contains("砸冰") || n.Contains("V4") || n.Contains("V6"))
                typeGroups["合并玩法 (二合/砸冰)"].Add(item.Name);
            else if (
                n.Contains("礼包")
                || n.Contains("商业化")
                || n.Contains("付费")
                || n.Contains("超级卡")
                || n.Contains("奖励翻转")
                || n.Contains("邮票")
                || n.Contains("线性")
            )
                typeGroups["礼包 / 商业化"].Add(item.Name);
            else if (n.Contains("联盟") || n.Contains("BumperHarvest") || n.Contains("Bumper"))
                typeGroups["BumperHarvest / 联盟"].Add(item.Name);
            else if (n.Contains("Bingo") || n.Contains("气球") || n.Contains("宝箱"))
                typeGroups["Bingo / 小游戏"].Add(item.Name);
            else
                typeGroups["其他需求"].Add(item.Name);
        }

        sb.AppendLine("### 活动类型分布\n");
        sb.AppendLine("| 类型 | 数量 | 占比 |");
        sb.AppendLine("|------|------|------|");
        foreach (
            var (type, items) in typeGroups
                .Where(kv => kv.Value.Count > 0)
                .OrderByDescending(kv => kv.Value.Count)
        )
        {
            var pct = (double)items.Count / pending.Count * 100;
            sb.AppendLine($"| {type} | {items.Count} | {pct:F0}% |");
        }

        // 续期 vs 首期
        var renewal = pending.Count(p =>
            p.Comment.Contains("续期")
            || System.Text.RegularExpressions.Regex.IsMatch(p.Name, @"第\d+期")
        );
        var firstTime = pending.Count - renewal;
        sb.AppendLine($"\n### 续期 vs 首期\n");
        sb.AppendLine($"- 首期/全新：**{firstTime}** 条");
        sb.AppendLine($"- 续期迭代：**{renewal}** 条");

        // 高频涉及表格 Top 5
        var tableFreq = pending
            .SelectMany(p => p.Tables)
            .GroupBy(t => t)
            .OrderByDescending(g => g.Count())
            .Take(5)
            .ToList();

        if (tableFreq.Count > 0)
        {
            sb.AppendLine($"\n### 高频配置表 Top 5\n");
            sb.AppendLine("| 表名 | 出现次数 |");
            sb.AppendLine("|------|---------|");
            foreach (var g in tableFreq)
                sb.AppendLine($"| {g.Key} | {g.Count()} |");
        }

        // 平均涉及表数
        var avgTables = pending.Average(p => p.Tables.Count);
        var maxItem = pending.MaxBy(p => p.Tables.Count);
        sb.AppendLine($"\n### 配置复杂度\n");
        sb.AppendLine($"- 平均涉及表数：**{avgTables:F1}** 张");
        if (maxItem is not null)
            sb.AppendLine(
                $"- 最复杂需求：**{Truncate(maxItem.Name.ReplaceLineEndings(" "), 40)}**（{maxItem.Tables.Count} 张表）"
            );

        return sb.ToString();
    }

    private static string BuildBugStatsSummary(List<PendingItem> pending)
    {
        if (pending.Count == 0)
            return "";

        var sb = new StringBuilder();
        sb.AppendLine("\n## 宏观统计\n");

        // ── 第一层：所属活动模块 ──────────────────────────────────────────────
        var moduleGroups = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in pending)
        {
            var n = item.Name;
            string module;
            if (
                n.Contains("LTE")
                || n.Contains("限时地图")
                || n.Contains("寻宝")
                || n.Contains("解谜玩法")
                || n.Contains("闯关")
            )
                module = "LTE 限时地图";
            else if (
                n.Contains("BP")
                || n.Contains("BattlePass")
                || n.Contains("通行证")
                || n.Contains("季节")
            )
                module = "BattlePass / 季节BP";
            else if (n.Contains("二合") || n.Contains("砸冰") || n.Contains("V4") || n.Contains("V6"))
                module = "合并玩法";
            else if (n.Contains("礼包") || n.Contains("商业化") || n.Contains("付费") || n.Contains("超级卡"))
                module = "礼包 / 商业化";
            else if (n.Contains("关卡岛") || n.Contains("地编") || n.Contains("地图"))
                module = "关卡岛 / 地编";
            else if (n.Contains("公会") || n.Contains("联盟") || n.Contains("社交"))
                module = "公会 / 社交";
            else
                module = "其他";
            moduleGroups[module] = moduleGroups.GetValueOrDefault(module) + 1;
        }

        sb.AppendLine("### 所属活动模块\n");
        sb.AppendLine("| 模块 | 数量 | 占比 |");
        sb.AppendLine("|------|------|------|");
        foreach (var (mod, cnt) in moduleGroups.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"| {mod} | {cnt} | {(double)cnt / pending.Count * 100:F0}% |");

        // ── 第二层：症状类型（宽泛匹配，覆盖自然语言描述） ─────────────────
        var symptomDefs = new (string Label, string[] Keywords)[]
        {
            ("多语言 / 文案", ["多语言", "翻译", "文案", "文字", "本地化", "语言缺失", "语言", "key"]),
            (
                "UI / 界面 / 位置",
                [
                    "界面",
                    "位置",
                    "UI",
                    "图标",
                    "样式",
                    "切图",
                    "排版",
                    "布局",
                    "弹窗",
                    "显示",
                    "展示",
                    "帮助图",
                    "遮挡",
                    "云层"
                ]
            ),
            ("缺失 / 不显示", ["缺失", "缺少", "不显示", "未显示", "没有", "丢失", "消失"]),
            ("合成 / 链条", ["合成", "主链", "链条", "三合", "二合", "合并", "merge", "产出不够", "材料不够"]),
            ("任务 / 触发", ["任务", "触发", "不触发", "不生效", "无法触发", "卡住", "卡死", "卡关", "寻找"]),
            ("奖励 / 数值", ["奖励", "数值", "数量", "价格", "积分", "得分"]),
            ("配置 / 数据", ["配置", "数据错误", "表格", "配表", "error", "报错"]),
            ("崩溃 / 闪退", ["崩溃", "闪退", "crash", "卡顿", "ANR"]),
            ("标识 / 图标", ["标识", "图标", "icon", "角标", "标记", "蛛网"]),
            ("NPC / 地标", ["NPC", "地标", "建造", "建筑"]),
        };

        // 每条 bug 按优先级匹配第一个命中的症状（避免重复计数）
        var symptomCounts = new Dictionary<string, int>();
        var uncategorized = 0;
        foreach (var item in pending)
        {
            var matched = false;
            foreach (var def in symptomDefs)
            {
                if (
                    def.Keywords.Any(kw =>
                        item.Name.Contains(kw, StringComparison.OrdinalIgnoreCase)
                    )
                )
                {
                    symptomCounts[def.Label] = symptomCounts.GetValueOrDefault(def.Label) + 1;
                    matched = true;
                    break;
                }
            }
            if (!matched)
                uncategorized++;
        }

        sb.AppendLine("\n### 缺陷症状分布\n");
        sb.AppendLine("| 症状 | 数量 | 占比 |");
        sb.AppendLine("|------|------|------|");
        foreach (var (label, count) in symptomCounts.OrderByDescending(kv => kv.Value))
            sb.AppendLine($"| {label} | {count} | {(double)count / pending.Count * 100:F0}% |");
        if (uncategorized > 0)
            sb.AppendLine(
                $"| 症状未识别 | {uncategorized} | {(double)uncategorized / pending.Count * 100:F0}% |"
            );

        // ── 可写入 vs 跳过 ──────────────────────────────────────────────────
        var writable = pending.Count(p => !p.SkipComment);
        var skipped = pending.Count(p => p.SkipComment);
        sb.AppendLine($"\n### 可分析率\n");
        sb.AppendLine(
            $"- 关键词命中（可写飞书）：**{writable}** 条 ({(double)writable / pending.Count * 100:F0}%)"
        );
        sb.AppendLine(
            $"- 关键词未命中（仅记录）：**{skipped}** 条 ({(double)skipped / pending.Count * 100:F0}%)"
        );

        // ── LTE 专项：涉及最多 bug 的活动 ID Top 5 ─────────────────────────
        var actIdPattern = new System.Text.RegularExpressions.Regex(@"\b7\d{5}\b");
        var actFreq = pending
            .SelectMany(p => actIdPattern.Matches(p.Name).Select(m => m.Value))
            .GroupBy(id => id)
            .OrderByDescending(g => g.Count())
            .Take(5)
            .ToList();

        if (actFreq.Count > 0)
        {
            sb.AppendLine("\n### Bug 最集中的活动 ID Top 5\n");
            sb.AppendLine("| 活动ID | Bug数量 |");
            sb.AppendLine("|--------|---------|");
            foreach (var g in actFreq)
                sb.AppendLine($"| {g.Key} | {g.Count()} |");
        }

        return sb.ToString();
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

    private static string Truncate(string s, int max) => s.Length <= max ? s : s[..max] + "…";

    private static string Now() => DateTime.Now.ToString("HH:mm:ss");
}
