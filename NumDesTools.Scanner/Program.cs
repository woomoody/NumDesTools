using System.Text;
using OfficeOpenXml;

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
    static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        if (args.Contains("--conflict-gui"))
            return ConflictGui.Run(args);

        if (args.Contains("--conflict-manager"))
            return ConflictManager.Run(args);

        if (args.Contains("--conflict-manager-tui"))
            return ConflictManagerTui.Run(args);

        if (args.Contains("--conflict"))
            return ConflictTui.Run(args);

        if (args.Contains("--search"))
            return SearchTui.Run(args);

        if (args.Contains("--build-index"))
            return IndexBuilder.Run(args);

        if (args.Contains("--validate"))
            return ConfigValidator.Run(args);

        if (args.Contains("--activity-batch"))
        {
            int idx = Array.IndexOf(args, "--activity-batch");
            var idsFile =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--")
                    ? args[idx + 1]
                    : @"C:\Users\cent\Documents\tmp\activity_ids_since_2026_03_01.txt";
            var mdPath =
                idx + 2 < args.Length && !args[idx + 2].StartsWith("--") ? args[idx + 2] : null;

            if (!File.Exists(idsFile))
            {
                Console.Error.WriteLine($"[ERROR] 活动 ID 清单文件不存在：{idsFile}");
                return 1;
            }

            var allIds = File.ReadAllLines(idsFile, Encoding.UTF8)
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s) && s.All(char.IsDigit))
                .ToHashSet();

            Console.WriteLine($"读取到 {allIds.Count} 个活动 ID，开始批量验证...");

            var excelPath = @"C:\M1Work\public\Excels\Tables\ActivityClientData.xlsx";
            var result = ActivityConfigValidator.RunHeadless(excelPath, allIds);

            Console.WriteLine();
            Console.WriteLine(
                result.Success
                    ? $"✅ 批量验证完成：{result.ErrorCount} 个配置问题"
                    : $"❌ 批量验证失败：{result.ErrorMessage}"
            );
            Console.WriteLine(
                $"详细报告：{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\\tmp\\activity_batch_result.txt"
            );

            if (mdPath != null && result.ReportText is { Length: > 0 })
            {
                Directory.CreateDirectory(Path.GetDirectoryName(mdPath)!);
                File.WriteAllText(mdPath, result.ReportText, Encoding.UTF8);
                Console.WriteLine($"📄 Markdown 报告：{mdPath}");
            }

            return result.Success ? 0 : 1;
        }

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

        if (args.Contains("--write-gossip"))
        {
            int idx = Array.IndexOf(args, "--write-gossip");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            GossipHarborWriter.Run(outDir);
            return 0;
        }

        if (args.Contains("--write-traveltown"))
        {
            int idx = Array.IndexOf(args, "--write-traveltown");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            TravelTownWriter.Run(outDir);
            return 0;
        }

        if (args.Contains("--write-mergecooking"))
        {
            int idx = Array.IndexOf(args, "--write-mergecooking");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            MergeCookingWriter.Run(outDir);
            return 0;
        }

        if (args.Contains("--write-tasty"))
        {
            int idx = Array.IndexOf(args, "--write-tasty");
            var outDir =
                idx + 1 < args.Length && !args[idx + 1].StartsWith("--") ? args[idx + 1] : null;
            TastyTravelsWriter.Run(outDir);
            return 0;
        }

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

        bool releaseMode = args.Contains("--release");
        bool confirmMode = args.Contains("--confirm");
        bool fetchAll = args.Contains("--all");
        bool bugsMode = args.Contains("--bugs");
        bool reviewOnly = args.Contains("--review");
        bool forceMode = args.Contains("--force");
        var xlsxIdx = Array.IndexOf(args, "--xlsx");
        var summaryXlsx = xlsxIdx >= 0 && xlsxIdx + 1 < args.Length ? args[xlsxIdx + 1] : null;
        var summaryApiKey = Environment.GetEnvironmentVariable("ANTHROPIC_AUTH_TOKEN") ?? "";
        var summaryBaseUrl = "https://litellm.solotopia.net/v1";
        List<string>? itemIds = null;
        int itemIdx = Array.IndexOf(args, "--item");
        if (itemIdx >= 0 && itemIdx + 1 < args.Length)
            itemIds = args[itemIdx + 1]
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .ToList();

        (FeishuMcpClient.McpToken, FeishuMcpClient.ProjectKey) = Helpers.LoadFeishuConfig();

        var rules = Helpers.LoadRules();
        var typeIndex = new ActivityTypeIndex();
        typeIndex.Load(Helpers.ActivityXlsx);

        var reviewer = new KnowledgeReviewer(Helpers.ConfigDir);

        if (reviewOnly)
        {
            await reviewer.RunReviewPassAsync(rules, typeIndex);
            return 0;
        }

        ScanOrchestrator.WarnIfConfigDirty(rules);

        var writtenItems = Helpers.LoadSet(Helpers.WrittenItemsPath);
        var noPermItems = Helpers.LoadSet(Helpers.NoPermItemsPath);

        if (bugsMode)
        {
            Console.WriteLine(
                $"\n[{Helpers.Now()}] {(itemIds != null ? $"缺陷分析 指定ID {string.Join(",", itemIds)}" : "缺陷全量分析")}..."
            );
            var kb = reviewer.LoadKnowledge();
            var pending = await ScanOrchestrator.ScanIssuesAsync(
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

            var docPath = ReportGenerator.WriteBugDoc(pending);
            if (summaryXlsx is not null)
                await ReportGenerator.GenerateSummaryAsync(
                    summaryXlsx,
                    summaryApiKey,
                    summaryBaseUrl
                );
            Console.WriteLine($"\n[INFO] 缺陷分析文档：{docPath}");

            if (!releaseMode)
            {
                Console.WriteLine("[INFO] 测试模式，以上为分析结果，不执行写入。");
                Console.WriteLine("[INFO] 使用 --release 参数启动可开启写入飞书缺陷评论。");
                return 0;
            }
            await ScanOrchestrator.WriteToFeishu(pending, confirmMode, "[缺陷评论]");
            return 0;
        }

        if (itemIds != null || fetchAll)
        {
            var label = itemIds != null ? $"指定ID {string.Join(",", itemIds)}" : "全量扫描";
            Console.WriteLine($"\n[{Helpers.Now()}] {label}...");
            var kb = reviewer.LoadKnowledge();
            var seen = new HashSet<string>();
            var pending = await ScanOrchestrator.ScanStoriesAsync(
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
            {
                await ScanOrchestrator.PromptAndWrite(pending, releaseMode, confirmMode);
                if (summaryXlsx is not null)
                    await ReportGenerator.GenerateSummaryAsync(
                        summaryXlsx,
                        summaryApiKey,
                        summaryBaseUrl
                    );
            }
            else
                Console.WriteLine("[INFO] 无待写入项。");
            return 0;
        }

        Console.WriteLine("[INFO] 请指定模式：--all | --bugs | --review | --item <id>");
        return 1;
    }
}
