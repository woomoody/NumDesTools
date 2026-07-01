using System.Net.Http;
using System.Text;
using System.Text.Json;
using OfficeOpenXml;

namespace NumDesTools.Scanner;

internal static class ReportGenerator
{
    internal static string WriteBugDoc(List<PendingItem> pending)
    {
        var today = DateTime.Today.ToString("yyyy-MM-dd");
        var docPath = Path.Combine(Helpers.ConfigDir, $"bug_analysis_{today}.md");
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

    internal static void WriteReviewDoc(List<PendingItem> pending, string docPath)
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

    internal static string BuildStoryStatsSummary(List<PendingItem> pending)
    {
        if (pending.Count == 0)
            return "";

        var sb = new StringBuilder();
        sb.AppendLine("\n## 宏观统计\n");

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
            else if (
                n.Contains("二合")
                || n.Contains("砸冰")
                || n.Contains("V4")
                || n.Contains("V6")
            )
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

        var renewal = pending.Count(p =>
            p.Comment.Contains("续期")
            || System.Text.RegularExpressions.Regex.IsMatch(p.Name, @"第\d+期")
        );
        var firstTime = pending.Count - renewal;
        sb.AppendLine($"\n### 续期 vs 首期\n");
        sb.AppendLine($"- 首期/全新：**{firstTime}** 条");
        sb.AppendLine($"- 续期迭代：**{renewal}** 条");

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

        var avgTables = pending.Average(p => p.Tables.Count);
        var maxItem = pending.MaxBy(p => p.Tables.Count);
        sb.AppendLine($"\n### 配置复杂度\n");
        sb.AppendLine($"- 平均涉及表数：**{avgTables:F1}** 张");
        if (maxItem is not null)
            sb.AppendLine(
                $"- 最复杂需求：**{Helpers.Truncate(maxItem.Name.ReplaceLineEndings(" "), 40)}**（{maxItem.Tables.Count} 张表）"
            );

        return sb.ToString();
    }

    internal static string BuildBugStatsSummary(List<PendingItem> pending)
    {
        if (pending.Count == 0)
            return "";

        var sb = new StringBuilder();
        sb.AppendLine("\n## 宏观统计\n");

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
            else if (
                n.Contains("二合")
                || n.Contains("砸冰")
                || n.Contains("V4")
                || n.Contains("V6")
            )
                module = "合并玩法";
            else if (
                n.Contains("礼包")
                || n.Contains("商业化")
                || n.Contains("付费")
                || n.Contains("超级卡")
            )
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

        var symptomDefs = new (string Label, string[] Keywords)[]
        {
            (
                "多语言 / 文案",
                ["多语言", "翻译", "文案", "文字", "本地化", "语言缺失", "语言", "key"]
            ),
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
                    "云层",
                ]
            ),
            ("缺失 / 不显示", ["缺失", "缺少", "不显示", "未显示", "没有", "丢失", "消失"]),
            (
                "合成 / 链条",
                ["合成", "主链", "链条", "三合", "二合", "合并", "merge", "产出不够", "材料不够"]
            ),
            (
                "任务 / 触发",
                ["任务", "触发", "不触发", "不生效", "无法触发", "卡住", "卡死", "卡关", "寻找"]
            ),
            ("奖励 / 数值", ["奖励", "数值", "数量", "价格", "积分", "得分"]),
            ("配置 / 数据", ["配置", "数据错误", "表格", "配表", "error", "报错"]),
            ("崩溃 / 闪退", ["崩溃", "闪退", "crash", "卡顿", "ANR"]),
            ("标识 / 图标", ["标识", "图标", "icon", "角标", "标记", "蛛网"]),
            ("NPC / 地标", ["NPC", "地标", "建造", "建筑"]),
        };

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

        var writable = pending.Count(p => !p.SkipComment);
        var skipped = pending.Count(p => p.SkipComment);
        sb.AppendLine($"\n### 可分析率\n");
        sb.AppendLine(
            $"- 关键词命中（可写飞书）：**{writable}** 条 ({(double)writable / pending.Count * 100:F0}%)"
        );
        sb.AppendLine(
            $"- 关键词未命中（仅记录）：**{skipped}** 条 ({(double)skipped / pending.Count * 100:F0}%)"
        );

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

    internal static async Task GenerateSummaryAsync(string xlsxPath, string apiKey, string baseUrl)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            Console.WriteLine("[INFO] GenerateSummaryAsync: apiKey 为空，跳过摘要生成。");
            return;
        }
        if (!File.Exists(xlsxPath))
        {
            Console.WriteLine($"[INFO] GenerateSummaryAsync: 文件不存在 {xlsxPath}，跳过。");
            return;
        }

        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var rows = new List<string[]>();
        using (var pkg = new ExcelPackage(new FileInfo(xlsxPath)))
        {
            var ws = pkg.Workbook.Worksheets.FirstOrDefault();
            if (ws is null)
            {
                Console.WriteLine("[INFO] GenerateSummaryAsync: xlsx 无工作表，跳过。");
                return;
            }
            var maxRow = Math.Min(ws.Dimension?.Rows ?? 0, 20);
            var maxCol = ws.Dimension?.Columns ?? 0;
            for (var r = 1; r <= maxRow; r++)
            {
                var cells = new string[maxCol];
                for (var c = 1; c <= maxCol; c++)
                    cells[c - 1] = ws.Cells[r, c].Text ?? "";
                rows.Add(cells);
            }
        }

        var tableText = string.Join("\n", rows.Select(r => string.Join("\t", r)));
        var prompt =
            $"以下是一份竞品分析 Excel 的前 20 行数据（制表符分隔列）：\n\n{tableText}\n\n"
            + "请用中文输出 100-200 字摘要，包含：主要发现、亮点、建议关注点。";

        var requestBody = JsonSerializer.Serialize(
            new
            {
                model = "deepseek-v4-flash",
                messages = new[] { new { role = "user", content = prompt } },
                max_tokens = 512,
            }
        );

        var client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
        var response = await client.PostAsync(
            baseUrl.TrimEnd('/') + "/chat/completions",
            new StringContent(requestBody, Encoding.UTF8, "application/json")
        );
        var responseText = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            Console.WriteLine(
                $"[WARN] GenerateSummaryAsync: API 返回 {(int)response.StatusCode}，跳过写文件。"
            );
            return;
        }

        using var doc = JsonDocument.Parse(responseText);
        var summary =
            doc.RootElement.GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString()
            ?? "";

        var today = DateTime.Today.ToString("yyyyMMdd");
        var outPath = Path.Combine(OutputPaths.Analysis, $"scanner-summary-{today}.md");
        var sb = new StringBuilder();
        sb.AppendLine($"# 竞品分析摘要 {today}");
        sb.AppendLine($"生成时间：{DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"来源文件：{xlsxPath}");
        sb.AppendLine();
        sb.AppendLine(summary);
        File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8);
        Console.WriteLine($"[INFO] 竞品分析摘要已写入：{outPath}");
    }
}
