using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// Gossip Harbor 竞品核心循环分析 xlsx 生成器。
/// 数据源：C:\tmp\gossipharbor_parsed\*.json（由 parse_gh_lua.py 生成）
/// 输出：竞品-GossipHarbor核心循环分析.xlsx → Documents\workspace\
/// 包含 5 个 sheet：合并链 / 棋盘布局 / 订单配置 / 顾客NPC / 排名赛数据
/// </summary>
public static class GossipHarborWriter
{
    private const string ParsedDir = @"C:\tmp\gossipharbor_parsed";

    private const string OutFileName = "竞品-GossipHarbor核心循环分析.xlsx";

    // ── 数据加载 ──────────────────────────────────────────────────────────────
    private static (List<string> Strings, List<double> Numbers) Load(string name)
    {
        var path = Path.Combine(ParsedDir, name + ".json");
        if (!File.Exists(path))
            return ([], []);
        var doc = JsonDocument.Parse(File.ReadAllText(path));
        var root = doc.RootElement;

        List<string> strings = [];
        List<double> numbers = [];

        if (root.TryGetProperty("strings", out var sa))
            foreach (var e in sa.EnumerateArray())
                if (e.ValueKind == JsonValueKind.String)
                    strings.Add(e.GetString()!);

        if (root.TryGetProperty("numbers", out var na))
            foreach (var e in na.EnumerateArray())
                if (e.ValueKind == JsonValueKind.Number)
                    numbers.Add(e.GetDouble());

        return (strings, numbers);
    }

    // ── 公开入口 ──────────────────────────────────────────────────────────────
    public static void Run(string? outputDir = null)
    {
        var dir = outputDir ?? OutputPaths.Reports;
        var outPath = Path.Combine(dir, OutFileName);

        using var pkg = new ExcelPackage();

        BuildCoreMechSheet(pkg);
        BuildMergeChainSheet(pkg);
        BuildBoardSheet(pkg);
        BuildOrderSheet(pkg);
        BuildCustomerSheet(pkg);
        BuildBattleRobotSheet(pkg);
        BuildWeightSystemSheet(pkg);

        pkg.SaveAs(new FileInfo(outPath));
        Console.WriteLine($"[GossipHarbor] 已生成：{outPath}");
        OutputPaths.GitCommit($"[GossipHarbor] 更新竞品分析报告 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── Sheet 1：核心机制（生成器→产出链→订单）────────────────────────────────
    private static void BuildCoreMechSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("核心机制");

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【核心机制：生成器→产出链→订单对接】数据截面：adv_sewing（缝纫）关卡。棋盘pb#格=生成器直接产出的L1物品；合并链=L1→L2→…→终点；订单需求终点级物品。生成器CD/容量数值：配置加密，N/A。",
            7
        );
        ws.Row(1).Height = 70;
        ws.Row(2).Height = 22;

        // ── 区块A：生成器产出物（pb#）─────────────────────────────────────────
        var (boardStrings, _) = Load("AdventureBoardModelConfig");
        var pbItems = boardStrings
            .Where(s => s.Contains("pb#"))
            .Select(s => s.Replace("pb#", "").Replace("c#", ""))
            .Where(s => !string.IsNullOrEmpty(s))
            .Distinct()
            .OrderBy(s => s)
            .ToList();

        // 合并链数据
        var (itemStrings, itemNumbers) = Load("AdventureItemModelConfig");
        var fieldKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "Type",
            "MergedType",
            "BubbleChance",
            "Spread_Auto",
            "Spread_Weight",
            "Spread_WeightType",
            "Spread_ItemMaxNumber",
            "Spread_StorageMaxNumber",
            "Spread_CostEnergy",
            "Spread_TransformNumber",
            "ProtectLevel",
            "Transform_Weight",
            "Swallow_Weight1",
            "Swallow_Number1",
        };
        var pairs = new List<(string Type, string Merged)>();
        string? curType = null;
        for (var i = 0; i < itemStrings.Count - 1; i++)
        {
            if (itemStrings[i] == "Type" && !fieldKeys.Contains(itemStrings[i + 1]))
                curType = itemStrings[i + 1];
            else if (
                itemStrings[i] == "MergedType"
                && curType != null
                && !fieldKeys.Contains(itemStrings[i + 1])
            )
            {
                pairs.Add((curType, itemStrings[i + 1]));
                curType = null;
            }
        }
        var mergedFrom = pairs.ToDictionary(p => p.Type, p => p.Merged);
        var isNotStart = pairs.Select(p => p.Merged).ToHashSet();

        // 找每个 pb# 物品所在链的全链
        Dictionary<string, List<string>> pbChains = [];
        foreach (var pb in pbItems)
        {
            // pb 可能是链中任意位置，向上找链头
            var head = pb;
            // 反向查找：找谁 merge 到 pb
            var reverseMap = pairs.ToDictionary(p => p.Merged, p => p.Type);
            var cur = pb;
            while (reverseMap.TryGetValue(cur, out var prev))
                cur = prev;
            head = cur;

            var chain = new List<string> { head };
            var c2 = head;
            while (mergedFrom.TryGetValue(c2, out var next))
            {
                chain.Add(next);
                c2 = next;
            }
            pbChains[pb] = chain;
        }

        // Spread_Weight 数值：从字符串序列中提取 (item → weight)
        // 格式：...item_id, "Spread_Weight", item_id2(下一条的类型或数值前缀)...
        // 实际数值在 numbers 数组中，按 BubbleChance/Spread_ 字段顺序排列
        // 简化：直接读字符串数组里紧跟 "Spread_Weight" 的值（如 "main_nugget_1-2" 表示产出包）
        var spreadWeightMap = new Dictionary<string, string>();
        for (var i = 0; i < itemStrings.Count - 1; i++)
        {
            if (itemStrings[i] == "Spread_Weight" && !fieldKeys.Contains(itemStrings[i + 1]))
                spreadWeightMap[itemStrings[i - 2]] = itemStrings[i + 1]; // [i-2]=Type value
        }

        // ── 标题行（Row 2）─────────────────────────────────────────────────────
        XlsxStyleHelper.Header(ws.Cells[2, 1], "生成器产出物(pb#)", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "链位置", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 3], "链长", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 4], "链头(L1)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 5], "链尾(终点)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 6], "产出包(Spread_Weight)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 7], "生成器CD / 容量", "595959");
        ws.View.FreezePanes(3, 1);

        var row = 3;
        foreach (var pb in pbItems)
        {
            var chain = pbChains.TryGetValue(pb, out var ch) ? ch : [pb];
            var pos = chain.IndexOf(pb) + 1;
            var spreadPkg = spreadWeightMap.TryGetValue(pb, out var sp) ? sp : "—";
            XlsxStyleHelper.Cell(ws.Cells[row, 1], pb, "EBF3FB");
            ws.Cells[row, 1].Style.Font.Bold = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 2], $"L{pos}", "F2F2F2");
            XlsxStyleHelper.Cell(ws.Cells[row, 3], chain.Count, "EBF3FB");
            XlsxStyleHelper.Cell(ws.Cells[row, 4], chain[0], "D9EAD3");
            XlsxStyleHelper.Cell(ws.Cells[row, 5], chain[^1], "FFF2CC");
            XlsxStyleHelper.Cell(ws.Cells[row, 6], spreadPkg, "F0F4FA");
            XlsxStyleHelper.Cell(ws.Cells[row, 7], "N/A（配置加密）", "F5F5F5");
            ws.Cells[row, 7].Style.Font.Italic = true;
            row++;
        }

        row += 2;

        // ── 区块B：按链族分组的完整链展开 ──────────────────────────────────────
        var blockBTitle = ws.Cells[row, 1, row, 7];
        blockBTitle.Merge = true;
        blockBTitle.Value = "▶ 完整合并链展开（按族群）— 含链头、各级、终点";
        blockBTitle.Style.Font.Bold = true;
        blockBTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        blockBTitle.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("2F5496"));
        blockBTitle.Style.Font.Color.SetColor(Color.White);
        XlsxStyleHelper.Border(blockBTitle);
        ws.Row(row).Height = 22;
        row++;

        // 按 pb# 物品所在的链族分组
        var allChains = new List<List<string>>();
        foreach (var (type, _) in pairs)
        {
            if (isNotStart.Contains(type))
                continue;
            var chain = new List<string> { type };
            var c = type;
            while (mergedFrom.TryGetValue(c, out var next))
            {
                chain.Add(next);
                c = next;
            }
            allChains.Add(chain);
        }
        allChains.Sort((a, b) => b.Count.CompareTo(a.Count));

        var maxLen = allChains.Count > 0 ? allChains.Max(c => c.Count) : 1;
        XlsxStyleHelper.Header(ws.Cells[row, 1], "链头(L1)", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[row, 2], "链长", "1F4E79");
        for (var j = 0; j < Math.Min(maxLen, 5); j++)
            XlsxStyleHelper.Header(ws.Cells[row, 3 + j], $"L{j + 1}", "2F5496");
        row++;

        foreach (var chain in allChains)
        {
            var isGenChain = chain.Any(item => pbItems.Contains(item));
            var rowHex = isGenChain ? "D9EAD3" : "F5F5F5";
            XlsxStyleHelper.Cell(ws.Cells[row, 1], chain[0], rowHex);
            ws.Cells[row, 1].Style.Font.Bold = isGenChain;
            XlsxStyleHelper.Cell(ws.Cells[row, 2], chain.Count, "EBF3FB");
            for (var j = 0; j < Math.Min(chain.Count, 5); j++)
            {
                var isLast = j == chain.Count - 1;
                XlsxStyleHelper.Cell(ws.Cells[row, 3 + j], chain[j], isLast ? "FFF2CC" : "EBF3FB");
            }
            if (chain.Count > 5)
                ws.Cells[row, 7].Value = $"…+{chain.Count - 5}级";
            row++;
        }

        row += 2;

        // ── 区块C：订单机制摘要 ───────────────────────────────────────────────
        var blockCTitle = ws.Cells[row, 1, row, 7];
        blockCTitle.Merge = true;
        blockCTitle.Value = "▶ 订单机制（ChestOrderRandomConfig）— 关键参数摘要";
        blockCTitle.Style.Font.Bold = true;
        blockCTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        blockCTitle.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("2F5496"));
        blockCTitle.Style.Font.Color.SetColor(Color.White);
        XlsxStyleHelper.Border(blockCTitle);
        ws.Row(row).Height = 22;
        row++;

        var orderKv = new[]
        {
            ("订单结构", "ChestOrder — 每个订单由1~3种物品组成，物品数量有Min/Max范围"),
            ("物品数权重", "OneItemWeight / TwoItemsWeight / ThreeItemsWeight — 决定订单复杂度分布"),
            ("刷新机制", "RefreshTime 字段控制（秒），具体值在加密配置中，字段名已确认"),
            ("解锁条件", "UnlockLevel / PlayerMin / PlayerMax — 订单池按玩家等级分段"),
            ("积分奖励", "Point 字段 — 完成订单获得积分，用于关卡推进"),
            ("月卡差异", "sMonthPay / eMonthPay — 月卡玩家有独立的订单区间配置"),
            ("与生成器关系", "订单需要的物品 = 合并链终点级物品；不感知当前棋盘状态（高置信，见Sheet7）"),
            ("与动态权重关系", "生成器产出受 CurWeight=BaseW×CapFactor×ChainMultiple×RepeatFactor 调控（见Sheet7）"),
        };

        XlsxStyleHelper.Header(ws.Cells[row, 1], "机制项", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[row, 2, row, 7], "说明", "2F5496");
        ws.Cells[row, 2, row, 7].Merge = true;
        row++;

        var orderColors = new[] { "EBF3FB", "F0F4FA" };
        for (var i = 0; i < orderKv.Length; i++)
        {
            var (key, val) = orderKv[i];
            XlsxStyleHelper.Cell(ws.Cells[row, 1], key, "F2F2F2");
            ws.Cells[row, 1].Style.Font.Bold = true;
            var valCell = ws.Cells[row, 2, row, 7];
            valCell.Merge = true;
            valCell.Value = val;
            valCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            valCell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(orderColors[i % 2]));
            valCell.Style.WrapText = true;
            XlsxStyleHelper.Border(valCell);
            ws.Row(row).Height = 30;
            row++;
        }

        // 列宽
        ws.Column(1).Width = 22;
        ws.Column(2).Width = 20;
        ws.Column(3).Width = 20;
        ws.Column(4).Width = 20;
        ws.Column(5).Width = 20;
        ws.Column(6).Width = 20;
        ws.Column(7).Width = 20;
    }

    // ── Sheet 2：合并链 ───────────────────────────────────────────────────────
    private static void BuildMergeChainSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("合并链");

        var (strings, _) = Load("AdventureItemModelConfig");

        // 从常量字符串中提取 Type→MergedType 对
        // 格式：..., "Type", "<item>", "MergedType", "<next_item>", ...
        var pairs = new List<(string Type, string Merged)>();
        var fieldKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "Type",
            "MergedType",
            "BubbleChance",
            "Spread_Auto",
            "Spread_Weight",
            "Spread_WeightType",
            "Spread_ItemMaxNumber",
            "Spread_StorageMaxNumber",
            "Spread_CostEnergy",
            "Spread_TransformNumber",
            "ProtectLevel",
            "Transform_Weight",
            "Swallow_Weight1",
            "Swallow_Number1",
        };

        string? currentType = null;
        for (var i = 0; i < strings.Count - 1; i++)
        {
            if (strings[i] == "Type" && !fieldKeys.Contains(strings[i + 1]))
                currentType = strings[i + 1];
            else if (
                strings[i] == "MergedType"
                && currentType != null
                && !fieldKeys.Contains(strings[i + 1])
            )
            {
                pairs.Add((currentType, strings[i + 1]));
                currentType = null;
            }
        }

        // 构建链：Type→Merged→Merged→...
        var mergedFrom = pairs.ToDictionary(p => p.Type, p => p.Merged);
        var isNotStart = pairs.Select(p => p.Merged).ToHashSet();
        var chains = new List<List<string>>();
        foreach (var (type, _) in pairs)
        {
            if (isNotStart.Contains(type))
                continue;
            var chain = new List<string> { type };
            var cur = type;
            while (mergedFrom.TryGetValue(cur, out var next))
            {
                chain.Add(next);
                cur = next;
            }
            chains.Add(chain);
        }
        chains.Sort((a, b) => b.Count.CompareTo(a.Count));

        var maxLen = chains.Count > 0 ? chains.Max(c => c.Count) : 1;
        var colCount = 2 + maxLen;

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【合并链】数据源：AdventureItemModelConfig（Frida内存扫描）。Type→MergedType逐级追踪，形成完整合并链。链长=合并级数，最终物品高亮黄色。",
            colCount
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        XlsxStyleHelper.Header(ws.Cells[2, 1], "链序号", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "链长", "1F4E79");
        for (var j = 0; j < maxLen; j++)
            XlsxStyleHelper.Header(ws.Cells[2, 3 + j], $"Lv.{j + 1}", "2F5496");
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        // 数据行（从 row 3 起）
        for (var i = 0; i < chains.Count; i++)
        {
            var row = i + 3;
            var chain = chains[i];
            XlsxStyleHelper.Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            XlsxStyleHelper.Cell(ws.Cells[row, 2], chain.Count, "EBF3FB");
            for (var j = 0; j < chain.Count; j++)
            {
                var color = j == chain.Count - 1 ? "FFF2CC" : "EBF3FB";
                XlsxStyleHelper.Cell(ws.Cells[row, 3 + j], chain[j], color);
            }
        }

        ws.Column(1).Width = 8;
        ws.Column(2).Width = 8;
        for (var j = 0; j < maxLen; j++)
            ws.Column(3 + j).Width = 26;

        // 统计行
        var statRow = chains.Count + 4;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"共 {chains.Count} 条链";
        ws.Cells[statRow, 3].Value = $"最长 {maxLen} 级";
        ws.Cells[statRow, 4].Value = $"总元素 {pairs.Count} 个";
    }

    // ── Sheet 2：棋盘布局 ────────────────────────────────────────────────────
    private static void BuildBoardSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("棋盘布局");
        var (strings, _) = Load("AdventureBoardModelConfig");

        // 格子格式：[pb#][c#]<item_name>
        // pb# = producer bubble（生成器掉落）
        // c#  = cloud locked（初始锁定，开云后出现）
        var cells = strings.Where(s => !string.IsNullOrEmpty(s)).ToList();

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【棋盘布局】数据源：AdventureBoardModelConfig。pb#=生成器初始产出格，c#=云锁格（云消除后出现）。元素名为去前缀后的物品ID。",
            4
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        XlsxStyleHelper.Header(ws.Cells[2, 1], "格子编码", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "是否初始生成(pb)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 3], "是否锁定(c)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 4], "元素名称", "2F5496");
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        for (var i = 0; i < cells.Count; i++)
        {
            var raw = cells[i];
            var row = i + 3;
            var isPb = raw.Contains("pb#");
            var isC = Regex.IsMatch(raw, @"(?<![a-z])c#");
            var itemName = raw.Replace("pb#", "").Replace("c#", "");

            XlsxStyleHelper.Cell(ws.Cells[row, 1], raw);
            XlsxStyleHelper.Cell(ws.Cells[row, 2], isPb ? "✓" : "", isPb ? "D9EAD3" : null);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], isC ? "✓" : "", isC ? "FCE5CD" : null);
            XlsxStyleHelper.Cell(ws.Cells[row, 4], itemName, "EBF3FB");
        }

        ws.Column(1).Width = 36;
        ws.Column(2).Width = 18;
        ws.Column(3).Width = 14;
        ws.Column(4).Width = 30;

        // 统计
        var statRow = cells.Count + 4;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"总格子数: {cells.Count}";
        ws.Cells[statRow, 3].Value = $"初始生成: {cells.Count(c => c.Contains("pb#"))}";
        ws.Cells[statRow, 4].Value = $"锁定格: {cells.Count(c => Regex.IsMatch(c, @"(?<![a-z])c#"))}";
    }

    // ── Sheet 3：订单配置 ────────────────────────────────────────────────────
    private static void BuildOrderSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("订单配置");
        var (strings, numbers) = Load("ChestOrderRandomConfig");

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【订单配置】数据源：ChestOrderRandomConfig（Frida内存扫描）。字段名为Lua常量字符串，数值为对应配置值。RefreshTime=刷新间隔秒数，Weight=权重，Min/Max=物品数量范围。",
            3
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        XlsxStyleHelper.Header(ws.Cells[2, 1], "字段名（字符串常量）", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "数值常量（按序）", "2F5496");
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        var maxRows = Math.Max(strings.Count, numbers.Count);
        for (var i = 0; i < maxRows; i++)
        {
            var row = i + 3;
            if (i < strings.Count)
                XlsxStyleHelper.Cell(ws.Cells[row, 1], strings[i], "EBF3FB");
            if (i < numbers.Count)
                XlsxStyleHelper.Cell(ws.Cells[row, 2], numbers[i], "FFF2CC");
        }

        // 字段说明（已知含义）
        var notes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["Type"] = "订单类型 ID",
            ["RefreshTime"] = "刷新间隔（秒），3600=1h，1800=30m",
            ["UnlockLevel"] = "解锁关卡",
            ["Point"] = "完成订单奖励积分",
            ["Filters"] = "物品过滤器（允许哪类物品）",
            ["PlayerMin"] = "玩家等级下限",
            ["PlayerMax"] = "玩家等级上限",
            ["ItemAmounts"] = "单次订单物品数",
            ["OneItemWeight"] = "1种物品权重",
            ["TwoItemsWeight"] = "2种物品权重",
            ["ThreeItemsWeight"] = "3种物品权重",
            ["FirstMin"] = "第1件物品最小数量",
            ["FirstMax"] = "第1件物品最大数量",
            ["SecondMin"] = "第2件物品最小数量",
            ["SecondMax"] = "第2件物品最大数量",
            ["ThirdMin"] = "第3件物品最小数量",
            ["ThirdMax"] = "第3件物品最大数量",
            ["sMonthPay"] = "月卡玩家订单区间起",
            ["eMonthPay"] = "月卡玩家订单区间止",
        };

        XlsxStyleHelper.Header(ws.Cells[2, 3], "字段说明", "2F5496");
        for (var i = 0; i < strings.Count; i++)
        {
            var row = i + 3;
            if (notes.TryGetValue(strings[i], out var note))
                XlsxStyleHelper.Cell(ws.Cells[row, 3], note, "F4CCCC");
        }

        ws.Column(1).Width = 22;
        ws.Column(2).Width = 18;
        ws.Column(3).Width = 36;
    }

    // ── Sheet 4：顾客 NPC ────────────────────────────────────────────────────
    private static void BuildCustomerSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("顾客NPC");
        var (strings, _) = Load("CustomerConfigName");

        // 过滤掉 Lua 框架字符串，只留顾客名
        var luaKeywords = new HashSet<string>
        {
            "CustomerConfigName",
            "GameConfig",
            "IsTestMode",
            "setmetatable",
            "__index",
            "HasConfig",
            "Log",
            "Error",
            "tostring",
            "rawget",
        };
        var customers = strings
            .Where(s => !luaKeywords.Contains(s) && !s.Contains(' ') && !s.Contains(':'))
            .ToList();

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【顾客NPC】数据源：CustomerConfigName（Frida内存扫描）。已过滤Lua框架关键字，仅保留顾客ID字符串。备注列可手动补充中文名或角色说明。",
            3
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        XlsxStyleHelper.Header(ws.Cells[2, 1], "#", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "顾客名（英文）", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 3], "备注", "2F5496");
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        var colors = new[] { "EBF3FB", "D9EAD3", "FFF2CC", "FCE5CD", "E8DAEF", "D5DBDB" };
        for (var i = 0; i < customers.Count; i++)
        {
            var row = i + 3;
            var hex = colors[i % colors.Length];
            XlsxStyleHelper.Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            XlsxStyleHelper.Cell(ws.Cells[row, 2], customers[i], hex);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], "");
        }

        ws.Column(1).Width = 6;
        ws.Column(2).Width = 18;
        ws.Column(3).Width = 30;
    }

    // ── Sheet 5：排名赛数据 ──────────────────────────────────────────────────
    private static void BuildBattleRobotSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("排名赛数据");

        var (strings, numbers) = Load("BattleRobotConfig");

        // 格式：score-time-difficulty (e.g. "201-3-2")
        // difficulty: 0=简单, 1=普通, 2=困难
        var entries = strings
            .Select(s =>
            {
                var m = Regex.Match(s, @"^(\d+)-(\d+)-([012])$");
                if (!m.Success)
                    return null;
                return new
                {
                    Score = int.Parse(m.Groups[1].Value),
                    Time = int.Parse(m.Groups[2].Value),
                    Difficulty = int.Parse(m.Groups[3].Value),
                };
            })
            .Where(e => e != null)
            .Select(e => e!)
            .OrderBy(e => e.Difficulty)
            .ThenBy(e => e.Score)
            .ToList();

        // groupId 数值（排名赛分组）
        var groupIds = numbers
            .Where(n => n > 100000)
            .Select(n => (int)n)
            .Distinct()
            .Order()
            .ToList();

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【排名赛数据】数据源：BattleRobotConfig（Frida内存扫描）。格式 score-time-difficulty，difficulty: 0=Easy, 1=Normal, 2=Hard。分钟速率=score/time，GroupId格式10x00y推测x=轮次。",
            8
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        XlsxStyleHelper.Header(ws.Cells[2, 1], "#", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 2], "分数", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, 3], "时间(分)", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 4], "难度", "2F5496");
        XlsxStyleHelper.Header(ws.Cells[2, 5], "分钟得分速率", "2F5496");
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        var diffColors = new Dictionary<int, string>
        {
            [0] = "D9EAD3",
            [1] = "FFF2CC",
            [2] = "FADADD"
        };
        var diffNames = new Dictionary<int, string>
        {
            [0] = "Easy",
            [1] = "Normal",
            [2] = "Hard"
        };

        for (var i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            var row = i + 3;
            var hex = diffColors.GetValueOrDefault(e.Difficulty, "EBF3FB");
            XlsxStyleHelper.Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            XlsxStyleHelper.Cell(ws.Cells[row, 2], e.Score, hex);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], e.Time, hex);
            XlsxStyleHelper.Cell(
                ws.Cells[row, 4],
                diffNames.GetValueOrDefault(e.Difficulty, e.Difficulty.ToString()),
                hex
            );
            var rate = e.Time > 0 ? Math.Round((double)e.Score / e.Time, 1) : 0;
            XlsxStyleHelper.Cell(ws.Cells[row, 5], rate, "EBF3FB");
        }

        ws.Column(1).Width = 6;
        ws.Column(2).Width = 14;
        ws.Column(3).Width = 12;
        ws.Column(4).Width = 10;
        ws.Column(5).Width = 16;

        // 右侧：分组 ID 列表
        var gCol = 7;
        XlsxStyleHelper.Header(ws.Cells[2, gCol], "GroupId", "1F4E79");
        XlsxStyleHelper.Header(ws.Cells[2, gCol + 1], "分组轮数", "2F5496");
        for (var i = 0; i < groupIds.Count; i++)
        {
            var row = i + 3;
            XlsxStyleHelper.Cell(ws.Cells[row, gCol], groupIds[i], "EBF3FB");
            // groupId 格式：10x00y → x 是轮次，y 是序号（推测）
            var round = groupIds[i] / 100000;
            XlsxStyleHelper.Cell(ws.Cells[row, gCol + 1], $"Round {round}", "F2F2F2");
        }

        // 统计摘要
        var statRow = entries.Count + 4;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"总条目: {entries.Count}";
        ws.Cells[statRow, 3].Value = $"Easy: {entries.Count(e => e.Difficulty == 0)}";
        ws.Cells[statRow, 4].Value = $"Normal: {entries.Count(e => e.Difficulty == 1)}";
        ws.Cells[statRow, 5].Value = $"Hard: {entries.Count(e => e.Difficulty == 2)}";
    }

    // ── Sheet 6：动态权重系统（Frida内存扫描结论）──────────────────────────────
    private static void BuildWeightSystemSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("动态权重系统");

        // Row 1：机制说明
        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【动态权重系统】Frida内存扫描置信度报告（2026-05-14，PID=3199）。CurWeight=BaseW×CapFactor×ChainMultiple×RepeatFactor。BaseW/ChainMultiple高置信，CapFactor中置信（隐式推断），BoardAware高置信负向（订单不感知棋盘）。",
            6
        );
        ws.Row(1).Height = 70;

        // Row 2：标题
        string[] headers = ["问题", "结论", "置信度", "关键证据字段/函数", "具体数值", "备注"];
        string[] hdrHex = ["1F4E79", "2F5496", "595959", "2F5496", "595959", "595959"];
        for (var j = 0; j < headers.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);
        ws.Row(2).Height = 22;

        ws.View.FreezePanes(3, 1);

        // 4大问题的扫描结论
        var findings = new[]
        {
            (
                "BaseW（基础权重值）",
                "找到 — 运行时字段确认",
                "高",
                "Spread_Weight(配置表) / spreadWeight(运行时JSON)",
                "pack_common_1: 75/25; pack_common_2: 55/45; pack_large: 60/…",
                "与 AdventureItemModelConfig.Spread_Weight 字段对应；gacha包权重用同一格式"
            ),
            (
                "CapFactor（容量上限因子）",
                "隐式实现 — 无独立变量名",
                "中",
                "Spread_ItemMaxNumber(配置) / spreadItemRestNumber(运行时) / _UpdateSpreadConfigSmartWeight",
                "CapFactor = 1 - spreadItemRestNumber/Spread_ItemMaxNumber（推断公式）",
                "无 'CapFactor' 字面量；JSON实例: {\"spreadItemRestNumber\":0,\"spreadState\":3,...}；权重更新入口已确认"
            ),
            (
                "ChainMultiple（链重复倍率）",
                "找到 — 变量名确认",
                "高",
                "itemChainMultiple / mapChainRepeatWeight / repeatWeightDecrease / mapItemCodeMultiple",
                "具体倍率数值未以明文浮点出现；repeatWeightDecrease证实存在重复惩罚递减",
                "位于 MainOrderCreatorRandom 模块 (addr 0x7ef20f14d528)；与多重字段共存"
            ),
            (
                "BoardAware（订单是否感知棋盘）",
                "否 — 高置信负向结论",
                "高（负）",
                "GetCachedItemCount（仅在spread/freefall，从未在订单附近） / GetBoardItemCount(全扫描未命中)",
                "N/A — 订单创建只关心 requireItems 和时间参数",
                "6批次扫描(r--/r-x/rw-)均未命中GetBoardItemCount；CreateOrder/CanCreateOrder/MainOrderCreatorRandom无棋盘调用"
            ),
        };

        var confColors = new Dictionary<string, string>
        {
            ["高"] = "D9EAD3",
            ["中"] = "FFF2CC",
            ["高（负）"] = "EBF3FB",
        };

        var row = 3;
        foreach (var (q, conclusion, conf, evidence, value, note) in findings)
        {
            var confHex = confColors.GetValueOrDefault(conf, "F5F5F5");
            XlsxStyleHelper.Cell(ws.Cells[row, 1], q, "F0F4FA");
            ws.Cells[row, 1].Style.Font.Bold = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 2], conclusion, confHex);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], conf, confHex);
            XlsxStyleHelper.Cell(ws.Cells[row, 4], evidence, "FAFAFA");
            ws.Cells[row, 4].Style.WrapText = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 5], value, "FFF8E8");
            ws.Cells[row, 5].Style.WrapText = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 6], note, "F5F5F5");
            ws.Cells[row, 6].Style.WrapText = true;
            ws.Row(row).Height = 52;
            row++;
        }

        row += 2;

        // 4因子 CurWeight 公式块
        var formulaTitle = ws.Cells[row, 1, row, 6];
        formulaTitle.Merge = true;
        formulaTitle.Value = "■ CurWeight 4因子公式（逆向推导，置信度：中-高）";
        formulaTitle.Style.Font.Bold = true;
        formulaTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        formulaTitle.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("4A235A"));
        formulaTitle.Style.Font.Color.SetColor(Color.White);
        XlsxStyleHelper.Border(formulaTitle);
        ws.Row(row).Height = 22;
        row++;

        var formulaFactors = new[]
        {
            (
                "BaseW",
                "Spread_Weight（配置表字段）",
                "高",
                "每条链在AdventureItemModelConfig配置的基础权重，例如 75 / 25",
                "直接运行时字段 spreadWeight 观测到"
            ),
            (
                "CapFactor",
                "1 - spreadItemRestNumber / Spread_ItemMaxNumber",
                "中",
                "当前链产出积压量越多，CapFactor越低，抑制继续产出同类元素",
                "无显式变量，从 _UpdateSpreadConfigSmartWeight + 字段组合推断"
            ),
            (
                "ChainMultiple",
                "itemChainMultiple（运行时）",
                "高（名称）/ 中（值）",
                "对同一链重复出现施加衰减倍率，mapChainRepeatWeight存储各链重复次数",
                "函数名确认，倍率数值未明文出现；repeatWeightDecrease证实递减机制"
            ),
            (
                "RepeatFactor",
                "推断 = f(mapItemCodeMultiple)",
                "中",
                "单品ID级别重复惩罚，防止同一具体元素连续产出",
                "mapItemCodeMultiple 字段与 itemChainMultiple 并列，同一模块"
            ),
        };

        string[] factorHeaders = ["因子名", "字段/公式", "置信度", "语义", "证据来源"];
        for (var j = 0; j < factorHeaders.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[row, j + 1], factorHeaders[j], "2F5496");
        row++;

        var fColors = new[] { "EBF3FB", "FFF2CC", "D9EAD3", "FCE5CD" };
        for (var fi = 0; fi < formulaFactors.Length; fi++)
        {
            var (fname, formula, fconf, fmean, fevidence) = formulaFactors[fi];
            var fhex = fColors[fi % fColors.Length];
            XlsxStyleHelper.Cell(ws.Cells[row, 1], fname, fhex);
            ws.Cells[row, 1].Style.Font.Bold = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 2], formula, "FAFAFA");
            ws.Cells[row, 2].Style.WrapText = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 3], fconf, fconf.StartsWith("高") ? "D9EAD3" : "FFF2CC");
            XlsxStyleHelper.Cell(ws.Cells[row, 4], fmean, "F5F5F5");
            ws.Cells[row, 4].Style.WrapText = true;
            XlsxStyleHelper.Cell(ws.Cells[row, 5], fevidence, "F0F0F0");
            ws.Cells[row, 5].Style.WrapText = true;
            ws.Row(row).Height = 38;
            row++;
        }

        row += 2;

        // 总结说明
        var summaryCell = ws.Cells[row, 1, row, 6];
        summaryCell.Merge = true;
        summaryCell.Value =
            "【总结】4因子动态权重已高置信度确认存在（BaseW+ChainMultiple命中，CapFactor隐式推断，RepeatFactor字段并列）。"
            + "订单不感知棋盘状态（高置信负向，6批扫描均未命中GetBoardItemCount于订单路径）。"
            + "这与 Merge Cooking 的静态概率池和 TravelTown 的固定产出形成鲜明对比——"
            + "Harbor是三款中运行时动态性最高的，对订单池设计成本最高，也是最难平衡的。"
            + "主要未获取：各因子的具体数值参数（配置表加密，内存中以IL2CPP偏移量而非字符串字段存储）。";
        summaryCell.Style.Font.Size = 10;
        summaryCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        summaryCell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("EBF3FB"));
        summaryCell.Style.WrapText = true;
        summaryCell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        XlsxStyleHelper.Border(summaryCell);
        ws.Row(row).Height = 60;

        ws.Column(1).Width = 18;
        ws.Column(2).Width = 26;
        ws.Column(3).Width = 12;
        ws.Column(4).Width = 32;
        ws.Column(5).Width = 28;
        ws.Column(6).Width = 30;
    }
}
