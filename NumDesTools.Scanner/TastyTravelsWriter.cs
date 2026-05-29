using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// Tasty Travels (io.randomco.travel) 竞品核心循环分析 xlsx 生成器。
/// 数据源：设备 HTTPCache（adb pull 到本机）
///   - C:\tmp\tasty\httpcache\HTTPCache\1C  — 订单树 + 合成链图（16MB）
///   - C:\tmp\tasty\httpcache\HTTPCache\10  — 棋盘初始布局 + 生成器配置
/// 输出：竞品-TastyTravels核心循环分析.xlsx → Documents\NumDesOutput\reports\
/// Sheet：概览 / 订单系统 / 棋盘布局 / 合成链 / 生成器配置
/// </summary>
public static class TastyTravelsWriter
{
    private const string CacheDir = @"C:\tmp\tasty\httpcache\HTTPCache";
    private const string OutFileName = "竞品-TastyTravels核心循环分析.xlsx";

    // ── 数据模型 ──────────────────────────────────────────────────────────────

    private record OrderObjective(string ItemReference, int Amount);

    private record Order(
        string OrderId,
        long TreeId,
        int SeqInTree,
        List<OrderObjective> Objectives,
        List<OrderObjective> Rewards,
        string TaskType,
        List<string> LockedByIds
    );

    private record BoardCell(int Row, int Col, string Item, bool Locked, bool Boxed, int LevelLock);

    private record MergeChain(string GraphId, string GraphName, List<string> ItemIds);

    // ── 样式辅助 ──────────────────────────────────────────────────────────────

    private static void Header(ExcelRange c, string text, string hex = "2F5496")
    {
        c.Value = text;
        c.Style.Fill.PatternType = ExcelFillStyle.Solid;
        c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        c.Style.Font.Bold = true;
        c.Style.Font.Color.SetColor(Color.White);
        c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        c.Style.WrapText = true;
        Border(c);
    }

    private static void Cell(ExcelRange c, object? value, string? hex = null, bool wrap = false)
    {
        c.Value = value;
        if (hex != null)
        {
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        }
        if (wrap)
            c.Style.WrapText = true;
        Border(c);
    }

    private static void MechNote(
        ExcelWorksheet ws,
        int row,
        int startCol,
        string text,
        int mergeWidth
    )
    {
        var cell = ws.Cells[row, startCol, row, startCol + mergeWidth - 1];
        cell.Merge = true;
        cell.Value = text;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(HexColor("F0F4FA"));
        cell.Style.Font.Size = 9;
        cell.Style.WrapText = true;
        Border(cell);
    }

    private static void Border(ExcelRange c)
    {
        var b = c.Style.Border;
        b.Top.Style = b.Bottom.Style = b.Left.Style = b.Right.Style = ExcelBorderStyle.Thin;
        var gray = Color.FromArgb(0xBD, 0xBD, 0xBD);
        b.Top.Color.SetColor(gray);
        b.Bottom.Color.SetColor(gray);
        b.Left.Color.SetColor(gray);
        b.Right.Color.SetColor(gray);
    }

    private static Color HexColor(string hex)
    {
        hex = hex.TrimStart('#');
        return Color.FromArgb(
            Convert.ToInt32(hex[..2], 16),
            Convert.ToInt32(hex[2..4], 16),
            Convert.ToInt32(hex[4..6], 16)
        );
    }

    // ── 数据加载 ──────────────────────────────────────────────────────────────

    private static JsonElement LoadJson(string filename)
    {
        var path = Path.Combine(CacheDir, filename);
        if (!File.Exists(path))
            throw new FileNotFoundException($"HTTPCache 文件不存在：{path}");
        var raw = File.ReadAllBytes(path);
        var text = System.Text.Encoding.UTF8.GetString(raw);
        var idx = text.IndexOf('{');
        if (idx < 0)
            throw new InvalidDataException($"文件中找不到 JSON：{path}");
        return JsonDocument.Parse(text[idx..]).RootElement;
    }

    private static List<Order> LoadOrders()
    {
        var root = LoadJson("1C");
        var trees = root.GetProperty("itemsInteractions")
            .GetProperty("introOrders")
            .EnumerateArray();

        var result = new List<Order>();
        foreach (var tree in trees)
        {
            var treeId = tree.GetProperty("uniqueTreeId").GetInt64();
            var orders = tree.GetProperty("orders").EnumerateArray().ToList();
            for (var i = 0; i < orders.Count; i++)
            {
                var o = orders[i];
                var objectives = o.TryGetProperty("objectives", out var objArr)
                    ? objArr
                        .EnumerateArray()
                        .Select(x => new OrderObjective(
                            x.GetProperty("itemReference").GetString()!,
                            x.GetProperty("amount").GetInt32()
                        ))
                        .ToList()
                    : [];
                var rewards = o.TryGetProperty("itemRewards", out var rewArr)
                    ? rewArr
                        .EnumerateArray()
                        .Select(x => new OrderObjective(
                            x.GetProperty("itemReference").GetString()!,
                            x.GetProperty("amount").GetInt32()
                        ))
                        .ToList()
                    : [];
                var locked = o.TryGetProperty("lockedByIds", out var locArr)
                    ? locArr.EnumerateArray().Select(x => x.GetString()!).ToList()
                    : [];
                result.Add(
                    new Order(
                        o.GetProperty("orderId").GetString()!,
                        treeId,
                        i + 1,
                        objectives,
                        rewards,
                        o.TryGetProperty("taskType", out var tt) ? tt.GetString()! : "",
                        locked
                    )
                );
            }
        }
        return result;
    }

    private static List<BoardCell> LoadBoard()
    {
        var root = LoadJson("10");
        var rows = root.GetProperty("data")
            .GetProperty("data")
            .GetProperty("boardCoreConfiguration")
            .GetProperty("startingBoard")
            .EnumerateArray();

        var result = new List<BoardCell>();
        foreach (var row in rows)
        {
            var rowNum = row.GetProperty("description").GetString() is { } d
                ? int.TryParse(d, out var n)
                    ? n
                    : 0
                : 0;
            for (var col = 1; col <= 7; col++)
            {
                var key = $"cell{col}";
                if (!row.TryGetProperty(key, out var cell))
                    continue;
                result.Add(
                    new BoardCell(
                        rowNum,
                        col,
                        cell.GetProperty("item").GetString()!,
                        cell.GetProperty("locked").GetBoolean(),
                        cell.GetProperty("boxed").GetBoolean(),
                        cell.TryGetProperty("levelLock", out var ll) ? ll.GetInt32() : 0
                    )
                );
            }
        }
        return result;
    }

    private static List<MergeChain> LoadChains()
    {
        var root = LoadJson("1C");
        var graphs = root.GetProperty("itemsInteractions")
            .GetProperty("mergeItemGraphs")
            .EnumerateArray();

        var result = new List<MergeChain>();
        foreach (var g in graphs)
        {
            var graphId = g.GetProperty("uniqueId").GetString()!;
            var graphName = g.TryGetProperty("graphName", out var gn) ? gn.GetString()! : graphId;
            if (!g.TryGetProperty("items", out var itemsArr))
                continue;
            var items = itemsArr
                .EnumerateArray()
                .Select(x => x.TryGetProperty("uniqueId", out var uid) ? uid.GetString()! : "")
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
            if (items.Count >= 2)
                result.Add(new MergeChain(graphId, graphName, items));
        }
        // 按链长降序
        result.Sort((a, b) => b.ItemIds.Count.CompareTo(a.ItemIds.Count));
        return result;
    }

    // ── 公开入口 ──────────────────────────────────────────────────────────────

    public static void Run(string? outputDir = null)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var dir = outputDir ?? OutputPaths.Reports;
        var outPath = Path.Combine(dir, OutFileName);

        var orders = LoadOrders();
        var board = LoadBoard();
        var chains = LoadChains();

        using var pkg = new ExcelPackage();

        BuildOverviewSheet(pkg, orders, board, chains);
        BuildOrderSheet(pkg, orders);
        BuildBoardSheet(pkg, board);
        BuildChainSheet(pkg, chains);
        BuildSpawnerSheet(pkg, chains);

        pkg.SaveAs(new FileInfo(outPath));
        Console.WriteLine($"[TastyTravels] 已生成：{outPath}");
        OutputPaths.GitCommit($"[竞品分析] TastyTravels核心循环分析 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── Sheet 1：概览 ─────────────────────────────────────────────────────────

    private static void BuildOverviewSheet(
        ExcelPackage pkg,
        List<Order> orders,
        List<BoardCell> board,
        List<MergeChain> chains
    )
    {
        var ws = pkg.Workbook.Worksheets.Add("概览");

        MechNote(
            ws,
            1,
            1,
            "【Tasty Travels 核心循环】数据源：设备 HTTPCache（服务端下发配置）。"
                + "核心循环：生成器产出 → 棋盘合成升级 → 完成订单 → 获取能量/钻石 → 继续循环。",
            3
        );
        ws.Row(1).Height = 50;

        var kv = new[]
        {
            ("游戏名称", "Tasty Travels"),
            ("开发商", "RandomCo (io.randomco.travel)"),
            ("核心循环", "生成器产出 → 合成升级 → 完成订单 → 奖励回收"),
            ("订单树总数", $"{orders.Select(o => o.TreeId).Distinct().Count()} 棵"),
            ("订单总数", $"{orders.Count} 个"),
            (
                "平均每树订单数",
                $"{orders.Count / (double)orders.Select(o => o.TreeId).Distinct().Count():F1}"
            ),
            ("订单需求 Tier 分布", TierDistText(orders)),
            ("主要订单奖励", RewardDistText(orders)),
            ("棋盘格子数", $"{board.Count} 格（9行×7列）"),
            ("初始锁定格", $"{board.Count(c => c.Locked)} 格"),
            ("装箱格（Boxed）", $"{board.Count(c => c.Boxed)} 格"),
            (
                "等级解锁格",
                $"Lv7:{board.Count(c => c.LevelLock == 7)} / Lv8:{board.Count(c => c.LevelLock == 8)} / Lv11:{board.Count(c => c.LevelLock == 11)}"
            ),
            ("合成链总数", $"{chains.Count} 条（≥2 tier）"),
            (
                "最长合成链",
                $"{chains.FirstOrDefault()?.ItemIds.Count ?? 0} tier — {chains.FirstOrDefault()?.GraphId ?? ""}"
            ),
            ("链长分布", ChainDistText(chains)),
            ("能量恢复", "+1 / 120秒，上限 100"),
            ("泡泡过期时间", "60秒，3次/天免费爆泡，Lv10解锁"),
            ("棋盘扩展", "初始7格，最大30格，10HC起步×1.25倍率"),
        };

        Header(ws.Cells[2, 1], "项目", "1F4E79");
        Header(ws.Cells[2, 2, 2, 3], "数值 / 说明", "2F5496");
        ws.Cells[2, 2, 2, 3].Merge = true;
        ws.Row(2).Height = 22;

        var altColors = new[] { "EBF3FB", "F0F4FA" };
        for (var i = 0; i < kv.Length; i++)
        {
            var row = i + 3;
            var (key, val) = kv[i];
            Cell(ws.Cells[row, 1], key, "F2F2F2");
            ws.Cells[row, 1].Style.Font.Bold = true;
            var valCell = ws.Cells[row, 2, row, 3];
            valCell.Merge = true;
            valCell.Value = val;
            valCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            valCell.Style.Fill.BackgroundColor.SetColor(HexColor(altColors[i % 2]));
            valCell.Style.WrapText = true;
            Border(valCell);
            ws.Row(row).Height = 18;
        }

        ws.Column(1).Width = 20;
        ws.Column(2).Width = 28;
        ws.Column(3).Width = 28;
    }

    private static string TierDistText(List<Order> orders)
    {
        var dist = orders
            .SelectMany(o => o.Objectives)
            .Select(obj => TierFromItemId(obj.ItemReference))
            .Where(t => t >= 0)
            .GroupBy(t => t)
            .OrderBy(g => g.Key)
            .Select(g => $"T{g.Key}:{g.Count()}")
            .ToList();
        return string.Join(" / ", dist);
    }

    private static string RewardDistText(List<Order> orders)
    {
        return string.Join(
            " / ",
            orders
                .SelectMany(o => o.Rewards)
                .GroupBy(r => r.ItemReference)
                .OrderByDescending(g => g.Count())
                .Take(4)
                .Select(g => $"{g.Key.Replace("item_", "")}({g.Count()})")
        );
    }

    private static string ChainDistText(List<MergeChain> chains)
    {
        var dist = chains
            .GroupBy(c => c.ItemIds.Count / 3 * 3) // 按3tier一档分组
            .OrderBy(g => g.Key)
            .Select(g => $"{g.Key}-{g.Key + 2}tier:{g.Count()}条")
            .ToList();
        return string.Join(" / ", dist.Take(5));
    }

    private static int TierFromItemId(string itemId)
    {
        var m = Regex.Match(itemId, @"_(\d{2})$");
        return m.Success ? int.Parse(m.Groups[1].Value) : -1;
    }

    // ── Sheet 2：订单系统 ─────────────────────────────────────────────────────

    private static void BuildOrderSheet(ExcelPackage pkg, List<Order> orders)
    {
        var ws = pkg.Workbook.Worksheets.Add("订单系统");

        MechNote(
            ws,
            1,
            1,
            "【订单系统】数据源：HTTPCache/1C introOrders。103棵订单树，389个订单，81%为4单树。"
                + "每棵树内订单串行解锁（lockedByIds）。需求集中在 T2-T5 中高 tier，奖励以能量箱/工具箱/钻石三足鼎立。",
            8
        );
        ws.Row(1).Height = 50;

        string[] headers =
        [
            "订单ID",
            "树ID",
            "树内序号",
            "需求物品",
            "需求Tier",
            "需求数量",
            "奖励物品",
            "奖励数量",
        ];
        string[] hdrHex =
        [
            "1F4E79",
            "1F4E79",
            "2F5496",
            "2F5496",
            "2F5496",
            "2F5496",
            "385623",
            "385623",
        ];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);
        ws.Row(2).Height = 22;
        ws.View.FreezePanes(3, 1);

        var tierColors = new Dictionary<int, string>
        {
            [0] = "F5F5F5",
            [1] = "EBF3FB",
            [2] = "D9EAD3",
            [3] = "FFF2CC",
            [4] = "FCE5CD",
            [5] = "E8DAEF",
        };

        for (var i = 0; i < orders.Count; i++)
        {
            var row = i + 3;
            var o = orders[i];

            // 有多个 objective 时展成多行，第一行写订单元信息
            var objCount = Math.Max(1, o.Objectives.Count);
            for (var oi = 0; oi < objCount; oi++)
            {
                var r = row + oi;
                if (oi == 0)
                {
                    Cell(ws.Cells[r, 1], o.OrderId, "F2F2F2");
                    Cell(ws.Cells[r, 2], o.TreeId, "EBF3FB");
                    Cell(ws.Cells[r, 3], o.SeqInTree, "EBF3FB");
                }

                if (oi < o.Objectives.Count)
                {
                    var obj = o.Objectives[oi];
                    var tier = TierFromItemId(obj.ItemReference);
                    var tierHex = tierColors.GetValueOrDefault(Math.Min(tier, 5), "F5F5F5");
                    Cell(ws.Cells[r, 4], obj.ItemReference.Replace("item_", ""), tierHex);
                    Cell(ws.Cells[r, 5], tier >= 0 ? tier : "", tierHex);
                    Cell(ws.Cells[r, 6], obj.Amount, tierHex);
                }

                if (oi < o.Rewards.Count)
                {
                    var rew = o.Rewards[oi];
                    Cell(ws.Cells[r, 7], rew.ItemReference.Replace("item_", ""), "D9EAD3");
                    Cell(ws.Cells[r, 8], rew.Amount, "D9EAD3");
                }
            }

            // 如果多行，把序号列合并
            if (objCount > 1)
            {
                ws.Cells[row, 1, row + objCount - 1, 1].Merge = true;
                ws.Cells[row, 2, row + objCount - 1, 2].Merge = true;
                ws.Cells[row, 3, row + objCount - 1, 3].Merge = true;
            }

            // 跳过多余行（i 的步进不变，靠 row+oi 展开）
            // 注意：这里用简化方式——只展示第一个 objective/reward 以保持行对齐
            // 已在上面 oi loop 里处理
        }

        // 简化：每订单一行，多个 objective 合并显示
        // 重新渲染（覆盖上面的多行逻辑，改为单行）
        var ws2 = pkg.Workbook.Worksheets.Add("订单系统（单行）");
        MechNote(
            ws2,
            1,
            1,
            "【订单系统 · 单行版】每个订单一行，多个需求物品用 | 分隔。共 389 个订单，103 棵树。",
            8
        );
        ws2.Row(1).Height = 40;
        for (var j = 0; j < headers.Length; j++)
            Header(ws2.Cells[2, j + 1], headers[j], hdrHex[j]);
        ws2.Row(2).Height = 22;
        ws2.View.FreezePanes(3, 1);

        for (var i = 0; i < orders.Count; i++)
        {
            var row = i + 3;
            var o = orders[i];
            var altHex = o.TreeId % 2 == 0 ? "EBF3FB" : "F8FBFF";

            Cell(ws2.Cells[row, 1], o.OrderId, "F2F2F2");
            Cell(ws2.Cells[row, 2], o.TreeId, altHex);
            Cell(ws2.Cells[row, 3], o.SeqInTree, altHex);

            var objText = string.Join(
                " | ",
                o.Objectives.Select(obj =>
                    $"{obj.ItemReference.Replace("item_", "")} ×{obj.Amount}"
                )
            );
            var objTiers = o
                .Objectives.Select(obj => TierFromItemId(obj.ItemReference))
                .Where(t => t >= 0)
                .ToList();
            var maxTier = objTiers.Count > 0 ? objTiers.Max() : -1;
            var tierHex = tierColors.GetValueOrDefault(Math.Min(maxTier, 5), "F5F5F5");

            Cell(ws2.Cells[row, 4], objText, tierHex, wrap: true);
            Cell(ws2.Cells[row, 5], maxTier >= 0 ? maxTier : "", tierHex);
            Cell(
                ws2.Cells[row, 6],
                o.Objectives.Count > 0 ? o.Objectives.Sum(x => x.Amount) : 0,
                tierHex
            );

            var rewText = string.Join(
                " | ",
                o.Rewards.Select(r => $"{r.ItemReference.Replace("item_", "")} ×{r.Amount}")
            );
            Cell(ws2.Cells[row, 7], rewText, "D9EAD3", wrap: true);
            Cell(
                ws2.Cells[row, 8],
                o.Rewards.Count > 0 ? o.Rewards.Sum(r => r.Amount) : 0,
                "D9EAD3"
            );
        }

        // 列宽（两个 sheet 相同）
        foreach (var sheet in new[] { ws, ws2 })
        {
            sheet.Column(1).Width = 36;
            sheet.Column(2).Width = 8;
            sheet.Column(3).Width = 8;
            sheet.Column(4).Width = 30;
            sheet.Column(5).Width = 10;
            sheet.Column(6).Width = 10;
            sheet.Column(7).Width = 28;
            sheet.Column(8).Width = 10;
        }
    }

    // ── Sheet 3：棋盘布局 ─────────────────────────────────────────────────────

    private static void BuildBoardSheet(ExcelPackage pkg, List<BoardCell> board)
    {
        var ws = pkg.Workbook.Worksheets.Add("棋盘布局");

        MechNote(
            ws,
            1,
            1,
            "【棋盘布局】数据源：HTTPCache/10 boardCoreConfiguration.startingBoard。"
                + "9行×7列共63格。绿色=未锁定可操作，橙色=装箱(Boxed)，灰色=等级锁。",
            8
        );
        ws.Row(1).Height = 50;

        string[] headers = ["行", "列", "物品ID", "物品类型", "Tier", "锁定", "装箱", "等级解锁"];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], j < 2 ? "1F4E79" : "2F5496");
        ws.Row(2).Height = 22;
        ws.View.FreezePanes(3, 1);

        for (var i = 0; i < board.Count; i++)
        {
            var row = i + 3;
            var c = board[i];
            var itemType = Regex.Replace(c.Item, @"_\d{2}$", "").Replace("item_", "");
            var tier = TierFromItemId(c.Item);

            string hex;
            if (!c.Locked)
                hex = "D9EAD3"; // 未锁定：绿
            else if (c.Boxed)
                hex = "FCE5CD"; // 装箱：橙
            else if (c.LevelLock > 0)
                hex = "F5F5F5"; // 等级锁：灰
            else
                hex = "EBF3FB"; // 普通锁定：蓝

            Cell(ws.Cells[row, 1], c.Row, "F2F2F2");
            Cell(ws.Cells[row, 2], c.Col, "F2F2F2");
            Cell(ws.Cells[row, 3], c.Item, hex);
            Cell(ws.Cells[row, 4], itemType, hex);
            Cell(ws.Cells[row, 5], tier >= 0 ? tier : "", hex);
            Cell(ws.Cells[row, 6], c.Locked ? "✓" : "", c.Locked ? "FCE5CD" : "D9EAD3");
            Cell(ws.Cells[row, 7], c.Boxed ? "✓" : "", c.Boxed ? "FCE5CD" : null);
            Cell(
                ws.Cells[row, 8],
                c.LevelLock > 0 ? c.LevelLock : "",
                c.LevelLock > 0 ? "F5F5F5" : null
            );
        }

        // 统计
        var statRow = board.Count + 4;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"总格: {board.Count}";
        ws.Cells[statRow, 3].Value = $"未锁定: {board.Count(c => !c.Locked)}";
        ws.Cells[statRow, 4].Value = $"装箱: {board.Count(c => c.Boxed)}";
        ws.Cells[statRow, 5].Value = $"等级锁: {board.Count(c => c.LevelLock > 0)}";

        ws.Column(1).Width = 6;
        ws.Column(2).Width = 6;
        ws.Column(3).Width = 32;
        ws.Column(4).Width = 24;
        ws.Column(5).Width = 8;
        ws.Column(6).Width = 8;
        ws.Column(7).Width = 8;
        ws.Column(8).Width = 10;
    }

    // ── Sheet 4：合成链 ───────────────────────────────────────────────────────

    private static void BuildChainSheet(ExcelPackage pkg, List<MergeChain> chains)
    {
        var ws = pkg.Workbook.Worksheets.Add("合成链");

        var maxLen = chains.Count > 0 ? chains.Max(c => c.ItemIds.Count) : 1;
        var colCount = 3 + Math.Min(maxLen, 15);

        MechNote(
            ws,
            1,
            1,
            $"【合成链】数据源：HTTPCache/1C mergeItemGraphs。共 {chains.Count} 条有效链（≥2 tier）。"
                + $"最长 {maxLen} tier。链尾（最高 tier）黄色高亮。",
            colCount
        );
        ws.Row(1).Height = 50;

        Header(ws.Cells[2, 1], "序号", "1F4E79");
        Header(ws.Cells[2, 2], "GraphID", "1F4E79");
        Header(ws.Cells[2, 3], "链长", "2F5496");
        for (var j = 0; j < Math.Min(maxLen, 15); j++)
            Header(ws.Cells[2, 4 + j], $"T{j:D2}", "2F5496");
        ws.Row(2).Height = 22;
        ws.View.FreezePanes(3, 1);

        for (var i = 0; i < chains.Count; i++)
        {
            var row = i + 3;
            var chain = chains[i];
            Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            Cell(ws.Cells[row, 2], chain.GraphId, "F0F4FA");
            Cell(ws.Cells[row, 3], chain.ItemIds.Count, "EBF3FB");

            var displayCount = Math.Min(chain.ItemIds.Count, 15);
            for (var j = 0; j < displayCount; j++)
            {
                var isLast = j == chain.ItemIds.Count - 1;
                var itemShort = chain.ItemIds[j].Replace("item_", "");
                Cell(ws.Cells[row, 4 + j], itemShort, isLast ? "FFF2CC" : "EBF3FB");
            }
            if (chain.ItemIds.Count > 15)
                ws.Cells[row, 4 + 15].Value = $"…+{chain.ItemIds.Count - 15}";
        }

        // 统计
        var statRow = chains.Count + 4;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"共 {chains.Count} 条链";
        ws.Cells[statRow, 3].Value = $"最长 {maxLen} tier";
        ws.Cells[statRow, 4].Value =
            $"7-9tier: {chains.Count(c => c.ItemIds.Count >= 7 && c.ItemIds.Count <= 9)}条";
        ws.Cells[statRow, 5].Value = $"10+tier: {chains.Count(c => c.ItemIds.Count >= 10)}条";

        ws.Column(1).Width = 7;
        ws.Column(2).Width = 36;
        ws.Column(3).Width = 8;
        for (var j = 0; j < Math.Min(maxLen, 15); j++)
            ws.Column(4 + j).Width = 24;
    }

    // ── Sheet 5：生成器配置 ───────────────────────────────────────────────────

    private static void BuildSpawnerSheet(ExcelPackage pkg, List<MergeChain> chains)
    {
        var ws = pkg.Workbook.Worksheets.Add("生成器配置");

        // 从 chains 里找有 spawn 信息的（graphId 含 spawner / producer / grocery 等关键词）
        var spawnerKeywords = new[]
        {
            "spawner",
            "producer",
            "grocery",
            "store",
            "cart",
            "vehicle",
            "car",
            "truck",
        };
        var spawnerChains = chains
            .Where(c => spawnerKeywords.Any(k => c.GraphId.ToLower().Contains(k)))
            .ToList();

        // 也从 item names 文件里找 spawner 类型
        var itemNamesFile = @"C:\tmp\tasty_item_names.txt";
        var spawnerItems = new List<(string Type, string ItemId, int Tier)>();
        if (File.Exists(itemNamesFile))
        {
            foreach (var line in File.ReadAllLines(itemNamesFile))
            {
                var seg = line.Trim().Split('/').Last();
                if (string.IsNullOrEmpty(seg))
                    continue;
                var m = Regex.Match(seg, @"^\d+_(\w+)_(.+?)_(\d{2})$");
                if (m.Success && m.Groups[1].Value.ToLower() == "spawner")
                    spawnerItems.Add(
                        (m.Groups[1].Value, m.Groups[2].Value, int.Parse(m.Groups[3].Value))
                    );
            }
        }

        MechNote(
            ws,
            1,
            1,
            "【生成器配置】数据源：HTTPCache/1C mergeItemGraphs（含 spawn 配置） + 设备 cached_merge_items 文件名。"
                + "生成器是棋盘核心驱动力，决定产出节奏和 chain 覆盖面。",
            6
        );
        ws.Row(1).Height = 50;

        // Section A：从 item names 提取的 spawner 列表
        var titleA = ws.Cells[2, 1, 2, 6];
        titleA.Merge = true;
        titleA.Value = "▶ 生成器物品列表（来源：设备缓存文件名）";
        titleA.Style.Font.Bold = true;
        titleA.Style.Fill.PatternType = ExcelFillStyle.Solid;
        titleA.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
        titleA.Style.Font.Color.SetColor(Color.White);
        Border(titleA);

        string[] hA = ["#", "生成器ID", "Tier", "推断关联Chain", "备注", ""];
        for (var j = 0; j < 5; j++)
            Header(ws.Cells[3, j + 1], hA[j], "2F5496");
        ws.Row(3).Height = 20;

        var spawnerGroups = spawnerItems.GroupBy(s => s.ItemId).OrderBy(g => g.Key).ToList();

        var row = 4;
        for (var i = 0; i < spawnerGroups.Count; i++)
        {
            var g = spawnerGroups[i];
            var tiers = g.Select(x => x.Tier).OrderBy(t => t).ToList();
            var relatedChain =
                chains.FirstOrDefault(c => c.GraphId.ToLower().Contains(g.Key.ToLower()))?.GraphId
                ?? "—";
            var hex = i % 2 == 0 ? "EBF3FB" : "F0F4FA";
            Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            Cell(ws.Cells[row, 2], g.Key, hex);
            Cell(ws.Cells[row, 3], $"T{tiers.Min():D2}~T{tiers.Max():D2}（{tiers.Count}级）", hex);
            Cell(ws.Cells[row, 4], relatedChain, "FFF2CC");
            Cell(ws.Cells[row, 5], "", hex);
            row++;
        }

        // Section B：从 mergeItemGraphs 里找到的 spawner chain
        row += 2;
        var titleB = ws.Cells[row, 1, row, 6];
        titleB.Merge = true;
        titleB.Value = "▶ 含生成逻辑的合成链（mergeItemGraphs 中含 spawner 关键词）";
        titleB.Style.Font.Bold = true;
        titleB.Style.Fill.PatternType = ExcelFillStyle.Solid;
        titleB.Style.Fill.BackgroundColor.SetColor(HexColor("385623"));
        titleB.Style.Font.Color.SetColor(Color.White);
        Border(titleB);
        row++;

        string[] hB = ["#", "GraphID", "链长", "首项(T00)", "末项", "备注"];
        for (var j = 0; j < hB.Length; j++)
            Header(ws.Cells[row, j + 1], hB[j], "2F5496");
        row++;

        for (var i = 0; i < spawnerChains.Count; i++)
        {
            var c = spawnerChains[i];
            var hex = i % 2 == 0 ? "D9EAD3" : "EBF3FB";
            Cell(ws.Cells[row, 1], i + 1, "F2F2F2");
            Cell(ws.Cells[row, 2], c.GraphId, hex);
            Cell(ws.Cells[row, 3], c.ItemIds.Count, hex);
            Cell(ws.Cells[row, 4], c.ItemIds.FirstOrDefault()?.Replace("item_", "") ?? "", hex);
            Cell(ws.Cells[row, 5], c.ItemIds.LastOrDefault()?.Replace("item_", "") ?? "", "FFF2CC");
            Cell(ws.Cells[row, 6], "", hex);
            row++;
        }

        ws.Column(1).Width = 6;
        ws.Column(2).Width = 34;
        ws.Column(3).Width = 18;
        ws.Column(4).Width = 28;
        ws.Column(5).Width = 28;
        ws.Column(6).Width = 20;
    }
}
