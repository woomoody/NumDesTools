using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// Travel Town 竞品核心循环分析 xlsx 生成器。
/// 数据源：C:\tmp\traveltown_lua\（由 ADB HTTPCache 提取 + Unity AssetBundle 解析）
/// 输出：竞品-TravelTown核心循环分析.xlsx → Documents\workspace\
/// </summary>
public static class TravelTownWriter
{
    private const string DataDir = @"C:\tmp\traveltown_lua";
    private const string OutFileName = "竞品-TravelTown核心循环分析.xlsx";

    // ── 数据模型 ──────────────────────────────────────────────────────────────
    private record ChainItem(
        string UniqueId,
        string GraphId,
        int LevelIndex,
        double AvgDifficulty,
        int SellAmount,
        string SellResource,
        bool RequestedByOrders,
        string Scope,
        // produce fields (for generator items)
        bool IsGenerator,
        int? CycleDelay,
        int? Capacity,
        int? ItemsPerSubCycle,
        int? EnergyCost,
        string? SpawnTarget,
        int? SpawnWeight,
        // spawner difficulty per generator level (for product items)
        List<(string SpawnerId, double Difficulty)> SpawnerDiffs
    );

    private record ChainInfo(
        string UniqueId,
        string GraphName,
        List<ChainItem> Items,
        bool HasProducers,
        bool IsGeneratorChain
    );

    private record OrderStep(
        string OrderId,
        string TreeId,
        int TreeIndex,
        List<(string ItemRef, int Amount)> Objectives,
        List<(string ItemRef, int Amount)> Rewards,
        List<string> LockedByIds,
        string TaskType
    );

    // ── 数据加载 ──────────────────────────────────────────────────────────────
    private static List<ChainInfo>? _chainsCache;

    private static List<ChainInfo> LoadChains()
    {
        if (_chainsCache is not null)
            return _chainsCache;

        var path = Path.Combine(DataDir, "item_merge_graphs_full.json");
        if (!File.Exists(path))
            return (_chainsCache = []);

        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        var root = doc.RootElement;

        // First pass: build item map (all items across all chains)
        var allItems = new Dictionary<string, (JsonElement El, string ChainId)>();
        foreach (var chain in root.EnumerateArray())
        {
            var chainId = chain.GetProperty("uniqueId").GetString() ?? "";
            foreach (var item in chain.GetProperty("items").EnumerateArray())
            {
                var uid = item.GetProperty("uniqueId").GetString() ?? "";
                if (!string.IsNullOrEmpty(uid))
                    allItems[uid] = (item, chainId);
            }
        }

        // Collect all spawner IDs (items referenced as spawner in difficulty.difficulties)
        var spawnerIds = new HashSet<string>();
        foreach (var (item, _) in allItems.Values)
        {
            if (
                item.TryGetProperty("difficulty", out var diff)
                && diff.TryGetProperty("difficulties", out var diffs)
            )
            {
                foreach (var d in diffs.EnumerateArray())
                {
                    if (d.TryGetProperty("spawner", out var sv))
                        spawnerIds.Add(sv.GetString() ?? "");
                }
            }
        }

        // Second pass: build ChainInfo
        var result = new List<ChainInfo>();
        foreach (var chain in root.EnumerateArray())
        {
            var chainId = chain.GetProperty("uniqueId").GetString() ?? "";
            var graphName = chain.TryGetProperty("graphName", out var gn)
                ? (gn.GetString() ?? "")
                : "";

            var rawItems = chain.GetProperty("items").EnumerateArray().ToList();
            var items = new List<ChainItem>();
            var isGeneratorChain = false;

            for (var idx = 0; idx < rawItems.Count; idx++)
            {
                var it = rawItems[idx];
                var uid = it.GetProperty("uniqueId").GetString() ?? "";

                // Difficulty
                double avgDiff = 0;
                var spawnerDiffs = new List<(string, double)>();
                if (it.TryGetProperty("difficulty", out var diffEl))
                {
                    if (diffEl.TryGetProperty("averageDifficulty", out var ad))
                        avgDiff = ad.GetDouble();
                    if (diffEl.TryGetProperty("difficulties", out var ds))
                    {
                        foreach (var d in ds.EnumerateArray())
                        {
                            var sp = d.TryGetProperty("spawner", out var sv)
                                ? (sv.GetString() ?? "")
                                : "";
                            var dv = d.TryGetProperty("difficulty", out var dv2)
                                ? dv2.GetDouble()
                                : 0;
                            if (!string.IsNullOrEmpty(sp))
                                spawnerDiffs.Add((sp, dv));
                        }
                    }
                }

                // Sell
                var sellAmt = 0;
                var sellRes = "";
                if (it.TryGetProperty("sell", out var sell))
                {
                    if (sell.TryGetProperty("amount", out var sa))
                        sellAmt = sa.GetInt32();
                    if (sell.TryGetProperty("resource", out var sr))
                        sellRes = sr.GetString() ?? "";
                }

                // Produce (generator)
                var isGen = false;
                int? cycleDelay = null,
                    capacity = null,
                    itemsPerSub = null,
                    energyCost = null;
                string? spawnTarget = null;
                int? spawnWeight = null;
                if (
                    it.TryGetProperty("produce", out var prod)
                    && prod.TryGetProperty("enabled", out var pe)
                    && pe.GetBoolean()
                )
                {
                    isGen = true;
                    isGeneratorChain = true;
                    if (prod.TryGetProperty("cycleDelay", out var cd))
                        cycleDelay = cd.GetInt32();
                    if (prod.TryGetProperty("capacity", out var cap))
                        capacity = cap.GetInt32();
                    if (prod.TryGetProperty("itemsPerSubCycle", out var isc))
                        itemsPerSub = isc.GetInt32();
                    if (prod.TryGetProperty("items", out var pi) && pi.GetArrayLength() > 0)
                    {
                        var first = pi[0];
                        if (first.TryGetProperty("itemReference", out var ir))
                            spawnTarget = ir.GetString();
                        if (first.TryGetProperty("weight", out var wt))
                            spawnWeight = wt.GetInt32();
                    }
                }

                // Interaction (energy cost)
                if (it.TryGetProperty("interaction", out var inter))
                {
                    if (inter.TryGetProperty("resourceAmountToConsume", out var ec))
                        energyCost = ec.GetInt32();
                }

                var scope = it.TryGetProperty("scope", out var sc) ? (sc.GetString() ?? "") : "";
                var reqByOrders =
                    it.TryGetProperty("requestedByOrders", out var ro) && ro.GetBoolean();

                items.Add(
                    new ChainItem(
                        uid,
                        chainId,
                        idx,
                        avgDiff,
                        sellAmt,
                        sellRes,
                        reqByOrders,
                        scope,
                        isGen,
                        cycleDelay,
                        capacity,
                        itemsPerSub,
                        energyCost,
                        spawnTarget,
                        spawnWeight,
                        spawnerDiffs
                    )
                );
            }

            var hasProducers =
                chain.TryGetProperty("producersStacks", out var ps)
                && ps.ValueKind != JsonValueKind.Null
                && ps.ValueKind == JsonValueKind.Object;

            result.Add(new ChainInfo(chainId, graphName, items, hasProducers, isGeneratorChain));
        }

        return (_chainsCache = result);
    }

    private static List<OrderStep>? _ordersCache;

    private static List<OrderStep> LoadOrders()
    {
        if (_ordersCache is not null)
            return _ordersCache;
        var path = Path.Combine(DataDir, "orders_intro.json");
        if (!File.Exists(path))
            return (_ordersCache = []);

        using var doc = JsonDocument.Parse(File.ReadAllText(path));
        var result = new List<OrderStep>();

        foreach (var tree in doc.RootElement.EnumerateArray())
        {
            var treeId = tree.GetProperty("uniqueTreeId").GetInt64().ToString();
            var treeIdx = (int)(tree.GetProperty("uniqueTreeId").GetInt64() % 100000);

            foreach (var order in tree.GetProperty("orders").EnumerateArray())
            {
                var orderId = order.TryGetProperty("orderId", out var oid)
                    ? (oid.GetString() ?? "")
                    : "";
                var taskType = order.TryGetProperty("taskType", out var tt)
                    ? (tt.GetString() ?? "")
                    : "";

                var objectives = new List<(string, int)>();
                if (order.TryGetProperty("objectives", out var objs))
                {
                    foreach (var obj in objs.EnumerateArray())
                    {
                        var ir = obj.TryGetProperty("itemReference", out var irv)
                            ? (irv.GetString() ?? "")
                            : "";
                        var amt = obj.TryGetProperty("amount", out var av) ? av.GetInt32() : 1;
                        objectives.Add((ir, amt));
                    }
                }

                var rewards = new List<(string, int)>();
                if (order.TryGetProperty("itemRewards", out var rwds))
                {
                    foreach (var rwd in rwds.EnumerateArray())
                    {
                        var ir = rwd.TryGetProperty("itemReference", out var irv)
                            ? (irv.GetString() ?? "")
                            : "";
                        var amt = rwd.TryGetProperty("amount", out var av) ? av.GetInt32() : 1;
                        rewards.Add((ir, amt));
                    }
                }

                var lockedBy = new List<string>();
                if (order.TryGetProperty("lockedByIds", out var lbIds))
                {
                    foreach (var lb in lbIds.EnumerateArray())
                        lockedBy.Add(lb.GetString() ?? "");
                }

                result.Add(
                    new OrderStep(orderId, treeId, treeIdx, objectives, rewards, lockedBy, taskType)
                );
            }
        }

        return (_ordersCache = result);
    }

    // Short display name from uniqueId: item_laundry-accessory_03 → "laundry-accessory Lv.3"
    private static string ShortName(string uid)
    {
        var m = Regex.Match(uid, @"^item_(.+?)_(\d+)$");
        if (m.Success)
            return $"{m.Groups[1].Value} Lv.{int.Parse(m.Groups[2].Value)}";
        return uid.Replace("item_", "").Replace("item-graph_", "");
    }

    private static string ChainShortName(string chainId) => chainId.Replace("item-graph_", "");

    // ── 公开入口 ──────────────────────────────────────────────────────────────
    public static void Run(string? outputDir = null)
    {
        var dir = outputDir ?? OutputPaths.Reports;
        var outPath = Path.Combine(dir, OutFileName);

        using var pkg = new ExcelPackage();

        BuildGeneratorSheet(pkg);
        BuildProductChainSheet(pkg);
        BuildOrderSheet(pkg);
        BuildBuildingMappingSheet(pkg);
        BuildPaymentSheet(pkg);
        BuildSummarySheet(pkg);

        pkg.SaveAs(new FileInfo(outPath));
        Console.WriteLine($"[TravelTown] 已生成：{outPath}");
        OutputPaths.GitCommit($"[TravelTown] 更新竞品分析报告 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── Sheet 1：生成器链（Generators）────────────────────────────────────────
    private static void BuildGeneratorSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("生成器链");
        ws.View.FreezePanes(3, 1);

        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【Travel Town 生成器机制】生成器是可升级的「机器」，分 7~9 个等级（Lv.0~Lv.6/8）。"
                + "每次点击消耗 1 Energy，固定产出1种初始元素（weight=100，无随机池）。"
                + "生成器本身也是合并对象：低级机器 × 2 → 合并为下一级机器，产出效率随等级提升（容量/冷却/体积）。"
                + "cycleDelay=0 表示点击即产，>0 表示自动填充冷却时间（秒）。"
                + "capacity=格子容量上限；itemsPerSubCycle=每次自动填充数量。",
            16
        );

        string[] headers =
        [
            "链名",
            "机器等级",
            "机器ID",
            "点击CD(s)",
            "容量(格)",
            "每次产出数",
            "能量消耗",
            "固定产出元素",
            "产出权重",
            "类型",
        ];
        string[] hdrHex =
        [
            "1F4E79",
            "595959",
            "2F5496",
            "2F5496",
            "2F5496",
            "2F5496",
            "595959",
            "1F4E79",
            "595959",
            "595959",
        ];
        for (var j = 0; j < headers.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var chains = LoadChains();
        var genChains = chains.Where(c => c.IsGeneratorChain).OrderBy(c => c.UniqueId).ToList();

        var chainColors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var curRow = 3;
        var ci = 0;

        foreach (var chain in genChains)
        {
            var baseHex = chainColors[ci % chainColors.Length];
            ci++;

            // Chain title row
            var titleCell = ws.Cells[curRow, 1, curRow, headers.Length];
            titleCell.Merge = true;
            titleCell.Value = $"▶ {ChainShortName(chain.UniqueId)}  共{chain.Items.Count}级";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleCell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("1F4E79"));
            titleCell.Style.Font.Color.SetColor(Color.White);
            XlsxStyleHelper.Border(titleCell);
            curRow++;

            foreach (var item in chain.Items)
            {
                if (!item.IsGenerator)
                    continue;

                var isFirst = item.LevelIndex == 0;
                var isLast = item.LevelIndex == chain.Items.Count - 1;
                var rowHex = isFirst
                    ? "D9EAD3"
                    : isLast
                        ? "FADADD"
                        : baseHex;

                var lvLabel = isFirst
                    ? "Lv.0(基础)"
                    : isLast
                        ? $"Lv.{item.LevelIndex}(最高)"
                        : $"Lv.{item.LevelIndex}";
                var typeLabel = item.CycleDelay is > 0 ? "自动填充" : "点击即产";
                var typeHex = item.CycleDelay is > 0 ? "EBF3FB" : "D9EAD3";

                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 1],
                    isFirst ? ChainShortName(chain.UniqueId) : "",
                    isFirst ? "D9EAD3" : "F5F5F5"
                );
                XlsxStyleHelper.Cell(ws.Cells[curRow, 2], lvLabel, "F5F5F5");
                XlsxStyleHelper.Cell(ws.Cells[curRow, 3], ShortName(item.UniqueId), rowHex);
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 4],
                    item.CycleDelay is > 0 ? (object)item.CycleDelay : "即时",
                    item.CycleDelay is > 0 ? "FFF2CC" : "D9EAD3"
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 5],
                    item.Capacity.HasValue ? (object)item.Capacity.Value : "—",
                    item.Capacity.HasValue ? "EBF3FB" : null
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 6],
                    item.ItemsPerSubCycle.HasValue ? (object)item.ItemsPerSubCycle.Value : "—",
                    null
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 7],
                    item.EnergyCost.HasValue ? (object)item.EnergyCost.Value : "—",
                    item.EnergyCost is 1 ? "D9EAD3" : null
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 8],
                    item.SpawnTarget is not null ? ShortName(item.SpawnTarget) : "—",
                    item.SpawnTarget is not null ? "F0F8FF" : null
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 9],
                    item.SpawnWeight.HasValue ? (object)item.SpawnWeight.Value : "—",
                    null
                );
                XlsxStyleHelper.Cell(ws.Cells[curRow, 10], typeLabel, typeHex);
                curRow++;
            }
            curRow++;
        }

        ws.Column(1).Width = 22;
        ws.Column(2).Width = 14;
        ws.Column(3).Width = 26;
        ws.Column(4).Width = 12;
        ws.Column(5).Width = 10;
        ws.Column(6).Width = 12;
        ws.Column(7).Width = 10;
        ws.Column(8).Width = 28;
        ws.Column(9).Width = 10;
        ws.Column(10).Width = 12;
        ws.Row(1).Height = 56;
        ws.Cells[1, 1].Style.WrapText = true;

        // Stats
        ws.Cells[curRow, 1].Value = $"合计 {genChains.Count} 条生成器链";
        ws.Cells[curRow, 1].Style.Font.Bold = true;
    }

    // ── Sheet 2：产品链（Product Chains）──────────────────────────────────────
    // 核心合并链：元素等级 × 难度系数（spawner难度）
    private static void BuildProductChainSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("产品合并链");
        ws.View.FreezePanes(3, 1);

        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【产品链与难度系数】每条产品链包含多个等级的元素（Lv.0~Lv.N），低等级元素 2个合并→高1级。"
                + "difficulty.averageDifficulty = 该等级元素的「平均产出难度」（需点击多少次生成器才能合出来）。"
                + "difficulty.difficulties[] = 各生成器等级对应的难度：数值越低越容易产出，越高越稀有。"
                + "requestedByOrders=true 表示该等级元素可能出现在订单需求中。"
                + "sell.resource=Smiley（好评点数），sell.amount=出售获得数量。"
                + "Travel Town 的元素难度是「硬编码到配置」的，不像 Gossip Harbor 有运行时动态权重。",
            16
        );

        // Headers — columns: 链名 | 等级 | 元素ID | 平均难度 | 出售 | 订单需求 | 生成器0难度 | 生成器1 | ...
        // Find max spawner count in core chains
        var chains = LoadChains();
        var coreChains = chains
            .Where(c =>
                !c.IsGeneratorChain && c.Items.Any(it => it.Scope == "core" && it.RequestedByOrders)
            )
            .OrderBy(c => c.UniqueId)
            .ToList();

        // Collect all spawner IDs in a consistent order
        var allSpawnerIds = new List<string>();
        var seenSpawners = new HashSet<string>();
        foreach (var chain in coreChains)
        {
            foreach (var item in chain.Items)
            {
                foreach (var (sid, _) in item.SpawnerDiffs)
                {
                    if (seenSpawners.Add(sid))
                        allSpawnerIds.Add(sid);
                }
            }
        }

        // Cap at 6 spawner columns for display
        var displaySpawners = allSpawnerIds.Take(6).ToList();
        var baseColCount = 6;
        var totalCols = baseColCount + displaySpawners.Count;

        string[] baseHeaders = ["链名", "等级", "元素ID(短)", "平均难度", "出售(笑脸)", "订单需求"];
        string[] baseHex = ["1F4E79", "595959", "2F5496", "2F5496", "595959", "1F4E79"];
        for (var j = 0; j < baseHeaders.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], baseHeaders[j], baseHex[j]);
        for (var j = 0; j < displaySpawners.Count; j++)
        {
            var spawnShort = ShortName(displaySpawners[j]);
            XlsxStyleHelper.Header(ws.Cells[2, baseColCount + j + 1], $"难度@\n{spawnShort}", "2F5496");
        }

        var chainColors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var curRow = 3;
        var ci = 0;

        foreach (var chain in coreChains)
        {
            var baseHexChain = chainColors[ci % chainColors.Length];
            ci++;

            var titleCell = ws.Cells[curRow, 1, curRow, totalCols];
            titleCell.Merge = true;
            titleCell.Value =
                $"▶ {ChainShortName(chain.UniqueId)}  共{chain.Items.Count}级  订单元素:{chain.Items.Count(i => i.RequestedByOrders)}种";
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleCell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("1F4E79"));
            titleCell.Style.Font.Color.SetColor(Color.White);
            XlsxStyleHelper.Border(titleCell);
            curRow++;

            foreach (var item in chain.Items)
            {
                var isFirst = item.LevelIndex == 0;
                var isLast = item.LevelIndex == chain.Items.Count - 1;
                var rowHex = isFirst
                    ? "EBF3FB"
                    : isLast
                        ? "FADADD"
                        : baseHexChain;
                var diffHex = item.AvgDifficulty switch
                {
                    < 0.05 => "D9EAD3", // easy
                    < 0.15 => "FFF2CC",
                    < 0.3 => "FCE5CD",
                    _ => "FADADD", // hard
                };

                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 1],
                    isFirst ? ChainShortName(chain.UniqueId) : "",
                    isFirst ? "D9EAD3" : "F5F5F5"
                );
                XlsxStyleHelper.Cell(ws.Cells[curRow, 2], $"Lv.{item.LevelIndex}", "F5F5F5");
                XlsxStyleHelper.Cell(ws.Cells[curRow, 3], ShortName(item.UniqueId), rowHex);
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 4],
                    item.AvgDifficulty > 0 ? (object)item.AvgDifficulty : "—",
                    diffHex
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 5],
                    item.SellAmount > 0 ? (object)item.SellAmount : "—",
                    item.SellAmount > 0 ? "FFF2CC" : null
                );
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 6],
                    item.RequestedByOrders ? "✓" : "",
                    item.RequestedByOrders ? "D9EAD3" : null
                );

                var diffMap = item.SpawnerDiffs.ToDictionary(d => d.SpawnerId, d => d.Difficulty);
                for (var j = 0; j < displaySpawners.Count; j++)
                {
                    var sid = displaySpawners[j];
                    if (diffMap.TryGetValue(sid, out var dv))
                    {
                        var dvHex = dv switch
                        {
                            <= 0.1 => "D9EAD3",
                            <= 0.5 => "FFF2CC",
                            <= 1.0 => "FCE5CD",
                            _ => "FADADD",
                        };
                        XlsxStyleHelper.Cell(ws.Cells[curRow, baseColCount + j + 1], dv, dvHex);
                    }
                    else
                    {
                        XlsxStyleHelper.Cell(ws.Cells[curRow, baseColCount + j + 1], "—", "F0F0F0");
                    }
                }
                curRow++;
            }
            curRow++;
        }

        ws.Column(1).Width = 22;
        ws.Column(2).Width = 8;
        ws.Column(3).Width = 26;
        ws.Column(4).Width = 12;
        ws.Column(5).Width = 10;
        ws.Column(6).Width = 10;
        for (var j = 0; j < displaySpawners.Count; j++)
            ws.Column(baseColCount + j + 1).Width = 16;

        ws.Row(1).Height = 56;
        ws.Row(2).Height = 36;
        ws.Cells[1, 1].Style.WrapText = true;

        ws.Cells[curRow, 1].Value = $"合计 {coreChains.Count} 条核心产品链";
        ws.Cells[curRow, 1].Style.Font.Bold = true;
    }

    // ── Sheet 3：订单（Order Trees）───────────────────────────────────────────
    private static void BuildOrderSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("订单任务树");
        ws.View.FreezePanes(3, 1);

        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【Travel Town 订单机制】订单以「任务树(Tree)」为单位，每棵树包含若干串联步骤（必须按顺序完成）。"
                + "数据来源：orders_intro.json 包含全部 103 棵树（含早期 ID 1~75 和后期 ID >1M），"
                + "「introOrders」只是系统内部命名，并非「教程专用」——这就是 TT 完整的订单系统，没有单独的主线循环。"
                + "每步骤：需提交 objectiveItems 指定的元素，获得 rewardItems 奖励（宝箱/钻石/工具箱等）。"
                + "lockedByIds = 前置订单ID（串联依赖，空=起始步骤）。"
                + "与 Gossip Harbor 对比：TT 是线性剧情推进（A完成才解锁B），Harbor 是动态权重随机刷新。",
            16
        );

        string[] headers = ["树ID", "步骤#", "订单ID(短)", "需求元素", "需求数", "奖励", "前置", "任务类型"];
        string[] hdrHex =
        [
            "1F4E79",
            "595959",
            "2F5496",
            "1F4E79",
            "595959",
            "2F5496",
            "595959",
            "595959"
        ];
        for (var j = 0; j < headers.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var orders = LoadOrders();
        var byTree = orders.GroupBy(o => o.TreeId).OrderBy(g => long.Parse(g.Key)).ToList();

        var treeColors = new[] { "EBF3FB", "FFF2CC", "D9EAD3", "FCE5CD", "E8DAEF", "F3F3F3" };
        var curRow = 3;
        var ti = 0;

        foreach (var treeGrp in byTree)
        {
            var treeHex = treeColors[ti % treeColors.Length];
            ti++;

            var treeSteps = treeGrp
                .OrderBy(o => o.LockedByIds.Count)
                .ThenBy(o => o.OrderId)
                .ToList();
            var stepNum = 0;

            foreach (var step in treeSteps)
            {
                stepNum++;
                var isFirst = stepNum == 1;

                var objStr = string.Join(", ", step.Objectives.Select(o => ShortName(o.ItemRef)));
                var objAmt = step.Objectives.Count > 0 ? (object)step.Objectives[0].Amount : "—";
                var rewardStr =
                    step.Rewards.Count > 0
                        ? string.Join(
                            ", ",
                            step.Rewards.Select(r => $"{ShortName(r.ItemRef)}×{r.Amount}")
                        )
                        : "—";
                var lockedStr =
                    step.LockedByIds.Count > 0
                        ? string.Join(
                            ", ",
                            step.LockedByIds.Select(id => id.Split('_').LastOrDefault() ?? id)
                        )
                        : "（起始）";

                // Short order ID: remove tree prefix
                var shortId = Regex.Replace(step.OrderId, @"^intro_", "");

                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 1],
                    isFirst ? $"Tree {step.TreeId}" : "",
                    isFirst ? "1F4E79" : null
                );
                if (isFirst)
                    ws.Cells[curRow, 1].Style.Font.Color.SetColor(Color.White);
                XlsxStyleHelper.Cell(ws.Cells[curRow, 2], stepNum, "F5F5F5");
                XlsxStyleHelper.Cell(ws.Cells[curRow, 3], shortId, treeHex);
                XlsxStyleHelper.Cell(ws.Cells[curRow, 4], objStr, treeHex, wrap: true);
                XlsxStyleHelper.Cell(ws.Cells[curRow, 5], objAmt, "F9F9F9");
                XlsxStyleHelper.Cell(
                    ws.Cells[curRow, 6],
                    rewardStr,
                    step.Rewards.Any(r => r.ItemRef.Contains("diamond")) ? "FFF2CC" : treeHex,
                    wrap: true
                );
                XlsxStyleHelper.Cell(ws.Cells[curRow, 7], lockedStr, lockedStr == "（起始）" ? "D9EAD3" : "F9F9F9");
                XlsxStyleHelper.Cell(ws.Cells[curRow, 8], step.TaskType, "F5F5F5");
                ws.Row(curRow).Height = 20;
                curRow++;
            }
            curRow++;
        }

        ws.Column(1).Width = 10;
        ws.Column(2).Width = 8;
        ws.Column(3).Width = 32;
        ws.Column(4).Width = 28;
        ws.Column(5).Width = 8;
        ws.Column(6).Width = 28;
        ws.Column(7).Width = 22;
        ws.Column(8).Width = 14;
        ws.Row(1).Height = 56;
        ws.Row(2).Height = 22;
        ws.Cells[1, 1].Style.WrapText = true;

        var statsRow = curRow;
        ws.Cells[statsRow, 1].Value = $"合计 {byTree.Count} 棵任务树 / {orders.Count} 个订单步骤";
        ws.Cells[statsRow, 1].Style.Font.Bold = true;
        var avgSteps = byTree.Count > 0 ? (double)orders.Count / byTree.Count : 0;
        ws.Cells[statsRow, 3].Value = $"平均每树 {avgSteps:F1} 步";
    }

    // ── Sheet 4：建筑→生成器→订单树 映射 ────────────────────────────────────────
    private static void BuildBuildingMappingSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("建筑→生成器映射");
        ws.View.FreezePanes(3, 1);

        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【建筑升级→生成器解锁映射】数据源：producers_inventory_slots.json（105条，含 unlockLevel）+ item_merge_graphs_full.json（orderSpawn.orderGraphTreeReference）。"
                + "unlockLevel = 玩家等级达到该值时解锁此生成器；treeId = 与该生成器绑定的订单任务树 ID（唯一对应关系，直接配置字段，高置信）。"
                + "小 ID（1~75）=早期内容，大 ID（>1M）=后期内容，均属 introOrders 同一系统。"
                + "前4条（backpack/bucket/jewelry-box/picnic-basket）unlockLevel=0，为游戏开局即有的生成器。",
            10
        );

        string[] headers = ["#", "解锁玩家等级", "生成器链(stem)", "生成器ItemID", "绑定订单树ID", "树类型", "备注",];
        string[] hdrHex = ["595959", "1F4E79", "1F4E79", "2F5496", "1F6E4A", "595959", "595959",];
        for (var j = 0; j < headers.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var producersPath = Path.Combine(DataDir, "producers_inventory_slots.json");
        var chainsPath = Path.Combine(DataDir, "item_merge_graphs_full.json");
        if (!File.Exists(producersPath) || !File.Exists(chainsPath))
        {
            ws.Cells[3, 1].Value = "数据文件缺失";
            return;
        }

        using var prodDoc = JsonDocument.Parse(File.ReadAllText(producersPath));
        using var chainDoc = JsonDocument.Parse(File.ReadAllText(chainsPath));

        // Build gen tree map: chainId → {generatorItemId, treeId}
        var genTreeMap = new Dictionary<string, (string GenItemId, string TreeId)>(
            StringComparer.Ordinal
        );
        foreach (var chain in chainDoc.RootElement.EnumerateArray())
        {
            var cid = chain.GetProperty("uniqueId").GetString() ?? "";
            foreach (var item in chain.GetProperty("items").EnumerateArray())
            {
                if (
                    !item.TryGetProperty("orderSpawn", out var os)
                    || !os.TryGetProperty("orderGraphTreeReference", out var trEl)
                )
                    continue;
                var treeId = trEl.GetString() ?? "";
                if (string.IsNullOrEmpty(treeId))
                    continue;
                var genItemId = item.GetProperty("uniqueId").GetString() ?? "";
                genTreeMap[cid] = (genItemId, treeId);
                break;
            }
        }

        var row = 3;
        var idx = 0;
        foreach (var producer in prodDoc.RootElement.EnumerateArray())
        {
            idx++;
            var producerName = producer.GetProperty("producerName").GetString() ?? "";
            var unlockLevel = producer.TryGetProperty("unlockLevel", out var ulEl)
                ? ulEl.GetInt32()
                : 0;
            var stem = producerName.Replace("item-graph_", "");

            genTreeMap.TryGetValue(producerName, out var genInfo);
            var genItemId = genInfo.GenItemId ?? "—";
            var treeId = genInfo.TreeId ?? "?";

            string treeType;
            string treeHex;
            if (treeId == "?")
            {
                treeType = "无订单树";
                treeHex = "F0F0F0";
            }
            else if (long.TryParse(treeId, out var tid))
            {
                if (tid <= 75)
                {
                    treeType = "早期内容";
                    treeHex = "D9EAD3";
                }
                else if (tid < 1_000_000)
                {
                    treeType = "中期内容";
                    treeHex = "FFF2CC";
                }
                else
                {
                    treeType = "后期内容";
                    treeHex = "FCE5CD";
                }
            }
            else
            {
                treeType = "特殊ID";
                treeHex = "E8DAEF";
            }

            var noteStr = unlockLevel == 0 ? "开局即有" : "";
            var rowHex = unlockLevel == 0 ? "EBF3FB" : null;

            XlsxStyleHelper.Cell(ws.Cells[row, 1], idx, "F5F5F5");
            XlsxStyleHelper.Cell(ws.Cells[row, 2], unlockLevel == 0 ? "—(初始)" : (object)unlockLevel, rowHex);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], stem, rowHex ?? "FAFAFA");
            XlsxStyleHelper.Cell(ws.Cells[row, 4], genItemId, "F8F8F8");
            XlsxStyleHelper.Cell(ws.Cells[row, 5], treeId, treeHex);
            XlsxStyleHelper.Cell(ws.Cells[row, 6], treeType, treeHex);
            XlsxStyleHelper.Cell(ws.Cells[row, 7], noteStr, null);
            row++;
        }

        ws.Column(1).Width = 5;
        ws.Column(2).Width = 14;
        ws.Column(3).Width = 28;
        ws.Column(4).Width = 34;
        ws.Column(5).Width = 16;
        ws.Column(6).Width = 12;
        ws.Column(7).Width = 12;
        ws.Row(1).Height = 56;
        ws.Row(2).Height = 22;
        ws.Cells[1, 1].Style.WrapText = true;

        ws.Cells[row + 1, 1].Value = $"合计 {idx} 条生成器（含 {idx - genTreeMap.Count} 条无订单树关联）";
        ws.Cells[row + 1, 1].Style.Font.Bold = true;
    }

    // ── Sheet 5：付费设计分析 ─────────────────────────────────────────────────
    private static void BuildPaymentSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("付费设计");
        ws.View.FreezePanes(3, 1);

        XlsxStyleHelper.MechNote(
            ws,
            1,
            1,
            "【Travel Town 付费设计】数据源：cache_22.dat_full.json → data.realCurrencyProducts（476条 IAP SKU，高置信直接配置）。"
                + "SKU 命名规律：traveltown.bundle.{price}.{variant} / traveltown.store.{price}.{variant}。"
                + "价格梯度：$0.49 ~ $99.99，共 54 个价格档位，476 个 SKU（每档多个变体，A/B 测试定价策略）。"
                + "另有 Stars Shop（月费忠诚积分商店，featureUserLevel=35 解锁）和卡牌册系统（cache_1A albums，7个赛季册）。",
            10
        );

        // Price tier summary
        string[] headers = ["价格(USD)", "SKU数量", "示例SKU", "类型"];
        string[] hdrHex = ["1F4E79", "2F5496", "2F5496", "595959"];
        for (var j = 0; j < headers.Length; j++)
            XlsxStyleHelper.Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var cache22Path = Path.Combine(DataDir, "cache_22.dat_full.json");
        if (!File.Exists(cache22Path))
        {
            ws.Cells[3, 1].Value = "数据文件缺失";
            return;
        }

        using var doc = JsonDocument.Parse(File.ReadAllText(cache22Path));
        var rcpEl = doc
            .RootElement.GetProperty("data")
            .GetProperty("data")
            .GetProperty("realCurrencyProducts");

        var priceTiers = new SortedDictionary<double, List<string>>();
        foreach (var p in rcpEl.EnumerateArray())
        {
            var price = p.GetProperty("price").GetDouble();
            var sku = p.GetProperty("sku").GetString() ?? "";
            if (!priceTiers.TryGetValue(price, out var list))
            {
                list = [];
                priceTiers[price] = list;
            }
            list.Add(sku);
        }

        // Color code by tier
        static string TierHex(double price) =>
            price switch
            {
                < 1.0 => "D9EAD3",
                < 5.0 => "EBF3FB",
                < 15.0 => "FFF2CC",
                < 30.0 => "FCE5CD",
                _ => "FADADD",
            };

        static string TierLabel(double price) =>
            price switch
            {
                < 1.0 => "入门档",
                < 5.0 => "低价档",
                < 15.0 => "中价档",
                < 30.0 => "高价档",
                _ => "顶级档",
            };

        var row = 3;
        foreach (var (price, skus) in priceTiers)
        {
            var hex = TierHex(price);
            var skuType = skus[0].Contains("bundle") ? "bundle包" : "store单品";
            XlsxStyleHelper.Cell(ws.Cells[row, 1], $"${price:F2}", hex);
            XlsxStyleHelper.Cell(ws.Cells[row, 2], skus.Count, hex);
            XlsxStyleHelper.Cell(ws.Cells[row, 3], skus[0], "F8F8F8");
            XlsxStyleHelper.Cell(ws.Cells[row, 4], $"{TierLabel(price)} / {skuType}", hex);
            row++;
        }

        // Summary block
        row += 2;
        var summaryItems = new[]
        {
            ("总SKU数", "476 条（直接配置数据，高置信）"),
            ("价格档位", "54 档，$0.49 ~ $99.99，每档2~29个变体（A/B测试用）"),
            ("最低价入口", "$0.49（5个SKU，降低首付门槛）"),
            ("核心主力档", "$4.99 / $9.99 / $19.99（SKU最多：16/29/29个，主推价位）"),
            ("高消费档", "$49.99（8个SKU）/ $99.99（2个SKU，大额鲸鱼用户）"),
            ("SKU命名", "bundle = 时限礼包（活动捆绑）；store = 常驻商店（钻石直购）"),
            ("Stars Shop", "月费忠诚积分商店，玩家等级35解锁，URL: loyalty.traveltowngame.com"),
            ("卡牌册系统", "7个赛季册（cache_1A albums），完成册子获得额外道具奖励"),
        };

        XlsxStyleHelper.Header(ws.Cells[row, 1, row, 4], "付费设计关键结论（高置信直接数据）", "1A3A5C");
        ws.Cells[row, 1, row, 4].Merge = true;
        row++;

        foreach (var (k, v) in summaryItems)
        {
            ws.Cells[row, 1].Value = k;
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("EBF3FB"));
            XlsxStyleHelper.Border(ws.Cells[row, 1]);

            var vc = ws.Cells[row, 2, row, 4];
            vc.Merge = true;
            vc.Value = v;
            vc.Style.WrapText = true;
            vc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            vc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("F8FAFF"));
            XlsxStyleHelper.Border(vc);
            ws.Row(row).Height = 18;
            row++;
        }

        ws.Column(1).Width = 14;
        ws.Column(2).Width = 10;
        ws.Column(3).Width = 44;
        ws.Column(4).Width = 18;
        ws.Row(1).Height = 56;
        ws.Row(2).Height = 22;
        ws.Cells[1, 1].Style.WrapText = true;
    }

    // ── Sheet 6：核心设计总结（含 Harbor 对比）────────────────────────────────
    private static void BuildSummarySheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("核心设计总结");
        var row = 1;

        void BigTitle(string text, int r)
        {
            var cell = ws.Cells[r, 1, r, 12];
            cell.Merge = true;
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 14;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("1F2D40"));
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            XlsxStyleHelper.Border(cell);
            ws.Row(r).Height = 32;
        }

        void SectionTitle(string text, int r, string hex = "2F5496")
        {
            var cell = ws.Cells[r, 1, r, 12];
            cell.Merge = true;
            cell.Value = text;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 11;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(hex));
            cell.Style.Font.Color.SetColor(Color.White);
            XlsxStyleHelper.Border(cell);
            ws.Row(r).Height = 22;
        }

        void TextBlock(
            string text,
            int r,
            string hex = "F8FAFF",
            int height = 18,
            bool bold = false
        )
        {
            var cell = ws.Cells[r, 1, r, 12];
            cell.Merge = true;
            cell.Value = text;
            cell.Style.Font.Size = 10;
            cell.Style.Font.Bold = bold;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(hex));
            cell.Style.WrapText = true;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            XlsxStyleHelper.Border(cell);
            ws.Row(r).Height = height;
        }

        // ── 大标题 ──
        BigTitle("Travel Town 核心玩法设计总结  ·  逆向工程分析报告", row++);
        TextBlock(
            "数据来源：MuMu模拟器 ADB → HTTPCache gzip-JSON + Unity AssetBundle 提取 | 游戏类型：IL2CPP Unity（非Lua）| 主配置：item_merge_graphs_full.json / orders_intro.json / boosters_config.json",
            row++,
            "E8EEF8",
            14
        );
        row++;

        // ── 一句话设计思想 ──
        SectionTitle("■ 一句话设计思想", row++, "1A3A5C");
        TextBlock(
            "「生成器升级解锁效率，元素合并推进难度，订单线性串联剧情」——整套系统靠「建筑升级路线」驱动，玩家主动推进比随机奖励更重要，每笔订单都是剧情节点而不是随机刷新。",
            row++,
            "EAF0FA",
            28,
            true
        );
        row++;

        // ── 核心循环流程图（EPPlus shapes）──
        SectionTitle("■ 完整核心循环流程图", row++, "1F4E79");

        const int StepRowCount = 3;
        const int StepRowH = 22;
        const int ArrowRowCount = 1;
        const int ArrowRowH = 12;
        const int TotalSteps = 7;
        const int BlockRows = TotalSteps * (StepRowCount + ArrowRowCount) + 2;

        for (var ir = 0; ir < BlockRows; ir++)
        {
            var isArrow = (ir % (StepRowCount + ArrowRowCount)) == StepRowCount;
            ws.Row(row + ir).Height = isArrow ? ArrowRowH : StepRowH;
        }

        var stepDefs = new (string Text, string Bg)[]
        {
            ("① 玩家点击生成器\n消耗 1 Energy → 固定产出1个初始元素（无随机，weight=100确定产出）", "1F4E79"),
            ("② 元素落盘\n棋盘上出现 Lv.0 元素，等待玩家操作", "2E75B6"),
            ("③ 玩家合并\n2个相同等级元素 → 合并为 Lv.+1 元素（Same-type merge）", "2F5496"),
            ("④ 达到订单要求等级\n根据当前任务树步骤，需要特定等级元素（difficulty 体现合成深度）", "4472C4"),
            ("⑤ 提交完成订单\n消耗元素 → 获得奖励（宝箱/钻石/工具箱）→ 解锁下一步骤", "1F6E4A"),
            ("⑥ 解锁建筑升级\n奖励工具箱 → 升级建筑 → 解锁更高等级生成器 → 产出效率提升", "276040"),
            ("⑦ 进入下一任务树\n当前树全部完成 → 自动推进到下一棵任务树（更高难度元素需求）", "145A32"),
        };

        for (var si = 0; si < stepDefs.Length; si++)
        {
            var (text, bg) = stepDefs[si];
            var baseRow = row + si * (StepRowCount + ArrowRowCount);
            var bgColor = XlsxStyleHelper.HexColor(bg);
            var darkColor = Color.FromArgb(
                Math.Max(0, bgColor.R - 25),
                Math.Max(0, bgColor.G - 25),
                Math.Max(0, bgColor.B - 25)
            );
            var boxH = StepRowCount * StepRowH;

            var shape = ws.Drawings.AddShape(
                $"TTFlowStep{si}",
                OfficeOpenXml.Drawing.eShapeStyle.RoundRect
            );
            shape.SetPosition(baseRow - 1, 3, 0, 10);
            shape.SetSize(700, boxH);
            shape.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            shape.Fill.Color = bgColor;
            shape.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            shape.Border.Fill.Color = darkColor;
            shape.Border.LineStyle = OfficeOpenXml.Drawing.eLineStyle.Solid;
            shape.Border.Width = 1.2;
            var rt = shape.RichText.Add(text);
            rt.Bold = false;
            rt.Size = 10;
            rt.Color = Color.White;
            shape.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Center;
            shape.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;

            var badge = ws.Drawings.AddShape(
                $"TTFlowBadge{si}",
                OfficeOpenXml.Drawing.eShapeStyle.Rect
            );
            badge.SetPosition(baseRow - 1, 3, 0, 10);
            badge.SetSize(26, boxH);
            badge.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            badge.Fill.Color = darkColor;
            badge.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.NoFill;

            if (si < stepDefs.Length - 1)
            {
                var arrRow = baseRow + StepRowCount;
                var arr = ws.Drawings.AddShape(
                    $"TTFlowArrow{si}",
                    OfficeOpenXml.Drawing.eShapeStyle.DownArrow
                );
                arr.SetPosition(arrRow - 1, 1, 2, 8);
                arr.SetSize(30, ArrowRowH + 2);
                arr.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                arr.Fill.Color = XlsxStyleHelper.HexColor("7F7F7F");
                arr.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.NoFill;
            }
        }

        var loopRow = row + stepDefs.Length * (StepRowCount + ArrowRowCount);
        var loopCell = ws.Cells[loopRow, 1, loopRow, 12];
        loopCell.Merge = true;
        loopCell.Value = "↑─────── 任务树完成后解锁新生成器，循环继续 ───────↑";
        loopCell.Style.Font.Bold = true;
        loopCell.Style.Font.Size = 10;
        loopCell.Style.Font.Color.SetColor(XlsxStyleHelper.HexColor("145A32"));
        loopCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        loopCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        loopCell.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("E8F5EE"));
        XlsxStyleHelper.Border(loopCell);
        ws.Row(loopRow).Height = 16;

        row += BlockRows + 1;

        // ── Travel Town vs Gossip Harbor 对比 ──
        SectionTitle("■ Travel Town vs Gossip Harbor  核心机制对比", row++, "4A235A");

        // 对比表头
        var colDefs = new[]
        {
            (1, 2, "对比维度"),
            (3, 5, "Travel Town"),
            (6, 8, "Gossip Harbor"),
            (9, 12, "设计意图差异")
        };
        foreach (var (c1, c2, label) in colDefs)
        {
            var hc = ws.Cells[row, c1, row, c2];
            hc.Merge = true;
            hc.Value = label;
            hc.Style.Font.Bold = true;
            hc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            hc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("2F5496"));
            hc.Style.Font.Color.SetColor(Color.White);
            hc.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            XlsxStyleHelper.Border(hc);
        }
        ws.Row(row++).Height = 20;

        var compRows = new[]
        {
            (
                "生成器产出",
                "固定产出（weight=100，唯一目标）\n点击即出，无随机",
                "动态权重随机（ListWeightSelectOne）\n4因子实时调整概率",
                "TT：确定性强，玩家知道在做什么\nHarbor：不确定性创造惊喜感"
            ),
            (
                "生成器升级",
                "生成器本身可合并升级（Lv.0~Lv.6）\n高级机器：容量更大/CD更短",
                "生成器不升级（固定6条链）\n通过权重配置区分链的定位",
                "TT：升级系统延伸玩家目标\nHarbor：权重系统更灵活可配"
            ),
            (
                "难度表达",
                "averageDifficulty 硬编码到配置\n（=合成路径深度，固定不变）",
                "CurWeight 运行时动态计算\n（4因子实时乘积，随行为变化）",
                "TT：可预测，设计师精确控制节奏\nHarbor：自适应，无需手动调关卡"
            ),
            (
                "订单机制",
                "线性任务树（必须按步骤顺序完成）\n103棵树，每树3~10步",
                "动态随机刷新（按进度分档）\n不看棋盘，看 DiffSum 累积",
                "TT：叙事驱动，每步是故事节点\nHarbor：无限循环，节奏靠难度档"
            ),
            (
                "订单与棋盘关系",
                "订单需要特定等级元素\n玩家必须合并到对应深度才能提交",
                "订单不看棋盘当前状态\n订单→反哺生成器权重",
                "TT：棋盘是完成目标的工具\nHarbor：棋盘状态影响权重闭环"
            ),
            (
                "进度节奏",
                "建筑升级驱动进度\n工具箱 → 建筑 → 新生成器",
                "DiffSum 累积驱动难度档升级\n无明确建筑升级路线",
                "TT：有形成就感（看得见的建筑变化）\nHarbor：渐进加速，难度无感上升"
            ),
            (
                "技术栈",
                "IL2CPP Unity，JSON配置（CDN下发）\n无Lua，C#运行时",
                "Lua 5.4 字节码，运行时动态逻辑\n配置+逻辑均在 Lua 层",
                "TT：配置与逻辑分离，运营更新方便\nHarbor：Lua热更灵活，逻辑更复杂"
            ),
        };

        foreach (var (dim, tt, harbor, diff) in compRows)
        {
            void FC(int c1, int c2, string v, string h)
            {
                var c = ws.Cells[row, c1, row, c2];
                c.Merge = true;
                c.Value = v;
                c.Style.Font.Size = 10;
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(h));
                c.Style.WrapText = true;
                c.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                XlsxStyleHelper.Border(c);
            }
            FC(1, 2, dim, "F0F4FA");
            ws.Cells[row, 1, row, 2].Style.Font.Bold = true;
            FC(3, 5, tt, "E8F5FF");
            FC(6, 8, harbor, "F0FFF0");
            FC(9, 12, diff, "FFFBEE");
            ws.Row(row++).Height = 44;
        }
        row++;

        // ── 亮点与借鉴 ──
        SectionTitle("■ Travel Town 设计亮点 & 对我们的借鉴", row++, "1A5E4A");

        var insights = new[]
        {
            (
                "亮点①  「机器即目标」",
                "生成器本身是合并对象，升级生成器就是目标。玩家始终清楚「我在做什么」——合并机器升级，不是漫无目的地刷元素。这比 Harbor 的隐性权重引导更直白。",
                "E8F8F0"
            ),
            (
                "亮点②  订单是叙事锚点",
                "103棵任务树，每棵树是一段小故事（美容、音乐、烘焙…）。每笔订单完成都有具体的叙事意义，不只是「消耗了元素」。长期留存靠故事解锁感，不靠数值刺激。",
                "E8F0F8"
            ),
            (
                "亮点③  确定性产出降低焦虑",
                "生成器 weight=100 固定产出，玩家知道点洗衣机一定出洗衣配件 Lv.0。不像 Harbor 是权重随机。确定性降低挫败感，适合休闲玩家。",
                "FFF8E8"
            ),
            (
                "可借鉴：难度配置方式",
                "difficulty.averageDifficulty 和 per-spawner difficulty 是很好的「元素获取难度」表达方式。可以在自己游戏里用类似字段：每个合成目标标注「需要点击生成器几次」，让策划精确控制每步的时间消耗。",
                "F8E8F8"
            ),
            (
                "可借鉴：生成器分层升级",
                "生成器本身分 7 级且可合并升级，为玩家创造了中期目标层次（除了元素合成之外还要合成机器）。这比 Harbor 的「静态生成器」在玩家目标多样性上更丰富。",
                "F0F0F0"
            ),
        };

        foreach (var (title, desc, hex) in insights)
        {
            var lc = ws.Cells[row, 1, row, 3];
            lc.Merge = true;
            lc.Value = title;
            lc.Style.Font.Bold = true;
            lc.Style.Font.Size = 10;
            lc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("1A5E4A"));
            lc.Style.Font.Color.SetColor(Color.White);
            lc.Style.WrapText = true;
            lc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            XlsxStyleHelper.Border(lc);

            var rc = ws.Cells[row, 4, row, 12];
            rc.Merge = true;
            rc.Value = desc;
            rc.Style.Font.Size = 10;
            rc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(hex));
            rc.Style.WrapText = true;
            rc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            XlsxStyleHelper.Border(rc);
            ws.Row(row++).Height = 44;
        }
        row++;

        // ── 数据来源 & 置信度 ──
        SectionTitle("■ 数据来源 & 置信度", row++, "595959");

        var sources = new[]
        {
            (
                "高置信（直接配置数据）",
                "produce.cycleDelay / capacity / itemsPerSubCycle / spawnTarget / weight：直接从 JSON 读取，100% 准确。\n"
                    + "difficulty.averageDifficulty 和 per-spawner difficulty：直接字段，100% 准确。\n"
                    + "orders_intro.json 全部103棵任务树字段：完整数据，「introOrders」就是完整订单系统，无单独主线循环。\n"
                    + "生成器解锁等级→订单树映射：producers_inventory_slots.json 直接字段（105条，含 unlockLevel + chainId），与 orderSpawn.orderGraphTreeReference 完全对齐（见「建筑→生成器映射」sheet）。\n"
                    + "付费 IAP 价格：cache_22.dat_full.json → realCurrencyProducts（476条 SKU，54档位，$0.49~$99.99，见「付费设计」sheet）。",
                "E8F5E8"
            ),
            (
                "高置信（结构推断）",
                "生成器产出「固定确定」（weight=100，items列表只有一项）：直接配置可验证。\n"
                    + "任务树线性依赖（lockedByIds 串联）：直接字段可验证。\n"
                    + "元素难度「硬编码」（不是运行时动态计算）：JSON 配置层无动态字段。\n"
                    + "订单系统完整性：cache_D / cache_1C 均包含相同103棵树，无其他隐藏订单系统。",
                "E8F5E8"
            ),
            (
                "中置信（推断）",
                "boosters_config.json 115个 merge-producer 的具体解锁条件：部分与 producers_inventory_slots 对应关系未完全验证。\n"
                    + "「建筑视觉升级」与「生成器解锁」的精确触发时机：meta_buildings 有视觉阶段数据，但 tasks 内 ConsumeResources 要求内容未完整解析。",
                "FFF8E8"
            ),
            (
                "暂未获取",
                "meta_buildings 建筑故事剧情文本（cache_1E 有本地化字符串但为编码乱码）。\n"
                    + "具体 IAP 礼包内容（bundle SKU对应的道具详情，只有 price 和 sku 字段，无 contents）。",
                "F0F0F0"
            ),
        };

        foreach (var (level, desc, hex) in sources)
        {
            var lc = ws.Cells[row, 1, row, 3];
            lc.Merge = true;
            lc.Value = level;
            lc.Style.Font.Bold = true;
            lc.Style.Font.Size = 10;
            lc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor("595959"));
            lc.Style.Font.Color.SetColor(Color.White);
            lc.Style.WrapText = true;
            lc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            XlsxStyleHelper.Border(lc);

            var rc = ws.Cells[row, 4, row, 12];
            rc.Merge = true;
            rc.Value = desc;
            rc.Style.Font.Size = 10;
            rc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rc.Style.Fill.BackgroundColor.SetColor(XlsxStyleHelper.HexColor(hex));
            rc.Style.WrapText = true;
            rc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            XlsxStyleHelper.Border(rc);
            ws.Row(row++).Height = 44;
        }

        // 列宽
        ws.Column(1).Width = 12;
        for (var j = 2; j <= 12; j++)
            ws.Column(j).Width = 16;
    }
}
