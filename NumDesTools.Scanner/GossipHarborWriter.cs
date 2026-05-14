using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// Gossip Harbor 竞品核心循环分析 xlsx 生成器。
/// 数据源：C:\tmp\gossipharbor_parsed2\all.json（由 parse_gh_lua.py 生成）
/// 输出：竞品-GossipHarbor核心循环分析.xlsx → Documents\workspace\
/// 包含 4 个 sheet：元素体系 / 棋盘布局 / 订单配置 / 生成器配置
/// </summary>
public static class GossipHarborWriter
{
    private const string ParsedFile = @"C:\tmp\gossipharbor_parsed2\all.json";
    private static readonly string OutputDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "workspace"
    );

    private const string OutFileName = "竞品-GossipHarbor核心循环分析.xlsx";

    // ── 样式辅助 ──────────────────────────────────────────────────────────────
    private static void Header(ExcelRange c, string text, string hex = "2F5496")
    {
        c.Value = text;
        c.Style.Fill.PatternType = ExcelFillStyle.Solid;
        c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        c.Style.Font.Bold = true;
        c.Style.Font.Color.SetColor(Color.White);
        c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        Border(c);
    }

    private static void Cell(ExcelRange c, object? value, string? hex = null)
    {
        c.Value = value;
        if (hex != null)
        {
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        }
        Border(c);
    }

    private static void Border(ExcelRange c)
    {
        var b = c.Style.Border;
        b.Top.Style = b.Bottom.Style = b.Left.Style = b.Right.Style = ExcelBorderStyle.Thin;
        b.Top.Color.SetColor(Color.FromArgb(0xBD, 0xBD, 0xBD));
        b.Bottom.Color.SetColor(Color.FromArgb(0xBD, 0xBD, 0xBD));
        b.Left.Color.SetColor(Color.FromArgb(0xBD, 0xBD, 0xBD));
        b.Right.Color.SetColor(Color.FromArgb(0xBD, 0xBD, 0xBD));
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

    private const string ParsedDir2 = @"C:\tmp\gossipharbor_parsed2";

    // ── 数据加载 ──────────────────────────────────────────────────────────────
    private static JsonElement? _root;

    private static JsonElement GetRoot()
    {
        if (_root is not null)
            return _root.Value;
        if (!File.Exists(ParsedFile))
            throw new FileNotFoundException($"找不到解析数据文件: {ParsedFile}");
        var doc = JsonDocument.Parse(File.ReadAllText(ParsedFile));
        _root = doc.RootElement;
        return _root.Value;
    }

    private static List<string> GetStrings(string key)
    {
        var root = GetRoot();
        if (!root.TryGetProperty(key, out var obj))
            return [];
        if (!obj.TryGetProperty("strings", out var arr))
            return [];
        return arr.EnumerateArray()
            .Where(e => e.ValueKind == JsonValueKind.String)
            .Select(e => e.GetString()!)
            .ToList();
    }

    private record ItemNameInfo(string Cn, string En, string Img);

    private static Dictionary<string, ItemNameInfo>? _itemNames;

    private static Dictionary<string, ItemNameInfo> GetItemNames()
    {
        if (_itemNames is not null)
            return _itemNames;
        var path = Path.Combine(ParsedDir2, "item_names_cn.json");
        if (!File.Exists(path))
            return (_itemNames = []);
        var doc = JsonDocument.Parse(File.ReadAllText(path));
        var result = new Dictionary<string, ItemNameInfo>();
        if (doc.RootElement.TryGetProperty("full", out var fullEl))
        {
            foreach (var kv in fullEl.EnumerateObject())
            {
                var cn = kv.Value.TryGetProperty("cn", out var cv) ? cv.GetString() ?? "" : "";
                var en = kv.Value.TryGetProperty("en", out var ev) ? ev.GetString() ?? "" : "";
                var img = kv.Value.TryGetProperty("img", out var iv) ? iv.GetString() ?? "" : "";
                result[kv.Name] = new ItemNameInfo(cn, en, img);
            }
        }
        return (_itemNames = result);
    }

    private static string ItemCn(string typeId) =>
        GetItemNames().TryGetValue(typeId, out var n) ? n.Cn : "";

    private static string ItemEn(string typeId) =>
        GetItemNames().TryGetValue(typeId, out var n) ? n.En : "";

    private static string ItemImg(string typeId) =>
        GetItemNames().TryGetValue(typeId, out var n) ? n.Img : "";

    // For N-M format: "一袋面粉 Lv.5" or just "一袋面粉" for base
    private static string ItemCnFull(string typeId)
    {
        var m = Regex.Match(typeId, @"^(\d+)-(\d+)$");
        if (m.Success)
        {
            var baseName = ItemCn(m.Groups[1].Value);
            return baseName.Length > 0 ? $"{baseName} Lv.{m.Groups[2].Value}" : typeId;
        }
        var name = ItemCn(typeId);
        return name.Length > 0 ? name : typeId;
    }

    private record OrderItemFull(
        string Code,
        int? UnlockLevel,
        int? Price,
        int? Weight,
        int? DiffScore,
        int? WeightMultiple,
        int? ProgressUnlock,
        int? RepeatWeightDecrease,
        int? RecommendedNumber,
        int? RecommendedPrice,
        int? DifficultyLevel
    );

    private static Dictionary<string, List<OrderItemFull>>? _orderItemsCache;

    private static Dictionary<string, List<OrderItemFull>> LoadAllOrderItems()
    {
        if (_orderItemsCache is not null)
            return _orderItemsCache;
        var path = Path.Combine(ParsedDir2, "order_items_full.json");
        if (!File.Exists(path))
            return (_orderItemsCache = new Dictionary<string, List<OrderItemFull>>());
        var root = JsonDocument.Parse(File.ReadAllText(path)).RootElement;
        var result = new Dictionary<string, List<OrderItemFull>>();
        foreach (var prop in root.EnumerateObject())
        {
            var list = new List<OrderItemFull>();
            foreach (var e in prop.Value.EnumerateArray())
            {
                int? GetInt(string name) =>
                    e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
                        ? v.GetInt32()
                        : null;
                list.Add(
                    new OrderItemFull(
                        e.GetProperty("Code").GetString()!,
                        GetInt("UnlockLevel"),
                        GetInt("Price"),
                        GetInt("Weight"),
                        GetInt("DiffScore"),
                        GetInt("WeightMultiple"),
                        GetInt("ProgressUnlock"),
                        GetInt("RepeatWeightDecrease"),
                        GetInt("RecommendedNumber"),
                        GetInt("RecommendedPrice"),
                        GetInt("DifficultyLevel")
                    )
                );
            }
            result[prop.Name] = list;
        }
        return (_orderItemsCache = result);
    }

    // ── 公开入口 ──────────────────────────────────────────────────────────────
    public static void Run(string? outputDir = null)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        var dir = outputDir ?? OutputDir;
        Directory.CreateDirectory(dir);
        var outPath = Path.Combine(dir, OutFileName);

        using var pkg = new ExcelPackage();

        BuildItemSystemSheet(pkg);
        BuildBoardSheet(pkg);
        BuildOrderSheet(pkg);
        BuildGeneratorSheet(pkg);
        BuildWeightAlgoSheet(pkg);

        pkg.SaveAs(new FileInfo(outPath));
        Console.WriteLine($"[GossipHarbor] 已生成：{outPath}");
    }

    // ── Sheet 1：元素体系（主游戏合并链）────────────────────────────────────
    // 数据源：all.json → ItemModelConfig（N-M 格式，如 101、201-35）
    private static void BuildItemSystemSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("元素体系");
        ws.View.FreezePanes(2, 1);

        var strings = GetStrings("ItemModelConfig");

        var chainMap = new Dictionary<int, SortedSet<int>>();
        foreach (var s in strings)
        {
            var m = Regex.Match(s, @"^(\d+)-(\d+)$");
            if (!m.Success)
                continue;
            var typeId = int.Parse(m.Groups[1].Value);
            var level = int.Parse(m.Groups[2].Value);
            if (level == 100)
                continue;
            if (!chainMap.ContainsKey(typeId))
                chainMap[typeId] = [];
            chainMap[typeId].Add(level);
        }

        var baseItems = new HashSet<int>(
            strings
                .Where(s =>
                    Regex.IsMatch(s, @"^\d+$")
                    && int.TryParse(s, out var n)
                    && n is >= 100 and <= 9999
                )
                .Select(int.Parse)
        );

        var chains = new List<(int TypeId, List<string> Levels)>();
        foreach (var (typeId, levels) in chainMap.OrderBy(kv => kv.Key))
        {
            var chain = new List<string> { typeId.ToString() };
            chain.AddRange(levels.Select(lv => $"{typeId}-{lv}"));
            chains.Add((typeId, chain));
        }

        var maxLen = chains.Count > 0 ? chains.Max(c => c.Levels.Count) : 1;

        Header(ws.Cells[1, 1], "TypeID", "1F4E79");
        Header(ws.Cells[1, 2], "中文名", "1F4E79");
        Header(ws.Cells[1, 3], "英文名", "1F4E79");
        Header(ws.Cells[1, 4], "图片资源名", "595959");
        Header(ws.Cells[1, 5], "链长", "1F4E79");
        Header(ws.Cells[1, 6], "最高级", "1F4E79");
        for (var j = 0; j < maxLen; j++)
            Header(ws.Cells[1, 7 + j], j == 0 ? "Lv.1(基础)" : $"Lv.{j + 1}", "2F5496");

        static string FamilyColor(int typeId) =>
            typeId switch
            {
                < 200 => "FFF2CC",
                < 300 => "D9EAD3",
                < 400 => "EBF3FB",
                < 500 => "E8DAEF",
                < 600 => "FCE5CD",
                < 700 => "D0E4F8",
                < 1000 => "FFF2CC",
                < 1200 => "E8F5E9",
                < 1500 => "FFF3E0",
                < 1700 => "E3F2FD",
                < 2000 => "FCE4EC",
                < 2500 => "F3E5F5",
                _ => "F5F5F5",
            };

        for (var i = 0; i < chains.Count; i++)
        {
            var (typeId, levels) = chains[i];
            var row = i + 2;
            var hex = FamilyColor(typeId);
            var tid = typeId.ToString();
            var maxLevel = levels.Count > 1 ? int.Parse(levels.Last().Split('-')[1]) : 1;

            Cell(ws.Cells[row, 1], typeId, "F2F2F2");
            Cell(ws.Cells[row, 2], ItemCn(tid), hex);
            Cell(ws.Cells[row, 3], ItemEn(tid), "F5F5F5");
            Cell(ws.Cells[row, 4], ItemImg(tid), "F5F5F5");
            Cell(ws.Cells[row, 5], levels.Count, hex);
            Cell(ws.Cells[row, 6], maxLevel, hex);
            for (var j = 0; j < levels.Count; j++)
                Cell(ws.Cells[row, 7 + j], levels[j], j == levels.Count - 1 ? "FADADD" : hex);
        }

        var singleItems = baseItems
            .Where(id => !chainMap.ContainsKey(id))
            .OrderBy(id => id)
            .ToList();

        var scatterCol = 7 + maxLen + 1;
        Header(ws.Cells[1, scatterCol], "散件TypeID", "595959");
        Header(ws.Cells[1, scatterCol + 1], "中文名", "595959");
        Header(ws.Cells[1, scatterCol + 2], "英文名", "595959");
        for (var i = 0; i < singleItems.Count; i++)
        {
            var row = i + 2;
            var tid = singleItems[i].ToString();
            Cell(ws.Cells[row, scatterCol], singleItems[i], "F5F5F5");
            Cell(ws.Cells[row, scatterCol + 1], ItemCn(tid), "EBF3FB");
            Cell(ws.Cells[row, scatterCol + 2], ItemEn(tid), "F5F5F5");
        }

        ws.Column(1).Width = 10;
        ws.Column(2).Width = 20;
        ws.Column(3).Width = 28;
        ws.Column(4).Width = 22;
        ws.Column(5).Width = 8;
        ws.Column(6).Width = 8;
        for (var j = 0; j < maxLen; j++)
            ws.Column(7 + j).Width = 12;
        ws.Column(scatterCol).Width = 10;
        ws.Column(scatterCol + 1).Width = 20;
        ws.Column(scatterCol + 2).Width = 26;

        var statRow = chains.Count + 3;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"合并链 {chains.Count} 条";
        ws.Cells[statRow, 5].Value = $"散件 {singleItems.Count} 个";
        ws.Cells[statRow, 6].Value = $"最长链 {maxLen} 级";
        ws.Cells[statRow, 7].Value =
            $"合计元素类型 {baseItems.Count + chainMap.Values.Sum(s => s.Count)} 种";
    }

    // ── 棋盘网格 JSON 加载 ────────────────────────────────────────────────────
    private static Dictionary<string, List<List<string>>>? _boardGridCache;

    private static Dictionary<string, List<List<string>>> LoadBoardGrids()
    {
        if (_boardGridCache is not null)
            return _boardGridCache;
        var path = Path.Combine(ParsedDir2, "board_grids.json");
        if (!File.Exists(path))
            return (_boardGridCache = []);
        var root = JsonDocument.Parse(File.ReadAllText(path)).RootElement;
        var result = new Dictionary<string, List<List<string>>>();
        foreach (var prop in root.EnumerateObject())
        {
            if (prop.Name.Contains("_sub"))
                continue;
            var grid = new List<List<string>>();
            foreach (var rowEl in prop.Value.EnumerateArray())
            {
                var row = rowEl.EnumerateArray().Select(e => e.GetString() ?? "").ToList();
                grid.Add(row);
            }
            result[prop.Name] = grid;
        }
        return (_boardGridCache = result);
    }

    // ── Sheet 2：棋盘布局（2D 网格）──────────────────────────────────────────
    // pb# = 生成器(producer), c# = 初始锁定(cloud-locked), 0 = 空格
    private static void BuildBoardSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("棋盘布局");

        MechNote(
            ws,
            1,
            1,
            "【棋盘格子说明】pb# = 场景内生成器（自动/手动产出初始元素）；c# = 初始锁定（需解锁区域才显现）；"
                + "pb#c# = 同时有生成器且初始锁定；0 = 当前空格（可填充区域）；"
                + "无前缀 = 已放置的元素/宝箱/传送门。颜色：绿=生成器 橙=锁定 蓝=传送门 灰=空格 白=普通元素。",
            20
        );

        var grids = LoadBoardGrids();
        var boardOrder = new[] { "bird_1", "bird_2", "bird_3", "bird_4" };
        var boardLabels = new Dictionary<string, string>
        {
            ["bird_1"] = "主棋盘 阶段1 (5×5)",
            ["bird_2"] = "主棋盘 阶段2 (8×5)",
            ["bird_3"] = "主棋盘 阶段3 (8×6)",
            ["bird_4"] = "主棋盘 阶段4 (8×6)",
        };

        var startRow = 2;

        foreach (var boardKey in boardOrder)
        {
            if (!grids.TryGetValue(boardKey, out var grid) || grid.Count == 0)
                continue;

            var label = boardLabels.GetValueOrDefault(boardKey, boardKey);
            var rows = grid.Count;
            var cols = grid.Max(r => r.Count);

            // Board title
            var titleCell = ws.Cells[startRow, 1, startRow, cols + 2];
            titleCell.Merge = true;
            titleCell.Value = label;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.Size = 11;
            titleCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleCell.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
            titleCell.Style.Font.Color.SetColor(Color.White);
            titleCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Border(titleCell);
            startRow++;

            // Column headers (1-based col numbers)
            Header(ws.Cells[startRow, 1], "行\\列", "595959");
            for (var c = 0; c < cols; c++)
                Header(ws.Cells[startRow, c + 2], $"C{c + 1}", "2F5496");
            startRow++;

            // Grid data rows
            for (var r = 0; r < rows; r++)
            {
                var rowLabel = ws.Cells[startRow + r, 1];
                rowLabel.Value = $"R{r + 1}";
                rowLabel.Style.Font.Bold = true;
                rowLabel.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rowLabel.Style.Fill.BackgroundColor.SetColor(HexColor("2F5496"));
                rowLabel.Style.Font.Color.SetColor(Color.White);
                rowLabel.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Border(rowLabel);

                for (var c = 0; c < cols; c++)
                {
                    var raw = c < grid[r].Count ? grid[r][c] : "";
                    var cell = ws.Cells[startRow + r, c + 2];

                    var isPb = raw.Contains("pb#");
                    var isLocked = Regex.IsMatch(raw, @"(?<![a-z])c#");
                    var isEmpty = raw == "0" || raw == "";
                    var isPortal = raw.Contains("portal");

                    // Cell background color
                    string? cellHex = (isPb, isLocked, isEmpty, isPortal) switch
                    {
                        (_, _, true, _) => "F0F0F0", // empty
                        (_, _, _, true) => "D0E4F8", // portal = blue
                        (true, true, _, _) => "F9CB9C", // generator + locked = orange-green
                        (true, false, _, _) => "D9EAD3", // generator only = green
                        (false, true, _, _) => "FCE5CD", // locked only = orange
                        _ => null, // normal element = white
                    };

                    // Display text: strip prefixes, show element code
                    var itemCode = raw.Replace("pb#", "").Replace("c#", "");
                    var displayText = isEmpty ? "" : itemCode;

                    // Try to get a short label
                    var cnName = ItemCn(itemCode);
                    if (cnName.Length == 0)
                    {
                        var baseCode = Regex.Match(itemCode, @"^(\d+)").Groups[1].Value;
                        if (baseCode.Length > 0)
                            cnName = ItemCn(baseCode);
                    }

                    // Two-line display: type indicator + code, then Chinese name
                    var typeTag = (isPb, isLocked) switch
                    {
                        (true, true) => "⚙🔒",
                        (true, false) => "⚙",
                        (false, true) => "🔒",
                        _ => "",
                    };
                    var line1 = isEmpty
                        ? ""
                        : (typeTag.Length > 0 ? typeTag + " " : "") + displayText;
                    var line2 = cnName.Length > 0 ? cnName : "";
                    cell.Value = isEmpty ? "" : (line2.Length > 0 ? line1 + "\n" + line2 : line1);
                    if (cellHex != null)
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(HexColor(cellHex));
                    }
                    cell.Style.Font.Size = 8;
                    cell.Style.WrapText = true;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    Border(cell);
                }

                ws.Row(startRow + r).Height = 38;
            }

            // Statistics row
            var statRow = startRow + rows;
            var allCells = grid.SelectMany(r => r).ToList();
            var pbCount = allCells.Count(s => s.Contains("pb#"));
            var lockCount = allCells.Count(s => Regex.IsMatch(s, @"(?<![a-z])c#"));
            var emptyCount = allCells.Count(s => s == "0");
            ws.Cells[statRow, 1].Value = "统计";
            ws.Cells[statRow, 1].Style.Font.Bold = true;
            ws.Cells[statRow, 2].Value =
                $"总格: {rows * cols} | 生成器(pb): {pbCount} | 锁定(c): {lockCount} | 空格(0): {emptyCount}";
            ws.Cells[statRow, 2, statRow, cols + 1].Merge = true;
            ws.Cells[statRow, 2].Style.Font.Size = 9;

            startRow = statRow + 2; // gap between boards
        }

        // Column widths: row label col + data cols
        ws.Column(1).Width = 6;
        for (var c = 2; c <= 8; c++)
            ws.Column(c).Width = 22;
        ws.Row(1).Height = 40;
        ws.Cells[1, 1].Style.WrapText = true;

        // Color legend row at the end
        var legendRow = startRow;
        ws.Cells[legendRow, 1].Value = "图例：";
        ws.Cells[legendRow, 1].Style.Font.Bold = true;
        var legends = new[]
        {
            ("[G] 绿=生成器", "D9EAD3"),
            ("[L] 橙=锁定", "FCE5CD"),
            ("[G+L] 橙绿=生成器+锁定", "F9CB9C"),
            ("蓝=传送门", "D0E4F8"),
            ("灰=空格", "F0F0F0"),
        };
        for (var i = 0; i < legends.Length; i++)
        {
            var (text, hex) = legends[i];
            var lCell = ws.Cells[legendRow, i + 2];
            lCell.Value = text;
            lCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lCell.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
            lCell.Style.Font.Size = 9;
            Border(lCell);
        }
    }

    // ── Sheet 3：订单配置（含机制说明） ────────────────────────────────────────
    private static void BuildOrderSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("订单配置");
        ws.View.FreezePanes(4, 1);

        var allItems = LoadAllOrderItems();
        var hardItems = allItems.TryGetValue("hard", out var h) ? h : [];
        var contiItems = allItems.TryGetValue("conti", out var c) ? c : [];

        // ── 机制说明区（Row 1-2）──
        MechNote(
            ws,
            1,
            1,
            "【订单生成机制】每轮随机 1~3 个物品槽，槽1(First)难度最高，槽2(Second)次之，槽3(Third)最低。"
                + "DiffScore=难度分，订单总难度=各槽DiffScore之和，决定刷新CD。"
                + "Weight=权重（100=正常），WeightMultiple=权重倍率（1=双倍），"
                + "RepeatWeightDecrease=连续抽中后降权幅度。",
            14
        );
        MechNote(
            ws,
            2,
            1,
            "【好评机制(ReviewRace)】完成订单获得1个好评(笑脸)。收集到目标数量后可激活【好评浪潮】(OrderBoost)获得奖励加成。"
                + "RecommendedNumber=此元素推荐同批提交订单数，RecommendedPrice=推荐价格(单件×推荐数)。"
                + "ProgressUnlock=1表示需进度解锁，DifficultyLevel=系统难度等级(越高越难出现)。",
            14
        );

        // ── 合并表头（Row 3）──
        var sections = new[]
        {
            (Col: 1, Width: 11, Label: "Hard模式（chesthard/宝箱订单）"),
            (Col: 13, Width: 11, Label: "Conti模式（持续订单/连续订单）"),
        };
        foreach (var s in sections)
        {
            ws.Cells[3, s.Col].Value = s.Label;
            ws.Cells[3, s.Col].Style.Font.Bold = true;
            ws.Cells[3, s.Col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[3, s.Col].Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
            ws.Cells[3, s.Col].Style.Font.Color.SetColor(Color.White);
            ws.Cells[3, s.Col, 3, s.Col + s.Width - 1].Merge = true;
            Border(ws.Cells[3, s.Col, 3, s.Col + s.Width - 1]);
        }

        // ── 列标题（Hard, Row 4）──
        int hc = 1;
        Header(ws.Cells[4, hc++], "#", "1F4E79");
        Header(ws.Cells[4, hc++], "Code", "2F5496");
        Header(ws.Cells[4, hc++], "中文名", "1F4E79");
        Header(ws.Cells[4, hc++], "英文名", "2F5496");
        Header(ws.Cells[4, hc++], "售价", "2F5496");
        Header(ws.Cells[4, hc++], "难度分", "2F5496");
        Header(ws.Cells[4, hc++], "权重W", "595959");
        Header(ws.Cells[4, hc++], "推荐数", "595959");
        Header(ws.Cells[4, hc++], "进度锁", "595959");
        Header(ws.Cells[4, hc++], "难等级", "595959");
        Header(ws.Cells[4, hc], "降权幅", "595959");

        // ── 列标题（Conti, Row 4）──
        int cc = 13;
        Header(ws.Cells[4, cc++], "#", "1F4E79");
        Header(ws.Cells[4, cc++], "Code", "2F5496");
        Header(ws.Cells[4, cc++], "中文名", "1F4E79");
        Header(ws.Cells[4, cc++], "英文名", "2F5496");
        Header(ws.Cells[4, cc++], "售价", "2F5496");
        Header(ws.Cells[4, cc++], "难度分", "2F5496");
        Header(ws.Cells[4, cc++], "权重W", "595959");
        Header(ws.Cells[4, cc++], "推荐数", "595959");
        Header(ws.Cells[4, cc++], "进度锁", "595959");
        Header(ws.Cells[4, cc++], "难等级", "595959");
        Header(ws.Cells[4, cc], "DiffΔ(vs Hard)", "2F5496");

        static string ItemHex(int code) =>
            code switch
            {
                < 300 => "EBF3FB",
                < 400 => "D9EAD3",
                < 600 => "FFF2CC",
                < 700 => "FCE5CD",
                < 1000 => "E8DAEF",
                < 1200 => "D5E8D4",
                < 1500 => "FFE6CC",
                < 1800 => "DAE8FC",
                < 2000 => "F8CECC",
                _ => "F5F5F5",
            };

        // Hard 侧按 TypeID 排序
        var hardSorted = hardItems.OrderBy(x => int.TryParse(x.Code, out var n) ? n : 0).ToList();
        var hardByCode = hardSorted.ToDictionary(x => x.Code);
        var contiSorted = contiItems.OrderBy(x => int.TryParse(x.Code, out var n) ? n : 0).ToList();

        for (var i = 0; i < hardSorted.Count; i++)
        {
            var it = hardSorted[i];
            var row = i + 5;
            int.TryParse(it.Code, out var codeInt);
            var hex = ItemHex(codeInt);

            int col = 1;
            Cell(ws.Cells[row, col++], i + 1, "F2F2F2");
            Cell(ws.Cells[row, col++], it.Code, hex);
            Cell(ws.Cells[row, col++], ItemCn(it.Code), hex);
            Cell(ws.Cells[row, col++], ItemEn(it.Code), "F5F5F5");
            Cell(ws.Cells[row, col++], it.Price, it.Price.HasValue ? hex : null);
            Cell(ws.Cells[row, col++], it.DiffScore, hex);
            Cell(
                ws.Cells[row, col++],
                it.Weight == 0 ? "禁用(0)" : it.Weight?.ToString() ?? "",
                it.Weight == 0 ? "FADADD" : "F9F9F9"
            );
            Cell(
                ws.Cells[row, col++],
                it.RecommendedNumber.HasValue ? $"×{it.RecommendedNumber}" : "",
                it.RecommendedNumber.HasValue ? "FFF2CC" : null
            );
            Cell(
                ws.Cells[row, col++],
                it.ProgressUnlock == 1 ? "✓" : "",
                it.ProgressUnlock == 1 ? "FCE5CD" : null
            );
            Cell(ws.Cells[row, col++], it.DifficultyLevel, "F9F9F9");
            Cell(ws.Cells[row, col], it.RepeatWeightDecrease, "F9F9F9");
        }

        for (var i = 0; i < contiSorted.Count; i++)
        {
            var it = contiSorted[i];
            var row = i + 5;
            int.TryParse(it.Code, out var codeInt);
            var hex = ItemHex(codeInt);

            hardByCode.TryGetValue(it.Code, out var hardIt);
            var diffDelta =
                hardIt?.DiffScore.HasValue == true && it.DiffScore.HasValue
                    ? it.DiffScore.Value - hardIt.DiffScore.Value
                    : (int?)null;
            var diffStr = diffDelta.HasValue
                ? diffDelta.Value == 0
                    ? "="
                    : diffDelta.Value > 0
                        ? $"+{diffDelta}"
                        : $"{diffDelta}"
                : "";
            var diffHex = diffDelta is > 0
                ? "FCE5CD"
                : diffDelta is < 0
                    ? "D9EAD3"
                    : hex;

            int col = 13;
            Cell(ws.Cells[row, col++], i + 1, "F2F2F2");
            Cell(ws.Cells[row, col++], it.Code, hex);
            Cell(ws.Cells[row, col++], ItemCn(it.Code), hex);
            Cell(ws.Cells[row, col++], ItemEn(it.Code), "F5F5F5");
            Cell(ws.Cells[row, col++], it.Price, it.Price.HasValue ? hex : null);
            Cell(ws.Cells[row, col++], it.DiffScore, hex);
            Cell(
                ws.Cells[row, col++],
                it.Weight == 0 ? "禁用(0)" : it.Weight?.ToString() ?? "",
                it.Weight == 0 ? "FADADD" : "F9F9F9"
            );
            Cell(
                ws.Cells[row, col++],
                it.RecommendedNumber.HasValue ? $"×{it.RecommendedNumber}" : "",
                it.RecommendedNumber.HasValue ? "FFF2CC" : null
            );
            Cell(
                ws.Cells[row, col++],
                it.ProgressUnlock == 1 ? "✓" : "",
                it.ProgressUnlock == 1 ? "FCE5CD" : null
            );
            Cell(ws.Cells[row, col++], it.DifficultyLevel, "F9F9F9");
            Cell(ws.Cells[row, col], diffStr, diffHex);
        }

        // 列宽
        int[] hardColWidths = [5, 8, 18, 26, 8, 8, 9, 9, 8, 8, 8];
        for (var j = 0; j < hardColWidths.Length; j++)
            ws.Column(1 + j).Width = hardColWidths[j];
        ws.Column(12).Width = 2;
        int[] contiColWidths = [5, 8, 18, 26, 8, 8, 9, 9, 8, 8, 14];
        for (var j = 0; j < contiColWidths.Length; j++)
            ws.Column(13 + j).Width = contiColWidths[j];

        // 统计
        var statRow = Math.Max(hardSorted.Count, contiSorted.Count) + 6;
        ws.Cells[statRow, 1].Value = "统计";
        ws.Cells[statRow, 1].Style.Font.Bold = true;
        ws.Cells[statRow, 2].Value = $"Hard: {hardSorted.Count} 种";
        ws.Cells[statRow, 3].Value = $"Conti: {contiSorted.Count} 种";
        if (hardSorted.Count > 0)
        {
            ws.Cells[statRow, 4].Value =
                $"Hard价格: {hardSorted.Min(x => x.Price ?? 0)}～{hardSorted.Max(x => x.Price ?? 0)}";
            ws.Cells[statRow, 5].Value =
                $"Hard难度分: {hardSorted.Min(x => x.DiffScore ?? 0)}～{hardSorted.Max(x => x.DiffScore ?? 0)}";
        }
        ws.Row(1).Height = 50;
        ws.Row(2).Height = 50;
        ws.Cells[1, 1].Style.WrapText = true;
        ws.Cells[2, 1].Style.WrapText = true;

        // ── 订单槽生成机制（Row: 追加在数据末尾）──
        var slotRow = Math.Max(hardSorted.Count, contiSorted.Count) + 8;
        MechNote(
            ws,
            slotRow,
            1,
            "【订单槽生成规则（ChestOrderRandomConfig / OrderRandomConfig）】"
                + "系统按玩家进度(PlayerLevel)选择对应配置行，每个订单随机1~3个物品槽。"
                + "FirstMin/Max=槽1难度分范围，SecondMin/Max=槽2，ThirdMin/Max=槽3。"
                + "OneItemWeight/TwoItemsWeight/ThreeItemsWeight=1格/2格/3格订单权重（宝箱模式专用）。"
                + "DiffSumMin/Max=该玩家总难度分范围门槛（普通模式专用，限制订单难度上下界）。"
                + "RefreshTime=订单刷新CD(秒)，Point=完成订单获得的好评点数。",
            14
        );

        // 槽配置表头
        slotRow++;
        var slotLabel = ws.Cells[slotRow, 1, slotRow, 14];
        slotLabel.Merge = true;
        slotLabel.Value = "▶ 订单槽概率配置（lesscd=标准普通订单 / hard=宝箱专属订单）";
        slotLabel.Style.Font.Bold = true;
        slotLabel.Style.Fill.PatternType = ExcelFillStyle.Solid;
        slotLabel.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
        slotLabel.Style.Font.Color.SetColor(Color.White);
        Border(slotLabel);
        slotRow++;

        // Two sub-tables side by side: lesscd (left) and hard (right)
        int sh = slotRow;
        string[] slotHeaders =
        [
            "来源",
            "Type",
            "玩家级别",
            "刷新CD",
            "难度总和范围",
            "槽1难度(Min-Max)",
            "槽2难度(Max)",
            "槽3难度(Max)",
            "1格权重",
            "2格权重",
            "3格权重",
        ];
        for (var j = 0; j < slotHeaders.Length; j++)
        {
            Header(ws.Cells[sh, j + 1], slotHeaders[j], "2F5496");
        }
        sh++;

        var orderRandomData = LoadOrderRandomConfig();
        static string PlayerRange(JsonElement rec)
        {
            var mn =
                rec.TryGetProperty("PlayerMin", out var v) && v.ValueKind == JsonValueKind.Number
                    ? v.GetInt32().ToString()
                    : "0";
            var mx =
                rec.TryGetProperty("PlayerMax", out var v2) && v2.ValueKind == JsonValueKind.Number
                    ? v2.GetInt32().ToString()
                    : "∞";
            return $"{mn}~{mx}";
        }

        static string DiffRange(JsonElement rec, string minKey, string maxKey)
        {
            var mn =
                rec.TryGetProperty(minKey, out var v) && v.ValueKind == JsonValueKind.Number
                    ? v.GetInt32().ToString()
                    : "—";
            var mx =
                rec.TryGetProperty(maxKey, out var v2) && v2.ValueKind == JsonValueKind.Number
                    ? v2.GetInt32().ToString()
                    : "—";
            if (mn == "—" && mx == "—")
                return "—";
            return $"{mn}~{mx}";
        }

        static string GetN(JsonElement rec, string key)
        {
            if (rec.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.Number)
                return v.GetInt32().ToString();
            return "—";
        }

        foreach (
            var (srcName, srcLabel) in new[] { ("lesscd", "lesscd(标准)"), ("hard", "hard(宝箱)") }
        )
        {
            if (!orderRandomData.TryGetValue(srcName, out var srcRecs))
                continue;
            foreach (var rec in srcRecs.EnumerateArray())
            {
                var hasSlotData =
                    rec.TryGetProperty("FirstMax", out _) || rec.TryGetProperty("FirstMin", out _);
                if (!hasSlotData)
                    continue;

                var type = GetN(rec, "Type");
                var cd = GetN(rec, "RefreshTime");
                var playerRange = PlayerRange(rec);
                var diffSum = DiffRange(rec, "DiffSumMin", "DiffSumMax");
                var slot1 = DiffRange(rec, "FirstMin", "FirstMax");
                var slot2Max = GetN(rec, "SecondMax");
                var slot3Max = GetN(rec, "ThirdMax");
                var w1 = GetN(rec, "OneItemWeight");
                var w2 = GetN(rec, "TwoItemsWeight");
                var w3 = GetN(rec, "ThreeItemsWeight");

                var rowHex =
                    srcName == "hard"
                        ? "FFF2CC"
                        : playerRange.StartsWith("0~") || playerRange.StartsWith("~")
                            ? "EBF3FB"
                            : "F5F5F5";

                Cell(ws.Cells[sh, 1], srcLabel, rowHex);
                Cell(ws.Cells[sh, 2], type == "—" ? "" : type, rowHex);
                Cell(ws.Cells[sh, 3], playerRange, rowHex);
                Cell(
                    ws.Cells[sh, 4],
                    cd == "—"
                        ? ""
                        : int.TryParse(cd, out var cdInt)
                            ? cdInt
                            : (object)cd,
                    rowHex
                );
                Cell(ws.Cells[sh, 5], diffSum, rowHex);
                Cell(ws.Cells[sh, 6], slot1, rowHex);
                Cell(ws.Cells[sh, 7], slot2Max == "—" ? "" : slot2Max, rowHex);
                Cell(ws.Cells[sh, 8], slot3Max == "—" ? "" : slot3Max, rowHex);
                Cell(
                    ws.Cells[sh, 9],
                    w1 == "—" ? "" : w1,
                    w1 != "—" && w1 != "0" ? "D9EAD3" : rowHex
                );
                Cell(
                    ws.Cells[sh, 10],
                    w2 == "—" ? "" : w2,
                    w2 != "—" && w2 != "0" ? "FFF2CC" : rowHex
                );
                Cell(
                    ws.Cells[sh, 11],
                    w3 == "—" ? "" : w3,
                    w3 != "—" && w3 != "0" ? "FCE5CD" : rowHex
                );
                sh++;
            }
        }

        // Column widths for slot table
        int[] slotColWidths = [14, 6, 14, 10, 20, 18, 12, 12, 10, 10, 10];
        for (var j = 0; j < slotColWidths.Length; j++)
            ws.Column(1 + j).Width = Math.Max(ws.Column(1 + j).Width, slotColWidths[j]);

        ws.Row(slotRow - 1).Height = 60;
    }

    private static JsonElement _orderRandomRoot;
    private static bool _orderRandomLoaded;

    private static Dictionary<string, JsonElement> LoadOrderRandomConfig()
    {
        if (_orderRandomLoaded)
            return _orderRandomRoot.EnumerateObject().ToDictionary(p => p.Name, p => p.Value);
        var path = Path.Combine(ParsedDir2, "order_random_config.json");
        if (!File.Exists(path))
        {
            _orderRandomLoaded = true;
            return [];
        }
        var doc = JsonDocument.Parse(File.ReadAllText(path));
        _orderRandomRoot = doc.RootElement;
        _orderRandomLoaded = true;
        return _orderRandomRoot.EnumerateObject().ToDictionary(p => p.Name, p => p.Value);
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

    // ── 主棋盘生成器链数据加载 ────────────────────────────────────────────────
    // 数据源：main_generators.json（由 merge_main_item.py 生成）
    // 结构：{ "链头TypeID": [ {Type, MergedType, Spread_*, ...}, ... ] }
    private record MainElemInfo(
        string Type,
        string MergedType,
        int SpreadWeightType,
        int SpreadItemMax,
        int SpreadStorageMax,
        int SpreadCostEnergy,
        int SpreadItemRecovery,
        int SpreadStorageRecovery,
        int SpreadItemSpeedUpCost,
        int SpreadStorageSpeedUpCost,
        int SellingPrice,
        int Rare,
        string GenChain1,
        string GenChain2,
        int BubbleChance,
        int LevelUpPiece,
        int LevelDownPiece,
        int SpreadAuto
    );

    private static Dictionary<string, List<MainElemInfo>>? _mainGenCache;

    private static Dictionary<string, List<MainElemInfo>> LoadMainGenerators()
    {
        if (_mainGenCache is not null)
            return _mainGenCache;
        var path = Path.Combine(ParsedDir2, "main_generators.json");
        if (!File.Exists(path))
            return (_mainGenCache = []);
        var root = JsonDocument.Parse(File.ReadAllText(path)).RootElement;
        var result = new Dictionary<string, List<MainElemInfo>>();
        foreach (var chainProp in root.EnumerateObject())
        {
            var chain = new List<MainElemInfo>();
            foreach (var el in chainProp.Value.EnumerateArray())
            {
                int GetI(string k) =>
                    el.TryGetProperty(k, out var v) && v.ValueKind == JsonValueKind.Number
                        ? v.GetInt32()
                        : 0;
                string GetS(string k) =>
                    el.TryGetProperty(k, out var v)
                        ? v.ValueKind == JsonValueKind.String
                            ? v.GetString() ?? ""
                            : v.ValueKind == JsonValueKind.Number
                                ? v.GetInt32().ToString()
                                : ""
                        : "";
                chain.Add(
                    new MainElemInfo(
                        GetS("Type"),
                        GetS("MergedType"),
                        GetI("Spread_WeightType"),
                        GetI("Spread_ItemMaxNumber"),
                        GetI("Spread_StorageMaxNumber"),
                        GetI("Spread_CostEnergy"),
                        GetI("Spread_ItemRecoveryDuration"),
                        GetI("Spread_StorageRecoveryDuration"),
                        GetI("Spread_ItemSpeedUpCost"),
                        GetI("Spread_StorageSpeedUpCost"),
                        GetI("SellingPrice"),
                        GetI("Rare"),
                        GetS("FirstLevelGeneratorChain"),
                        GetS("SecondLevelGeneratorChain"),
                        GetI("BubbleChance"),
                        GetI("LevelUpPiece"),
                        GetI("LevelDownPiece"),
                        GetI("Spread_Auto")
                    )
                );
            }
            result[chainProp.Name] = chain;
        }
        return (_mainGenCache = result);
    }

    // ── Sheet 4：生成器配置（主棋盘产出链）────────────────────────────────────
    private static void BuildGeneratorSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("生成器配置");
        ws.View.FreezePanes(3, 1);

        MechNote(
            ws,
            1,
            1,
            "【主棋盘生成器产出机制】每次生成器触发时，从整条产出链中按权重随机选一个元素产出（不是固定产最低级）。"
                + "权重类型(WeightType)：1=高频(最常出) / 2=普通 / 3=低频 / 4=极低(稀有)。"
                + "同一条链里 WeightType 越小的元素被抽中概率越高，越高级的元素 WeightType 越大概率越低。"
                + "GenChain=生成器链ID（每条链的标识符，FirstLevel=主力生成器，SecondLevel=辅助生成器）。"
                + "格子CD=自动产出间隔秒数；格子上限=棋盘同时存放上限；仓库上限=仓库存放上限。",
            12
        );

        var chains = LoadMainGenerators();

        // 列标题
        string[] colHeaders =
        [
            "链头编码",
            "链中位置",
            "元素编码",
            "中文名",
            "售价(金币)",
            "稀有度",
            "产出权重(1高~4低)",
            "格子上限",
            "仓库上限",
            "格子CD(s)",
            "仓库CD(s)",
            "加速消耗",
            "气泡概率%",
            "升级所需",
            "降级所需",
            "自动产出",
            "合并后变为",
        ];
        string[] colHexes =
        [
            "1F4E79",
            "595959",
            "2F5496",
            "1F4E79",
            "2F5496",
            "595959",
            "595959",
            "595959",
            "595959",
            "595959",
            "595959",
            "595959",
            "2F5496",
            "595959",
            "595959",
            "595959",
            "2F5496",
        ];
        for (var j = 0; j < colHeaders.Length; j++)
            Header(ws.Cells[2, j + 1], colHeaders[j], colHexes[j]);

        // 链颜色（按链头区分）
        var chainColors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var curRow = 3;
        var chainIdx = 0;

        foreach (var (headType, chain) in chains.OrderBy(kv => int.Parse(kv.Key)))
        {
            var baseHex = chainColors[chainIdx % chainColors.Length];
            chainIdx++;

            // 区块标题行
            var headElem = chain.Count > 0 ? chain[0] : null;
            var gc1 = headElem?.GenChain1 is { Length: > 0 } g1 ? $"主力生成器链ID={g1}" : "";
            var gc2 = headElem?.GenChain2 is { Length: > 0 } g2 ? $"  辅助链ID={g2}" : "";
            var titleCell = ws.Cells[curRow, 1, curRow, colHeaders.Length];
            titleCell.Merge = true;
            titleCell.Value =
                $"▶ 产出链  链头={headType}({ItemCn(headType)})  共{chain.Count}级"
                + (gc1.Length > 0 ? $"  {gc1}{gc2}" : "");
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            titleCell.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
            titleCell.Style.Font.Color.SetColor(Color.White);
            Border(titleCell);
            curRow++;

            for (var step = 0; step < chain.Count; step++)
            {
                var elem = chain[step];
                var isHead = step == 0;
                var isTail = step == chain.Count - 1;
                var rowHex = isHead
                    ? "D9EAD3"
                    : isTail
                        ? "FADADD"
                        : baseHex;

                var posLabel = isHead
                    ? "链头"
                    : isTail
                        ? $"Lv.{step + 1}(末端)"
                        : $"Lv.{step + 1}";

                var cnName = ItemCn(elem.Type);

                var mergedDisplay = elem.MergedType.Length > 0 ? elem.MergedType : "（终点）";
                var mergedHex = elem.MergedType.Length > 0 ? "EBF3FB" : "F0F0F0";

                Cell(ws.Cells[curRow, 1], isHead ? headType : "", isHead ? "D9EAD3" : "F5F5F5");
                Cell(ws.Cells[curRow, 2], posLabel, "F5F5F5");
                Cell(ws.Cells[curRow, 3], elem.Type, rowHex);
                Cell(ws.Cells[curRow, 4], cnName.Length > 0 ? cnName : "—", rowHex);
                Cell(
                    ws.Cells[curRow, 5],
                    elem.SellingPrice > 0 ? elem.SellingPrice : (object)"—",
                    elem.SellingPrice > 0 ? "FFF2CC" : null
                );
                Cell(
                    ws.Cells[curRow, 6],
                    elem.Rare > 0 ? elem.Rare : (object)"—",
                    elem.Rare >= 3 ? "E8DAEF" : null
                );
                Cell(
                    ws.Cells[curRow, 7],
                    elem.SpreadWeightType > 0 ? elem.SpreadWeightType : (object)"—",
                    null
                );
                Cell(
                    ws.Cells[curRow, 8],
                    elem.SpreadItemMax > 0 ? elem.SpreadItemMax : (object)"—",
                    elem.SpreadItemMax > 0 ? "EBF3FB" : null
                );
                Cell(
                    ws.Cells[curRow, 9],
                    elem.SpreadStorageMax > 0 ? elem.SpreadStorageMax : (object)"—",
                    elem.SpreadStorageMax > 0 ? "EBF3FB" : null
                );
                Cell(
                    ws.Cells[curRow, 10],
                    elem.SpreadItemRecovery > 0 ? elem.SpreadItemRecovery : (object)"—",
                    elem.SpreadItemRecovery > 0 ? "FFF2CC" : null
                );
                Cell(
                    ws.Cells[curRow, 11],
                    elem.SpreadStorageRecovery > 0 ? elem.SpreadStorageRecovery : (object)"—",
                    elem.SpreadStorageRecovery > 0 ? "FFF2CC" : null
                );
                var speedUp =
                    elem.SpreadItemSpeedUpCost > 0 || elem.SpreadStorageSpeedUpCost > 0
                        ? $"格:{elem.SpreadItemSpeedUpCost} 仓:{elem.SpreadStorageSpeedUpCost}"
                        : "—";
                Cell(ws.Cells[curRow, 12], speedUp, speedUp != "—" ? "FCE5CD" : null);
                Cell(
                    ws.Cells[curRow, 13],
                    elem.BubbleChance > 0 ? $"{elem.BubbleChance}%" : (object)"—",
                    elem.BubbleChance > 0 ? "D9EAD3" : null
                );
                Cell(
                    ws.Cells[curRow, 14],
                    elem.LevelUpPiece > 0 ? elem.LevelUpPiece : (object)"—",
                    null
                );
                Cell(
                    ws.Cells[curRow, 15],
                    elem.LevelDownPiece > 0 ? elem.LevelDownPiece : (object)"—",
                    null
                );
                Cell(
                    ws.Cells[curRow, 16],
                    elem.SpreadAuto == 1 ? "✓自动" : "手动",
                    elem.SpreadAuto == 1 ? "D9EAD3" : null
                );
                Cell(ws.Cells[curRow, 17], mergedDisplay, mergedHex);
                curRow++;
            }

            curRow++; // 链间空行
        }

        ws.Column(1).Width = 12;
        ws.Column(2).Width = 14;
        ws.Column(3).Width = 10;
        ws.Column(4).Width = 16;
        ws.Column(5).Width = 12;
        ws.Column(6).Width = 8;
        ws.Column(7).Width = 12;
        ws.Column(8).Width = 10;
        ws.Column(9).Width = 10;
        ws.Column(10).Width = 12;
        ws.Column(11).Width = 10;
        ws.Column(12).Width = 18;
        ws.Column(13).Width = 10;
        ws.Column(14).Width = 10;
        ws.Column(15).Width = 10;
        ws.Column(16).Width = 10;
        ws.Column(17).Width = 14;
        ws.Row(1).Height = 72;
        ws.Cells[1, 1].Style.WrapText = true;
    }

    // ── Sheet 5：生成器权重算法说明 ─────────────────────────────────────────────
    private static void BuildWeightAlgoSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("权重算法");
        ws.View.FreezePanes(3, 1);

        // ── 算法总览说明（第1行）──
        MechNote(
            ws,
            1,
            1,
            "【生成器↔订单闭环机制】每次生成器触发，用 ListWeightSelectOne 从产出链中随机抽1个元素。"
                + "抽取权重 = 基础权重(WeightType) × 棋盘容量衰减系数 × 订单加成倍率(chainMultiple)。"
                + "订单系统同时用 RepeatWeightDecrease 对已连续出现的合成线降权，形成闭环：「棋盘满了→权重降→出高级」「订单需要→权重升→定向产出」。",
            18
        );

        // ── 算法步骤表（第2行起）──
        var titleCell = ws.Cells[2, 1, 2, 18];
        titleCell.Merge = true;
        titleCell.Value = "▶ CurWeight 计算公式（每次触发时实时计算）";
        titleCell.Style.Font.Bold = true;
        titleCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        titleCell.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
        titleCell.Style.Font.Color.SetColor(Color.White);
        Border(titleCell);

        // 公式步骤
        var steps = new[]
        {
            ("步骤", "公式/参数", "来源", "说明", "配置字段", "hex"),
            (
                "① 基础权重",
                "BaseW = WeightType枚举值映射",
                "ItemModelConfig_orange",
                "WeightType 1→高频 / 2→普通 / 3→低频 / 4→极低(稀有)。实际映射值在游戏代码中，配置只存档位。",
                "Spread_WeightType",
                "D9EAD3"
            ),
            (
                "② 容量衰减",
                "CapFactor = 1 - (当前棋盘数量 / ItemMax)",
                "实时棋盘状态",
                "格子或仓库接近上限时该元素权重自动降低。当棋盘数达到 ItemMax 时 CapFactor=0，该元素从池中移除。",
                "Spread_ItemMaxNumber / Spread_StorageMaxNumber",
                "FFF2CC"
            ),
            (
                "③ 订单加成",
                "ChainMultiple = 合成线倍率",
                "OrderRandomHelper",
                "当该元素所在合成线出现在当前订单时，系统对该线所有元素权重×倍率(chainMultipleExistMinLevelMap)。订单需要→权重高→更常出。",
                "itemChainMultiple / chainMultipleRequiredState",
                "EBF3FB"
            ),
            (
                "④ 重复衰减",
                "RepeatFactor = repeatWeightDecrease 倍率",
                "TestOrderLogInfo",
                "近期已经频繁产出了某条合成线的元素后，mapRepeatWeightDecreaseInfo 记录下来，对该线权重连续打折，避免同类元素刷屏。",
                "mapRepeatWeightDecrease / RepeatWeightDecrease",
                "FCE5CD"
            ),
            (
                "⑤ 最终权重",
                "CurWeight = BaseW × CapFactor × ChainMultiple × RepeatFactor",
                "所有来源汇总",
                "ListWeightSelectOne(所有候选元素, CurWeight) → 按比例随机选1个。CurWeight=0 的元素不参与抽取。",
                "—",
                "E8DAEF"
            ),
        };

        var headerRow = 3;
        string[] hdrCols = ["步骤", "公式/参数", "来源", "说明", "关联配置字段"];
        int[] hdrWidths = [14, 36, 22, 58, 42];
        for (var j = 0; j < hdrCols.Length; j++)
            Header(ws.Cells[headerRow, j + 1], hdrCols[j], "2F5496");

        for (var i = 1; i < steps.Length; i++)
        {
            var (step, formula, source, desc, field, hex) = steps[i];
            var r = headerRow + i;
            Cell(ws.Cells[r, 1], step, hex);
            Cell(ws.Cells[r, 2], formula, hex);
            Cell(ws.Cells[r, 3], source, "F5F5F5");
            Cell(ws.Cells[r, 4], desc, "F5F5F5");
            Cell(ws.Cells[r, 5], field, "F5F5F5");
            ws.Cells[r, 4].Style.WrapText = true;
            ws.Row(r).Height = 40;
        }

        // ── 参数速查表：6条链全量数据 ──
        var paramTitle = ws.Cells[
            headerRow + steps.Length + 1,
            1,
            headerRow + steps.Length + 1,
            18
        ];
        paramTitle.Merge = true;
        paramTitle.Value = "▶ 产出链参数速查（CurWeight 计算所需输入参数）";
        paramTitle.Style.Font.Bold = true;
        paramTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        paramTitle.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
        paramTitle.Style.Font.Color.SetColor(Color.White);
        Border(paramTitle);

        string[] paramHeaders =
        [
            "链",
            "等级",
            "Type",
            "中文名",
            "WeightType\n(①基础权重档位)",
            "ItemMax\n(②棋盘容量上限)",
            "StorageMax\n(②仓库容量上限)",
            "格子CD(s)\n(产出间隔)",
            "仓库CD(s)",
            "BubbleChance%\n(气泡掉落概率)",
            "LevelUpPiece\n(升级所需数)",
            "加速钻石消耗",
            "GenChain1\n(主力生成器ID)",
            "GenChain2\n(辅助生成器ID)",
            "售价",
            "合并→",
        ];
        var phr = headerRow + steps.Length + 2;
        for (var j = 0; j < paramHeaders.Length; j++)
        {
            Header(ws.Cells[phr, j + 1], paramHeaders[j], j < 4 ? "1F4E79" : "2F5496");
            ws.Cells[phr, j + 1].Style.WrapText = true;
        }
        ws.Row(phr).Height = 42;

        var chains = LoadMainGenerators();
        var chainColors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var pr = phr + 1;
        var ci = 0;

        foreach (var (headType, chain) in chains.OrderBy(kv => int.Parse(kv.Key)))
        {
            var baseHex = chainColors[ci % chainColors.Length];
            ci++;
            for (var step = 0; step < chain.Count; step++)
            {
                var elem = chain[step];
                var isHead = step == 0;
                var isTail = step == chain.Count - 1;
                var rowHex = isHead
                    ? "D9EAD3"
                    : isTail
                        ? "FADADD"
                        : baseHex;
                var cn = ItemCn(elem.Type);
                var speedUp =
                    elem.SpreadItemSpeedUpCost > 0 || elem.SpreadStorageSpeedUpCost > 0
                        ? $"格:{elem.SpreadItemSpeedUpCost} 仓:{elem.SpreadStorageSpeedUpCost}"
                        : "—";

                Cell(ws.Cells[pr, 1], isHead ? headType : "", isHead ? "D9EAD3" : "F5F5F5");
                Cell(
                    ws.Cells[pr, 2],
                    isHead
                        ? "链头"
                        : isTail
                            ? $"Lv.{step + 1}(末)"
                            : $"Lv.{step + 1}",
                    "F5F5F5"
                );
                Cell(ws.Cells[pr, 3], elem.Type, rowHex);
                Cell(ws.Cells[pr, 4], cn.Length > 0 ? cn : "—", rowHex);

                // WeightType — 用颜色区分档位
                var wtHex = elem.SpreadWeightType switch
                {
                    1 => "D9EAD3", // 高频 绿
                    2 => "FFF2CC", // 普通 黄
                    3 => "FCE5CD", // 低频 橙
                    4 => "FADADD", // 极低 红
                    _ => null,
                };
                var wtLabel = elem.SpreadWeightType switch
                {
                    1 => "1-高频",
                    2 => "2-普通",
                    3 => "3-低频",
                    4 => "4-稀有",
                    _ => "—",
                };
                Cell(ws.Cells[pr, 5], wtLabel, wtHex);

                Cell(
                    ws.Cells[pr, 6],
                    elem.SpreadItemMax > 0 ? elem.SpreadItemMax : (object)"—",
                    elem.SpreadItemMax > 0 ? "EBF3FB" : null
                );
                Cell(
                    ws.Cells[pr, 7],
                    elem.SpreadStorageMax > 0 ? elem.SpreadStorageMax : (object)"—",
                    elem.SpreadStorageMax > 0 ? "EBF3FB" : null
                );
                Cell(
                    ws.Cells[pr, 8],
                    elem.SpreadItemRecovery > 0 ? elem.SpreadItemRecovery : (object)"—",
                    elem.SpreadItemRecovery > 0 ? "FFF2CC" : null
                );
                Cell(
                    ws.Cells[pr, 9],
                    elem.SpreadStorageRecovery > 0 ? elem.SpreadStorageRecovery : (object)"—",
                    elem.SpreadStorageRecovery > 0 ? "FFF2CC" : null
                );
                Cell(
                    ws.Cells[pr, 10],
                    elem.BubbleChance > 0 ? $"{elem.BubbleChance}%" : (object)"—",
                    elem.BubbleChance > 0 ? "D9EAD3" : null
                );
                Cell(
                    ws.Cells[pr, 11],
                    elem.LevelUpPiece > 0 ? elem.LevelUpPiece : (object)"—",
                    null
                );
                Cell(ws.Cells[pr, 12], speedUp, speedUp != "—" ? "FCE5CD" : null);
                Cell(
                    ws.Cells[pr, 13],
                    elem.GenChain1.Length > 0 ? elem.GenChain1 : "—",
                    elem.GenChain1.Length > 0 ? "D9EAD3" : null
                );
                Cell(
                    ws.Cells[pr, 14],
                    elem.GenChain2.Length > 0 ? elem.GenChain2 : "—",
                    elem.GenChain2.Length > 0 ? "FFF2CC" : null
                );
                Cell(
                    ws.Cells[pr, 15],
                    elem.SellingPrice > 0 ? elem.SellingPrice : (object)"—",
                    elem.SellingPrice > 0 ? "FFF2CC" : null
                );
                var mergedDisplay = elem.MergedType.Length > 0 ? elem.MergedType : "（终点）";
                Cell(
                    ws.Cells[pr, 16],
                    mergedDisplay,
                    elem.MergedType.Length > 0 ? "EBF3FB" : "F0F0F0"
                );
                pr++;
            }
            pr++; // 链间空行
        }

        // 统计行
        var total = chains.Values.Sum(c => c.Count);
        ws.Cells[pr, 1].Value = $"合计 {chains.Count} 条链 / {total} 个元素";
        ws.Cells[pr, 1].Style.Font.Bold = true;

        // 列宽
        int[] colWidths = [10, 12, 10, 16, 14, 12, 12, 12, 10, 14, 12, 20, 14, 14, 10, 12];
        for (var j = 0; j < colWidths.Length; j++)
            ws.Column(j + 1).Width = colWidths[j];
        ws.Row(1).Height = 56;
        ws.Cells[1, 1].Style.WrapText = true;
        for (var j = 0; j < hdrCols.Length; j++)
            ws.Column(j + 1).Width = Math.Max(ws.Column(j + 1).Width, hdrWidths[j]);
    }
}
