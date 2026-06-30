using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// Merge Cooking 竞品核心循环分析 xlsx 生成器。
/// 数据源：MuMu模拟器 ADB → shared_prefs playerprefs.xml（URL编码JSON，~1.5MB）
/// 技术栈：IL2CPP Unity，无Lua，本地数据存 SharedPreferences
/// 输出：竞品-MergeCooking核心循环分析.xlsx → Documents\workspace\
/// </summary>
public static class MergeCookingWriter
{
    private const string OutFileName = "竞品-MergeCooking核心循环分析.xlsx";

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
        if (hex is not null)
        {
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
        }
        if (wrap)
            c.Style.WrapText = true;
        Border(c);
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

    // ── 内嵌数据（ADB playerprefs.xml + UnityPy 资产解包 + 语言包提取）─────────

    // 生成器：goodsID → (名称, surplus容量, 历史产出统计)
    // 名称来源：configlanguagefolder UnityPy解包 → lang_ui_en.json（G_vegetable / G_eggmeat / G_drink系列）
    private static readonly GeneratorDef[] Generators =
    [
        new(
            100005,
            "Vegetable Generator A (G_vegetable series)",
            30,
            [("100020 基础蔬菜", 50.0), ("100022 Cabbage/C_leaf", 33.3), ("100040 高级蔬菜", 16.7)]
        ),
        new(100007, "Vegetable Generator B (G_vegetable series)", 32, [("100020 基础蔬菜", 100.0)]),
        new(
            100103,
            "Protein Generator A (G_eggmeat: Pork/Lamb/Poultry/Eggs)",
            24,
            [("100120 面粉初级", 25.0), ("100140 面粉中级", 75.0)]
        ),
        new(
            100105,
            "Protein Generator B (G_eggmeat series)",
            36,
            [
                ("100120 面粉初级", 23.1),
                ("100140 面粉中级", 53.8),
                ("100141 面粉高级", 15.4),
                ("100142 面粉精品", 7.7)
            ]
        ),
        new(100183, "水产机A (Seafood chain)", 60, [("（surplus未清空，无产出记录）", 0.0)]),
        new(100184, "水产机B (Seafood chain)", 60, [("（surplus未清空，无产出记录）", 0.0)]),
        new(
            100204,
            "Drink Generator / 奶制品机 (G_drink: Soft Drinks+Additives+Coffee)",
            30,
            [("100220 C_milk Frothed Milk/Cheese/Butter链", 100.0)]
        ),
        new(100293, "综合机A (multi-chain)", 50, [("（surplus未清空，无产出记录）", 0.0)]),
        new(101554, "Tourist Baggage Lv1 (G_america theme generator)", 24, [("（surplus未清空）", 0.0)]),
        new(
            101555,
            "Tourist Baggage Lv2 (G_america series)",
            30,
            [("101580 主题元素L1", 50.0), ("101581 主题元素L2", 27.8), ("101611 特殊元素", 22.2)]
        ),
        new(
            101556,
            "Tourist Baggage Lv3 (G_america series)",
            36,
            [
                ("101580 主题元素L1", 55.6),
                ("101581 主题元素L2", 11.1),
                ("101582 主题元素L3", 11.1),
                ("101611 特殊元素", 22.2)
            ]
        ),
        new(
            101557,
            "Tourist Baggage Lv4 (G_america series)",
            42,
            [("101580 主题元素L1", 50.0), ("101611 特殊元素", 50.0)]
        ),
        new(101655, "Tourist Baggage Lv5 (G_america MAX)", 50, [("（surplus未清空）", 0.0)]),
    ];

    private record GeneratorDef(
        int GoodsId,
        string Name,
        int Surplus,
        (string Item, double Pct)[] ProduceProbs
    );

    // 合并链（语言包 lang_ui_en.json 真实英文名 — 置信度：高）
    // 链长度来源：lang_ui_en.json 枚举计数（19806条）
    // C_leaf=10步, C_root=12步, C_meat=12步, C_egg=10步, C_milk=9步
    // C_softdrink=13步, C_coffee=17步, C_glass=7步, C_coconutmilk=7步, C_coconutshell=7步
    private static readonly ChainDef[] MergeChains =
    [
        new(
            "A-Leaf Vegetables (C_leaf, 10级)",
            "G_vegetable (100005/100007) 叶菜产出链",
            [
                ("100022", "Lettuce → Cabbage (C_leaf Lv1→2)", 1, true),
                ("100029", "Broccoli (C_leaf Lv3)", 3, false),
                ("100025", "Red Cabbage (C_leaf Lv4)", 4, true),
                ("100031", "Arugula (C_leaf Lv5)", 5, true),
                ("100032", "Spinach (C_leaf Lv6)", 6, true),
                ("—", "Beet (C_leaf Lv7)", 7, false),
                ("—", "Button Mushroom (C_leaf Lv8)", 8, false),
                ("—", "Cauliflower (C_leaf Lv9)", 9, true),
                ("—", "Artichoke (C_leaf Lv10 MAX)", 10, false),
            ]
        ),
        new(
            "B-Root Vegetables (C_root, 12级)",
            "G_vegetable (100005/100007) 根菜产出链",
            [
                ("100020", "Tomato (C_root Lv1)", 1, true),
                ("100021", "Cucumber (C_root Lv2)", 2, false),
                ("100023", "Asparagus (C_root Lv3)", 3, false),
                ("100024", "Carrot (C_root Lv4)", 3, false),
                ("100026", "Onion (C_root Lv5)", 4, true),
                ("100027", "Bell Pepper (C_root Lv6)", 5, true),
                ("100028", "Eggplant (C_root Lv7)", 5, true),
                ("—", "Potato (C_root Lv8)", 6, false),
                ("—", "Okra (C_root Lv9)", 7, false),
                ("—", "Cherry Tomato (C_root Lv10)", 8, false),
                ("—", "Pea (C_root Lv11)", 9, false),
                ("—", "Corn (C_root Lv12 MAX)", 10, false),
            ]
        ),
        new(
            "C-Meats (C_meat, 12级)",
            "G_eggmeat (100103/100105) 肉类产出链",
            [
                ("100120", "Bacon (C_meat Lv1)", 1, true),
                ("100121", "Sausage (C_meat Lv2)", 2, false),
                ("100122", "Ham (C_meat Lv3)", 2, true),
                ("100123", "Patty (C_meat Lv4)", 3, true),
                ("100124", "Pork Ribs (C_meat Lv5)", 4, true),
                ("—", "Pork Knuckle (C_meat Lv6)", 5, false),
                ("—", "Lamb Chop (C_meat Lv7)", 6, false),
                ("—", "Leg of Lamb (C_meat Lv8)", 7, false),
                ("—", "Suckling Pig (C_meat Lv9)", 8, false),
                ("—", "Whole Lamb (C_meat Lv10)", 9, false),
                ("—", "Beef (C_meat Lv11)", 10, false),
                ("—", "Lamb (C_meat Lv12 MAX)", 11, false),
            ]
        ),
        new(
            "D-Poultry/Eggs (C_egg, 10级)",
            "G_eggmeat (100103/100105) 禽蛋产出链",
            [
                ("100140", "Egg (C_egg Lv1)", 1, true),
                ("100141", "Chicken Wing (C_egg Lv2)", 2, true),
                ("100142", "Chicken Drumstick (C_egg Lv3)", 3, true),
                ("100143", "Chicken Breast (C_egg Lv4)", 4, false),
                ("100144", "Whole Chicken (C_egg Lv5)", 5, true),
                ("—", "Duck Breast (C_egg Lv6)", 6, false),
                ("—", "Whole Duck (C_egg Lv7)", 7, false),
                ("—", "Goose (C_egg Lv8)", 8, false),
                ("—", "Quail (C_egg Lv9)", 9, false),
                ("100145", "Turkey (C_egg Lv10 MAX)", 10, true),
            ]
        ),
        new(
            "E-Dairy (C_milk, 9级)",
            "G_drink/100204 奶制品产出链",
            [
                ("100220", "Frothed Milk (C_milk Lv1)", 1, true),
                ("100221", "Whipped Cream (C_milk Lv2)", 2, true),
                ("100222", "Cheese (C_milk Lv3)", 3, true),
                ("100223", "Butter (C_milk Lv4)", 4, true),
                ("100224", "Cream Cheese (C_milk Lv5)", 5, true),
                ("—", "Mozzarella Cheese (C_milk Lv6)", 6, false),
                ("—", "Feta (C_milk Lv7)", 7, false),
                ("—", "Cheddar (C_milk Lv8)", 8, false),
                ("—", "Parmesan (C_milk Lv9 MAX)", 9, false),
            ]
        ),
        new(
            "F-Soft Drinks (C_softdrink, 13级)",
            "G_drink (100204) 软饮产出链",
            [
                ("100036", "Water (C_softdrink Lv1)", 1, true),
                ("100041", "Ice Cube (C_softdrink Lv2)", 2, true),
                ("100044", "Sparkling Water (C_softdrink Lv3)", 3, false),
                ("100050", "Cola (C_softdrink Lv4)", 4, true),
                ("100055", "Soda (C_softdrink Lv5)", 5, true),
                ("100056", "Orange Soda (C_softdrink Lv6)", 6, true),
                ("—", "Grape Soda (C_softdrink Lv7)", 7, false),
                ("—", "Black Tea (C_softdrink Lv8)", 8, false),
                ("—", "Green Tea (C_softdrink Lv9)", 9, false),
                ("—", "Lactic Acid Drink (C_softdrink Lv10)", 10, false),
                ("—", "Bubble Tea (C_softdrink Lv11)", 11, false),
                ("—", "Floral Tea (C_softdrink Lv12)", 12, false),
                ("—", "Energy Drink (C_softdrink Lv13 MAX)", 13, false),
            ]
        ),
        new(
            "G-Coffee (C_coffee, 17级)",
            "G_drink (100204) 咖啡产出链",
            [
                ("101100", "Coffee Bean (C_coffee Lv1)", 1, true),
                ("101110", "Ground Coffee (C_coffee Lv2)", 2, true),
                ("—", "Cup of Coffee (C_coffee Lv3)", 3, false),
                ("—", "Caramel Macchiato (C_coffee Lv4)", 4, false),
                ("—", "Latte (C_coffee Lv5)", 5, false),
                ("—", "Cappuccino (C_coffee Lv6)", 6, false),
                ("—", "Mocha (C_coffee Lv7)", 7, false),
                ("—", "Frappe (C_coffee Lv8)", 8, false),
                ("—", "Mocha Frappe (C_coffee Lv9)", 9, false),
                ("—", "Caramel Frappe (C_coffee Lv10)", 10, false),
                ("—", "Mug of Coffee (C_coffee Lv11)", 11, false),
                ("—", "Coffee on the Go (C_coffee Lv12)", 12, false),
                ("—", "Cold Brew (C_coffee Lv13)", 13, false),
                ("—", "Pour-over Coffee (C_coffee Lv14)", 14, false),
                ("—", "French Press Coffee (C_coffee Lv15)", 15, false),
                ("—", "Moka Pot Coffee (C_coffee Lv16)", 16, false),
                ("—", "Siphon Coffee (C_coffee Lv17 MAX)", 17, false),
            ]
        ),
        new(
            "H-Glassware (C_glass, 7级)",
            "G_french (Admission Ticket) 玻璃餐具链",
            [
                ("100500", "Broken Glass (C_glass Lv1)", 1, true),
                ("100501", "Glass Bowl (C_glass Lv2)", 2, false),
                ("100502", "Crystal Glass (C_glass Lv3)", 3, true),
                ("100503", "Glass Plate (C_glass Lv4)", 4, true),
                ("100504", "Glass Jar (C_glass Lv5)", 5, false),
                ("100510", "Glass Cake Stand (C_glass Lv6)", 6, true),
                ("100511", "Glass Dessert Stand (C_glass Lv7 MAX)", 7, true),
            ]
        ),
        new(
            "I-Coconut Tableware (C_coconutshell, 7级)",
            "G_america/G_french 椰壳餐具链",
            [
                ("100512", "Coconut Shell (C_coconutshell Lv1)", 1, true),
                ("100513", "Coconut Spoon (C_coconutshell Lv2)", 2, false),
                ("100520", "Coconut Fork (C_coconutshell Lv3)", 3, true),
                ("100521", "Coconut Saucer (C_coconutshell Lv4)", 4, true),
                ("100522", "Coconut Bowl (C_coconutshell Lv5)", 5, false),
                ("—", "Coconut Teapot (C_coconutshell Lv6)", 6, false),
                ("—", "Coconut Teacup (C_coconutshell Lv7 MAX)", 7, false),
            ]
        ),
        new(
            "J-Coconut Beverages (C_coconutmilk, 7级)",
            "G_america 椰汁饮品链",
            [
                ("—", "Glass of Coconut Water (C_coconutmilk Lv1)", 1, false),
                ("—", "Coconut Sago Soup (C_coconutmilk Lv2)", 2, false),
                ("—", "Coconut Soda (C_coconutmilk Lv3)", 3, false),
                ("—", "Coconut Yogurt (C_coconutmilk Lv4)", 4, false),
                ("—", "Coconut Jelly (C_coconutmilk Lv5)", 5, false),
                ("—", "Coconut Pudding (C_coconutmilk Lv6)", 6, false),
                ("—", "Coconut Candy (C_coconutmilk Lv7 MAX)", 7, false),
            ]
        ),
        new(
            "K-Theme / Tourist Baggage (G_america)",
            "101554-101657 主题装饰机产出链",
            [
                ("101554", "Tourist Baggage Lv1 (G_america1, GEN)", 1, false),
                ("101555", "Tourist Baggage Lv2 (G_america2, GEN)", 2, true),
                ("101556", "Tourist Baggage Lv3 (G_america3, GEN)", 3, true),
                ("101557", "Tourist Baggage Lv4 (G_america4, GEN)", 4, false),
                ("101655", "Tourist Baggage Lv5 (G_america5 MAX, GEN)", 5, false),
                ("101580", "Theme Element L1 (c_americagift1: Bald Eagle Stamp)", 1, true),
                ("101581", "Theme Element L2 (c_americagift2: Hollywood Postcard)", 2, true),
                ("101582", "Theme Element L3 (c_americagift3: Cowboy Hat)", 3, true),
                ("101583", "Theme Element L4 (c_americagift4: American Football)", 4, true),
                ("101584", "Theme Element L5 (c_americagift5: I Love MC Mug)", 5, false),
                ("101586", "Theme Element L6 (c_americagift6: Fridge Magnet)", 6, false),
            ]
        ),
    ];

    private record ChainDef(
        string ChainId,
        string Source,
        (string Id, string Name, int Level, bool OrderUsed)[] Items
    );

    // 动态订单样本（从 GameOrdersModel_KEY 提取，mainOrderID=-1 表示全动态刷新）
    private static readonly OrderSample[] OrderSamples =
    [
        new(1, "waveGroup=6", 8, [(100042, 1), (100018, 1)], [(1002, 64), (1001, 93)], "groupDef"),
        new(2, "waveGroup=6", 8, [(100022, 1), (100039, 1)], [(1002, 65), (1001, 39)], "groupDef"),
        new(3, "waveGroup=6", 8, [(100012, 1), (100050, 1)], [(1002, 29), (1001, 60)], "groupDef"),
        new(4, "waveGroup=6", 8, [(100106, 1), (100058, 1)], [(1002, 38), (1001, 37)], "groupDef"),
        new(5, "waveGroup=6", 8, [(100079, 1), (100069, 1)], [(1002, 34), (1001, 27)], "groupDef"),
    ];

    private record OrderSample(
        int Idx,
        string WaveDesc,
        int TotalWave,
        (int ItemId, int Num)[] Requirements,
        (int RewardId, int Num)[] Rewards,
        string GroupName
    );

    // 箱子掉落样本
    private static readonly BoxDrop[] BoxDrops =
    [
        new(100758, 9, [(100520, 5), (100521, 1)]),
        new(100758, 11, [(100511, 4), (100512, 1), (100513, 1)]),
        new(100620, 59, [(100510, 4), (100511, 1), (100512, 1)]),
        new(100756, 9, [(100520, 6), (100521, 2)]),
    ];

    private record BoxDrop(int BoxId, int Seq, (int ItemId, int Count)[] Drops);

    // ── 公开入口 ──────────────────────────────────────────────────────────────

    public static void Run(string? outputDir = null)
    {
        var dir = outputDir ?? OutputPaths.Reports;
        var outPath = Path.Combine(dir, OutFileName);

        using var pkg = new ExcelPackage();

        BuildGeneratorSheet(pkg);
        BuildChainSheet(pkg);
        BuildOrderSheet(pkg);
        BuildActivitySheet(pkg);
        BuildSummarySheet(pkg);

        pkg.SaveAs(new FileInfo(outPath));
        Console.WriteLine($"[MergeCooking] 已生成：{outPath}");
        OutputPaths.GitCommit($"[MergeCooking] 更新竞品分析报告 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── Sheet 1：生成器配置 ───────────────────────────────────────────────────

    private static void BuildGeneratorSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("生成器配置");
        ws.View.FreezePanes(3, 1);

        MechNote(
            ws,
            1,
            1,
            "【Merge Cooking 生成器机制 — 置信度：高（ADB运行时数据 + UnityPy语言包双重验证）】"
                + "生成器类型：GoodsState=6（激活中），GoodsState=8（CD冷却中）。"
                + "CookingType=1（全部相同，无区分generator vs item，均为time驱动）。"
                + "InitiativeSurplusNumber = 当前积压产出数量（玩家未点击收集）。"
                + "产出概率来自 InitiativeProduceList 历史记录统计（多个generator产出均摊推算）。"
                + "生成器英文名来源：lang_ui_en.json UnityPy解包（configlanguagefolder.unity3d，XOR key无需，直接UnityFS头）。"
                + "G_vegetable=Vegetable Generator(Lv1~10), G_eggmeat=Protein Generator(Pork/Lamb/Poultry/Eggs), "
                + "G_drink=Drink Generator(SoftDrinks+Additives+Coffee), G_america=Tourist Baggage(Lv1~7合并升级)。"
                + "global-metadata.dat 已通过4字节循环XOR(key=0x75 0xE5 0xE6 0xE9)解密，magic验证通过但字符串池二次混淆。",
            14
        );

        string[] headers = ["生成器ID", "推断名称", "当前surplus", "产出项目", "产出概率%", "置信度", "备注",];
        string[] hdrHex = ["1F4E79", "2F5496", "2F5496", "1F4E79", "2F5496", "595959", "595959"];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var row = 3;
        var colors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var ci = 0;

        foreach (var gen in Generators)
        {
            var hex = colors[ci % colors.Length];
            ci++;
            var validProbs = gen.ProduceProbs.Where(p => p.Item2 > 0).ToArray();
            var rowSpan = Math.Max(1, validProbs.Length);

            var conf =
                validProbs.Length > 0 && validProbs.Any(p => p.Item2 < 99)
                    ? "高（历史记录统计）"
                    : validProbs.Length > 0
                        ? "高（100%单一产出）"
                        : "中（surplus积压未触发）";
            var confHex = conf.StartsWith("高")
                ? "D9EAD3"
                : conf.StartsWith("中")
                    ? "FFF2CC"
                    : "FCE5CD";

            if (rowSpan > 1)
            {
                ws.Cells[row, 1, row + rowSpan - 1, 1].Merge = true;
                ws.Cells[row, 2, row + rowSpan - 1, 2].Merge = true;
                ws.Cells[row, 3, row + rowSpan - 1, 3].Merge = true;
                ws.Cells[row, 6, row + rowSpan - 1, 6].Merge = true;
                ws.Cells[row, 7, row + rowSpan - 1, 7].Merge = true;
            }

            Cell(ws.Cells[row, 1], gen.GoodsId, hex);
            Cell(ws.Cells[row, 2], gen.Name, hex);
            Cell(ws.Cells[row, 3], gen.Surplus, hex);
            Cell(ws.Cells[row, 6], conf, confHex);
            Cell(
                ws.Cells[row, 7],
                gen.GoodsId >= 101554
                    ? "主题/限时活动专属机器"
                    : gen.GoodsId >= 100100
                        ? "食材加工类生成器"
                        : "基础食材生成器",
                "F5F5F5"
            );

            for (var k = 0; k < rowSpan; k++)
            {
                var prob = k < validProbs.Length ? validProbs[k] : ("—", 0.0);
                var probPct = prob.Item2;
                var probItem = prob.Item1;
                var probHex =
                    probPct >= 70
                        ? "D9EAD3"
                        : probPct >= 30
                            ? "FFF2CC"
                            : probPct > 0
                                ? "FCE5CD"
                                : "F0F0F0";
                Cell(ws.Cells[row + k, 4], probItem, k == 0 ? hex : "FAFAFA");
                Cell(
                    ws.Cells[row + k, 5],
                    probPct > 0 ? (object)Math.Round(probPct, 1) : "N/A",
                    probHex
                );
                Border(ws.Cells[row + k, 1]);
                Border(ws.Cells[row + k, 2]);
                Border(ws.Cells[row + k, 3]);
            }

            row += rowSpan;
        }

        ws.Column(1).Width = 13;
        ws.Column(2).Width = 20;
        ws.Column(3).Width = 14;
        ws.Column(4).Width = 26;
        ws.Column(5).Width = 14;
        ws.Column(6).Width = 20;
        ws.Column(7).Width = 22;
        ws.Row(1).Height = 70;
        ws.Row(2).Height = 22;

        ws.Cells[row + 1, 1].Value =
            $"合计 {Generators.Length} 个生成器  |  数据来源：ADB GAME_LEVEL_MAP_DATA + lang_ui_en.json（UnityPy解包，G_vegetable/G_eggmeat/G_drink/G_america系列英文名）";
        ws.Cells[row + 1, 1].Style.Font.Bold = true;
    }

    // ── Sheet 2：元素合并链 ────────────────────────────────────────────────────

    private static void BuildChainSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("元素合并链");
        ws.View.FreezePanes(3, 1);

        MechNote(
            ws,
            1,
            1,
            "【Merge Cooking 合并链机制 — 置信度：高（UnityPy解包 lang_ui_en.json 19806条确认）】"
                + "合并规则：2个相同等级元素 → 合并为+1级元素（标准 2-to-1 merge）。"
                + "链长度（精确，来自lang_ui_en枚举）："
                + "C_leaf=10级(Lettuce→Artichoke), C_root=12级(Tomato→Corn), "
                + "C_meat=12级(Bacon→Lamb), C_egg=10级(Egg→Turkey), "
                + "C_milk=9级(Frothed Milk→Parmesan), C_softdrink=13级(Water→Energy Drink), "
                + "C_coffee=17级(Coffee Bean→Siphon Coffee), C_glass=7级(Broken Glass→Glass Dessert Stand), "
                + "C_coconutshell=7级(Coconut Shell→Coconut Teacup), C_coconutmilk=7级(Glass of Coconut Water→Coconut Candy)。"
                + "G_america主题机：Lv1~7均为'Tourist Baggage'合并升级，产出c_americagift1~10礼品元素。"
                + "订单消耗列仅基于5个当前动态订单样本，非全量分布。",
            14
        );

        string[] headers = ["链名", "来源", "元素ID", "推断名称", "链内等级", "订单消耗", "备注"];
        string[] hdrHex = ["1F4E79", "595959", "2F5496", "2F5496", "595959", "1F4E79", "595959"];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var chainColors = new[] { "D9EAD3", "FFF2CC", "EBF3FB", "FCE5CD", "E8DAEF", "F3F3F3" };
        var row = 3;
        var ci = 0;

        foreach (var chain in MergeChains)
        {
            var hex = chainColors[ci % chainColors.Length];
            ci++;

            var title = ws.Cells[row, 1, row, 7];
            title.Merge = true;
            title.Value = $"▶ {chain.ChainId}  共 {chain.Items.Length} 种元素  |  来源：{chain.Source}";
            title.Style.Font.Bold = true;
            title.Style.Fill.PatternType = ExcelFillStyle.Solid;
            title.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
            title.Style.Font.Color.SetColor(Color.White);
            Border(title);
            row++;

            foreach (var (id, name, level, orderUsed) in chain.Items)
            {
                var rowHex =
                    level == 1
                        ? "EBF3FB"
                        : level >= 6
                            ? "FADADD"
                            : hex;
                Cell(ws.Cells[row, 1], chain.ChainId, "F5F5F5");
                Cell(ws.Cells[row, 2], chain.Source, "F5F5F5", wrap: true);
                Cell(ws.Cells[row, 3], id, rowHex);
                Cell(ws.Cells[row, 4], name, rowHex);
                Cell(
                    ws.Cells[row, 5],
                    $"L{level}",
                    level <= 2
                        ? "D9EAD3"
                        : level >= 5
                            ? "FADADD"
                            : "FFF2CC"
                );
                Cell(ws.Cells[row, 6], orderUsed ? "是" : "", orderUsed ? "D9EAD3" : null);
                Cell(
                    ws.Cells[row, 7],
                    name.Contains("GEN")
                        ? "该元素本身是生成器"
                        : orderUsed
                            ? "订单需求元素"
                            : "",
                    "F5F5F5"
                );
                row++;
            }
            row++;
        }

        ws.Column(1).Width = 16;
        ws.Column(2).Width = 20;
        ws.Column(3).Width = 12;
        ws.Column(4).Width = 20;
        ws.Column(5).Width = 10;
        ws.Column(6).Width = 10;
        ws.Column(7).Width = 22;
        ws.Row(1).Height = 70;
        ws.Row(2).Height = 22;

        ws.Cells[row + 1, 1].Value =
            $"合计 {MergeChains.Length} 条合并链  |  数据来源：GAME_BASIC_UNLOCK_GOODS_DATA + GAME_LEVEL_MAP_DATA + lang_ui_en.json（UnityPy解包，真实英文名）";
        ws.Cells[row + 1, 1].Style.Font.Bold = true;
    }

    // ── Sheet 3：订单配置 ─────────────────────────────────────────────────────

    private static void BuildOrderSheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("订单配置");
        ws.View.FreezePanes(3, 1);

        MechNote(
            ws,
            1,
            1,
            "【Merge Cooking 订单机制 — 置信度：高（直接读取 GameOrdersModel_KEY）】"
                + "关键发现：mDynamicOrdersOpen=true，全部5个当前订单的 mainOrderID=-1（动态随机，非固定剧情）。"
                + "Wave系统：当前 waveGroupID=6，totalWave=8（一轮波次共8个订单），virtualWave=6（第6波）。"
                + "每个订单需要2种元素各1件（双元素需求设计，不像 TravelTown 单一需求）。"
                + "奖励类型：type=1,id=1001=金币；type=1,id=1002=能量。订单同时给金币+能量。"
                + "RestaurantID=3（第3个餐厅），玩家当前level=7。"
                + "对比 Gossip Harbor：Harbor 动态权重随机；Merge Cooking 是波次组(wave group)随机——"
                + "先确定波次组，波次组内按inWaveOrderIndex顺序提供订单，类似分段随机。"
                + "主线订单ID 10301~10806 为分餐厅线性推进（区别于wave动态订单）。",
            14
        );

        string[] headers = ["订单序号", "波次描述", "波次总数", "需求元素1", "需求元素2", "奖励能量", "奖励金币", "来源类型",];
        string[] hdrHex =
        [
            "1F4E79",
            "595959",
            "2F5496",
            "1F4E79",
            "1F4E79",
            "2F5496",
            "2F5496",
            "595959",
        ];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var row = 3;
        var colors = new[] { "EBF3FB", "FFF2CC", "D9EAD3", "FCE5CD" };
        var ci = 0;

        foreach (var order in OrderSamples)
        {
            var hex = colors[ci % colors.Length];
            ci++;

            (int ItemId, int Num) req1 =
                order.Requirements.Length > 0 ? order.Requirements[0] : (0, 0);
            (int ItemId, int Num) req2 =
                order.Requirements.Length > 1 ? order.Requirements[1] : (0, 0);
            (int RewardId, int Num) energyReward = order.Rewards.FirstOrDefault(r =>
                r.RewardId == 1002
            );
            (int RewardId, int Num) coinReward = order.Rewards.FirstOrDefault(r =>
                r.RewardId == 1001
            );

            Cell(ws.Cells[row, 1], $"波次订单#{order.Idx}", hex);
            Cell(ws.Cells[row, 2], order.WaveDesc, "F5F5F5");
            Cell(ws.Cells[row, 3], order.TotalWave, "F5F5F5");
            Cell(ws.Cells[row, 4], $"ID:{req1.ItemId} ×{req1.Num}", hex);
            Cell(ws.Cells[row, 5], req2.ItemId > 0 ? $"ID:{req2.ItemId} ×{req2.Num}" : "—", hex);
            Cell(ws.Cells[row, 6], energyReward.Num > 0 ? (object)energyReward.Num : "—", "D9EAD3");
            Cell(ws.Cells[row, 7], coinReward.Num > 0 ? (object)coinReward.Num : "—", "FFF2CC");
            Cell(ws.Cells[row, 8], "动态随机(mainOrderID=-1)", "EBF3FB");
            row++;
        }

        // 空行分隔
        row++;

        // 主线订单结构说明
        var noteCell = ws.Cells[row, 1, row, 8];
        noteCell.Merge = true;
        noteCell.Value =
            "【主线订单结构】m_completeMainOrdersDic 中记录了已解锁的主线订单ID（10301~10806），"
            + "共35个已跟踪，按餐厅编号分段：103xx=餐厅3第1段…108xx=餐厅8。这是类似 TravelTown 线性推进的部分，"
            + "但玩家同时有动态波次订单（随机刷新）和主线进度订单（线性推进）——双轨并行设计。";
        noteCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        noteCell.Style.Fill.BackgroundColor.SetColor(HexColor("F0F4FA"));
        noteCell.Style.Font.Size = 9;
        noteCell.Style.WrapText = true;
        Border(noteCell);
        ws.Row(row).Height = 52;
        row += 2;

        // 箱子掉落
        var boxTitle = ws.Cells[row, 1, row, 8];
        boxTitle.Merge = true;
        boxTitle.Value = "■ 箱子掉落概率分析（来源：GAME_DYNAMIC_BOX_DROP_DATA）";
        boxTitle.Style.Font.Bold = true;
        boxTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        boxTitle.Style.Fill.BackgroundColor.SetColor(HexColor("4472C4"));
        boxTitle.Style.Font.Color.SetColor(Color.White);
        Border(boxTitle);
        ws.Row(row).Height = 20;
        row++;

        string[] boxHeaders = ["箱子ID", "第几次开", "掉落元素", "次数", "概率%", "置信度", "", ""];
        for (var j = 0; j < 6; j++)
            Header(ws.Cells[row, j + 1], boxHeaders[j], "2F5496");
        row++;

        foreach (var box in BoxDrops)
        {
            var totalDrops = box.Drops.Sum(d => d.Count);
            var isFirst = true;
            foreach (var (itemId, count) in box.Drops)
            {
                var pct = totalDrops > 0 ? (double)count / totalDrops * 100 : 0;
                var probHex =
                    pct >= 70
                        ? "D9EAD3"
                        : pct >= 30
                            ? "FFF2CC"
                            : "FCE5CD";
                if (isFirst)
                {
                    Cell(ws.Cells[row, 1], box.BoxId, "EBF3FB");
                    Cell(ws.Cells[row, 2], box.Seq, "EBF3FB");
                    isFirst = false;
                }
                else
                {
                    Cell(ws.Cells[row, 1], "", "F5F5F5");
                    Cell(ws.Cells[row, 2], "", "F5F5F5");
                }
                Cell(ws.Cells[row, 3], itemId, "FAFAFA");
                Cell(ws.Cells[row, 4], count, "FAFAFA");
                Cell(ws.Cells[row, 5], Math.Round(pct, 1), probHex);
                Cell(ws.Cells[row, 6], "中（样本量小）", "FFF8E8");
                row++;
            }
        }

        ws.Column(1).Width = 14;
        ws.Column(2).Width = 18;
        ws.Column(3).Width = 16;
        ws.Column(4).Width = 16;
        ws.Column(5).Width = 14;
        ws.Column(6).Width = 14;
        ws.Column(7).Width = 14;
        ws.Column(8).Width = 20;
        ws.Row(1).Height = 80;
        ws.Row(2).Height = 22;
    }

    // ── Sheet 4：活动日历 ────────────────────────────────────────────────────

    private static void BuildActivitySheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("活动日历");
        ws.View.FreezePanes(3, 1);

        MechNote(
            ws,
            1,
            1,
            "【Merge Cooking 活动配置 — 置信度：高（USER_CONSTANT_CONFIG_DATA2服务端下发，2026-05-14实时数据）】"
                + "配置来源：服务器实时下发 180 条配置，ActivityTime_xxxx 格式：开始日期_结束日期|参数…。"
                + "波次能量数据：waveEnergyUse=127（当轮波次总能耗），waveOnlineTime=231704秒（~64.4小时在线时长）。"
                + "StuckspotDic：GeneratorCd=62次（玩家等待CD最多），None=109次，CookingStart=40次，OutofEnergy=7次——"
                + "证实生成器CD是主要卡点，远超能量不足。weightGroup=build_1039000_default，ABTestGroup=default。",
            14
        );

        // Activity schedule entries (from USER_CONSTANT_CONFIG_DATA2, 2026-05-14)
        var activities = new (string Id, string Schedule, string Note)[]
        {
            (
                "1130",
                "20260505-20260509|Teaaroma2026; 20260510-20260515|MotherLove2026; 20260516-20260521|RideSpring2026; 20260522-20260526|TennisBall2026",
                "每5~7天切换主题"
            ),
            ("1270", "20260416-20260424|2组; 20260425-20260428|2组(续)", "排行赛/联赛系列"),
            (
                "1260",
                "20260421|16; 20260426|16; 20260428|16; 20260501|17; 20260503|17…",
                "单日活动，每2~3天一次"
            ),
            ("1490", "20260321-20260428|16; 20260429-20260606|17", "长周期(~6周)赛季活动"),
            (
                "1570",
                "20260504-20260510|14|3; 20260511-20260517|15|3; 20260518-20260524|15|3",
                "每周一期，3参数"
            ),
            ("1550", "20260504-20260510|2|0-604800-600|1 (周期重复)", "限时能量/特殊规则活动"),
            ("1560", "20260504-20260510|2|0-604800-600|1 (同上)", "与1550并行运行"),
            ("1710", "20260416-20260419|4×4天; 20260430-20260501|5×2天…", "每日活动，4~5参数"),
            ("1650", "20260423-20260426|3×4天", "短期4天活动"),
            ("1700", "20260413-20260419|4; 20260427-20260503|5", "每隔1周一次，共7天"),
            ("1400", "20260504-20260510|2", "每周一次"),
            ("1440", "20260504-20260506|1; 20260511-20260513|1; 20260518-20260520|1", "每周3天"),
            ("1430", "20260507-20260510|2×4天", "每周4天活动"),
            ("1540", "20260501-20260503|1; …(每周)", "每周3天"),
            ("1680", "20260507|1; 20260514|1; 20260521|1", "每周四单日"),
            ("1780", "20260512; 20260519; 20260526; 20260602", "每周二单日"),
            ("1170", "20260501; 20260508; 20260515; 20260522", "每周四单日（不同于1680）"),
            ("1840", "20260514-20260517|1×4天", "连续4天活动"),
            ("1231", "20260509-20260510; 20260516-20260517…每周末", "每周末2天"),
            ("1620", "20260518-20260524|7", "单周，7参数"),
            ("1590", "20260501-20260504|2; 20260515-20260518|2", "两期，4天各"),
            ("1610", "20260511-20260515|1", "5天短活动"),
            ("1830", "20260511-20260517|1", "7天活动"),
            ("1340", "20260412-20260620|Default", "长期常驻活动(~10周)"),
            ("1230", "20260508-20260528|21|20260527", "3周活动，带截止节点"),
            ("1020", "20260508-20260511|1; 20260522-20260525|1", "隔两周各4天"),
            ("1820", "20260525-20260527|1; 20260601-20260603|1; 20260608-20260610|1", "每隔一周3天"),
        };

        string[] headers = ["活动ID", "活动周期/参数", "规律/备注", "周期类型"];
        string[] hdrHex = ["1F4E79", "2F5496", "595959", "2F5496"];
        for (var j = 0; j < headers.Length; j++)
            Header(ws.Cells[2, j + 1], headers[j], hdrHex[j]);

        var colors = new[] { "EBF3FB", "FFF2CC", "D9EAD3", "FCE5CD", "E8DAEF", "F3F3F3" };
        var row = 3;

        for (var i = 0; i < activities.Length; i++)
        {
            var (id, schedule, note) = activities[i];
            var hex = colors[i % colors.Length];

            string cycleType;
            if (note.Contains("每周") || note.Contains("每周一") || note.Contains("周末"))
                cycleType = "每周";
            else if (note.Contains("每日") || note.Contains("单日"))
                cycleType = "每日";
            else if (note.Contains("常驻") || note.Contains("长期"))
                cycleType = "常驻";
            else if (note.Contains("赛季") || schedule.Contains("0606"))
                cycleType = "赛季(~6周)";
            else
                cycleType = "短期";

            var typeHex = cycleType switch
            {
                "每周" => "D9EAD3",
                "每日" => "FFF2CC",
                "常驻" => "E8F5FF",
                "赛季(~6周)" => "F3D6E4",
                _ => "F5F5F5"
            };

            Cell(ws.Cells[row, 1], $"Act_{id}", hex);
            Cell(ws.Cells[row, 2], schedule, "FAFAFA", wrap: true);
            Cell(ws.Cells[row, 3], note, "F0F0F0", wrap: true);
            Cell(ws.Cells[row, 4], cycleType, typeHex);
            ws.Row(row).Height = 30;
            row++;
        }

        // Wave energy stats block
        row += 2;
        var statTitle = ws.Cells[row, 1, row, 4];
        statTitle.Merge = true;
        statTitle.Value = "■ 波次能量统计 & 玩家行为数据（来源：WAVE_ORDER_DATA_KEY + USER_DATA）";
        statTitle.Style.Font.Bold = true;
        statTitle.Style.Fill.PatternType = ExcelFillStyle.Solid;
        statTitle.Style.Fill.BackgroundColor.SetColor(HexColor("1F4E79"));
        statTitle.Style.Font.Color.SetColor(Color.White);
        Border(statTitle);
        row++;

        var waveStats = new (string Key, string Value, string Desc)[]
        {
            ("waveEnergyUse", "127", "当轮波次累计消耗能量"),
            ("waveOnlineTime", "231704秒 (~64.4小时)", "当轮波次玩家在线总时长"),
            ("orderTotal", "71", "玩家总完成订单数（含历史）"),
            ("MergeConsumeTotal", "42", "累计合并操作次数"),
            ("MergeConsumeRound", "31", "本轮合并次数"),
            ("UseMachineNumber", "39", "总使用机器次数（点击生成器）"),
            ("DiamondConsumeTotal", "31", "累计消耗钻石数"),
            ("playDay", "2", "玩家游戏天数"),
            ("GeneratorCd (StuckspotDic)", "62次", "等待生成器CD次数（主要卡点）"),
            ("None (StuckspotDic)", "109次", "无明确卡点操作次数"),
            ("CookingStart (StuckspotDic)", "40次", "开始烹饪操作次数"),
            ("OutofEnergy (StuckspotDic)", "7次", "能量不足卡点次数"),
        };

        string[] statHeaders = ["指标", "值", "说明", ""];
        for (var j = 0; j < 3; j++)
            Header(ws.Cells[row, j + 1], statHeaders[j], "2F5496");
        row++;

        foreach (var (key, value, desc) in waveStats)
        {
            Cell(ws.Cells[row, 1], key, "EBF3FB");
            Cell(ws.Cells[row, 2], value, "FFF2CC");
            Cell(ws.Cells[row, 3], desc, "F5F5F5");
            row++;
        }

        ws.Column(1).Width = 14;
        ws.Column(2).Width = 55;
        ws.Column(3).Width = 28;
        ws.Column(4).Width = 14;
        ws.Row(1).Height = 70;
        ws.Row(2).Height = 22;
    }

    // ── Sheet 5：核心设计总结 ─────────────────────────────────────────────────

    private static void BuildSummarySheet(ExcelPackage pkg)
    {
        var ws = pkg.Workbook.Worksheets.Add("核心设计总结");
        var row = 1;

        void BigTitle(string text, int r)
        {
            var c = ws.Cells[r, 1, r, 12];
            c.Merge = true;
            c.Value = text;
            c.Style.Font.Bold = true;
            c.Style.Font.Size = 14;
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor("1F2D40"));
            c.Style.Font.Color.SetColor(Color.White);
            c.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Border(c);
            ws.Row(r).Height = 32;
        }

        void SectionTitle(string text, int r, string hex = "2F5496")
        {
            var c = ws.Cells[r, 1, r, 12];
            c.Merge = true;
            c.Value = text;
            c.Style.Font.Bold = true;
            c.Style.Font.Size = 11;
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
            c.Style.Font.Color.SetColor(Color.White);
            Border(c);
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
            var c = ws.Cells[r, 1, r, 12];
            c.Merge = true;
            c.Value = text;
            c.Style.Font.Size = 10;
            c.Style.Font.Bold = bold;
            c.Style.Fill.PatternType = ExcelFillStyle.Solid;
            c.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
            c.Style.WrapText = true;
            c.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Border(c);
            ws.Row(r).Height = height;
        }

        // ── 大标题 ──
        BigTitle("Merge Cooking  核心玩法设计总结  ·  逆向工程分析报告", row++);
        TextBlock(
            "数据来源：①ADB playerprefs.xml（25个JSON文件，运行时SharedPreferences）"
                + "  ②UnityPy解包 configlanguagefolder.unity3d（lang_ui_en.json，19806条英文名称）"
                + "  ③global-metadata.dat 4字节XOR解密（key=75E5E6E9）"
                + "  ④USER_CONSTANT_CONFIG_DATA2（180条服务端配置）"
                + "  |  游戏：com.merge.cooking.theme.restaurant.food  |  技术栈：IL2CPP Unity（无Lua）"
                + "  |  版本：1.39.0  |  分析日期：2026-05-14  |  数据置信度：总体高",
            row++,
            "E8EEF8",
            22
        );
        row++;

        // ── 一句话设计思想 ──
        SectionTitle("■ 一句话设计思想", row++, "1A3A5C");
        TextBlock(
            "「双轨订单（动态波次+线性主线）× 概率产出生成器 × 多餐厅渐进解锁」——"
                + "Merge Cooking 是 Gossip Harbor 动态随机风格与 TravelTown 线性推进风格的混合体："
                + "短期靠波次订单（动态刷新能量/金币）驱动高频操作，长期靠餐厅主线（固定ID顺序）提供成就感锚点。",
            row++,
            "EAF0FA",
            32,
            true
        );
        row++;

        // ── 核心循环流程图 ──
        SectionTitle("■ 核心循环流程图（EPPlus Shape）", row++, "1F4E79");

        const int StepRowCount = 3;
        const int StepRowH = 22;
        const int ArrowRowH = 12;
        const int TotalSteps = 8;
        const int BlockRows = TotalSteps * (StepRowCount + 1) + 2;
        for (var ir = 0; ir < BlockRows; ir++)
        {
            var isArrow = ir % (StepRowCount + 1) == StepRowCount;
            ws.Row(row + ir).Height = isArrow ? ArrowRowH : StepRowH;
        }

        var stepDefs = new (string Text, string Bg)[]
        {
            ("① 玩家点击生成器\n消耗能量 → 概率产出食材元素\n（多档概率：常见L1约50%，稀有L3约8%）", "1F4E79"),
            ("② 元素落盘\n棋盘出现食材，等待合并操作\n（GoodsState=1 可操作元素）", "2E75B6"),
            ("③ 玩家合并\n2个相同等级元素 → Lv+1（标准2合1）\n链深度：各链约4~8级，稀有度随级升", "2F5496"),
            ("④ 波次订单需求匹配\n当前波次(wave group)中提供5个动态订单\n每单需2种不同元素，全随机刷新(mainOrderID=-1)", "4472C4"),
            ("⑤ 提交订单\n消耗2种食材元素 → 获得能量+金币\n能量用于继续点击生成器（能量闭环）", "1F6E4A"),
            ("⑥ 波次推进\n完成 inWaveOrderIndex 顺序积累\n达到 totalWave(8) → 进入下一波组，难度提升", "276040"),
            ("⑦ 主线餐厅推进（并行长期线）\n完成主线订单 10301~10806（按餐厅线性解锁）\n解锁新餐厅 → 新食材链 → 新生成器", "7B3F00"),
            (
                "⑧ 建造餐厅 / 活动参与\n金币投入建筑解锁，限时活动(LimitTimeOrder)提供特殊资源\n活动链（101554~101557）提供主题元素路线",
                "4A235A"
            ),
        };

        for (var si = 0; si < stepDefs.Length; si++)
        {
            var (text, bg) = stepDefs[si];
            var baseRow = row + si * (StepRowCount + 1);
            var bgColor = HexColor(bg);
            var darkColor = Color.FromArgb(
                Math.Max(0, bgColor.R - 25),
                Math.Max(0, bgColor.G - 25),
                Math.Max(0, bgColor.B - 25)
            );
            var boxH = StepRowCount * StepRowH;

            var shape = ws.Drawings.AddShape(
                $"MCFlowStep{si}",
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
                $"MCFlowBadge{si}",
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
                    $"MCFlowArrow{si}",
                    OfficeOpenXml.Drawing.eShapeStyle.DownArrow
                );
                arr.SetPosition(arrRow - 1, 1, 2, 8);
                arr.SetSize(30, ArrowRowH + 2);
                arr.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
                arr.Fill.Color = HexColor("7F7F7F");
                arr.Border.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.NoFill;
            }
        }

        var loopRow = row + stepDefs.Length * (StepRowCount + 1);
        var loopCell = ws.Cells[loopRow, 1, loopRow, 12];
        loopCell.Merge = true;
        loopCell.Value = "↑───── 订单产出能量 → 点击生成器继续产出 → 短期能量闭环；餐厅升级 → 新链解锁 → 长期扩展 ─────↑";
        loopCell.Style.Font.Bold = true;
        loopCell.Style.Font.Size = 10;
        loopCell.Style.Font.Color.SetColor(HexColor("145A32"));
        loopCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        loopCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        loopCell.Style.Fill.BackgroundColor.SetColor(HexColor("E8F5EE"));
        Border(loopCell);
        ws.Row(loopRow).Height = 16;

        row += BlockRows + 2;

        // ── 三方对比表 ──
        SectionTitle("■ Merge Cooking vs Travel Town vs Gossip Harbor  三方机制对比", row++, "4A235A");

        var colDefs = new[]
        {
            (1, 2, "对比维度"),
            (3, 4, "Merge Cooking"),
            (5, 7, "Travel Town"),
            (8, 12, "Gossip Harbor")
        };
        foreach (var (c1, c2, label) in colDefs)
        {
            var hc = ws.Cells[row, c1, row, c2];
            hc.Merge = true;
            hc.Value = label;
            hc.Style.Font.Bold = true;
            hc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            hc.Style.Fill.BackgroundColor.SetColor(
                HexColor(
                    label == "Merge Cooking"
                        ? "C55A11"
                        : label == "Travel Town"
                            ? "1F4E79"
                            : label == "Gossip Harbor"
                                ? "375623"
                                : "2F5496"
                )
            );
            hc.Style.Font.Color.SetColor(Color.White);
            hc.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            Border(hc);
        }
        ws.Row(row++).Height = 20;

        var compRows = new[]
        {
            (
                "生成器产出",
                "概率随机池\n多档概率（L1≈50%,L2≈28%,L3≈8%）\n无4因子，固定概率",
                "固定确定产出\nweight=100，唯一目标\n完全无随机",
                "动态权重随机\n4因子(BaseW/CapFactor/\nChainMultiple/RepeatFactor)\n实时调整"
            ),
            (
                "订单模式",
                "双轨并行：\n短期：波次动态随机(wave group)\n长期：餐厅主线线性推进",
                "纯线性任务树\n103棵树，每树3~10步\n强叙事锚点",
                "纯动态随机刷新\nDiffSum分档驱动\n无叙事锚点"
            ),
            ("订单需求", "2种元素各1件\n（双元素需求）", "1~3种元素\n（多件数量需求）", "2~4种元素\n按权重池随机组合"),
            (
                "难度表达",
                "概率×链深度（隐式）\n无averageDifficulty字段\n无动态状态感知",
                "averageDifficulty 硬编码\nper-spawner难度矩阵\n设计师精确控制",
                "CurWeight运行时计算\n4因子动态乘积\n自适应调整"
            ),
            (
                "生成器升级",
                "同类型机器存在级别差异\n(101554-101557,surplus差异)\n但无明确合并升级路线",
                "机器本身可合并升级\nLv.0~Lv.6,高级机CD更短",
                "生成器不升级\n通过权重区分6条链定位"
            ),
            (
                "棋盘互动",
                "GoodsState区分状态\n波次订单感知棋盘元素\n有bag存储系统(GoodsState=11)",
                "棋盘是提交工具\n不影响订单难度计算",
                "棋盘状态影响权重\n棋盘→反哺生成器权重闭环"
            ),
            (
                "技术栈",
                "IL2CPP Unity\nSharedPreferences持久化\nJSON本地化存储",
                "IL2CPP Unity\nHTTPCache CDN下发JSON\n配置与逻辑分离",
                "Lua 5.4字节码\n动态权重运行时计算\n配置+逻辑均在Lua层"
            ),
        };

        foreach (var (dim, mc, tt, harbor) in compRows)
        {
            void FC(int c1, int c2, string v, string h)
            {
                var c = ws.Cells[row, c1, row, c2];
                c.Merge = true;
                c.Value = v;
                c.Style.Font.Size = 10;
                c.Style.Fill.PatternType = ExcelFillStyle.Solid;
                c.Style.Fill.BackgroundColor.SetColor(HexColor(h));
                c.Style.WrapText = true;
                c.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                Border(c);
            }
            FC(1, 2, dim, "F0F4FA");
            ws.Cells[row, 1, row, 2].Style.Font.Bold = true;
            FC(3, 4, mc, "FEF0E7");
            FC(5, 7, tt, "E8F5FF");
            FC(8, 12, harbor, "F0FFF0");
            ws.Row(row++).Height = 52;
        }
        row++;

        // ── 关键发现 ──
        SectionTitle("■ 三大关键机制发现（对标 Harbor/TT）", row++, "7B3F00");

        var findings = new[]
        {
            (
                "发现①  双轨订单是独特设计",
                "Merge Cooking 同时运行「波次动态订单」和「餐厅主线订单」两条轨道。"
                    + "Harbor 只有动态随机（无叙事锚点，长期体验空洞）；TT 只有线性树（短期反馈慢，挫败感强）。"
                    + "MC 用波次订单保证高频反馈（能量循环短），主线提供成就感，兼顾了两者优点。"
                    + "设计建议：我方游戏可参考此双轨结构，短期订单（5~8个/波次随机刷新）+ 长期主线（5~10个里程碑/餐厅）。",
                "FEF0E7"
            ),
            (
                "发现②  概率生成器介于 TT 与 Harbor 之间",
                "TT 固定产出（无惊喜），Harbor 动态权重（高度不确定）。"
                    + "MC 采用静态概率池（L1≈50%,L2≈28%,L3≈8%），保留随机乐趣（偶尔出高级元素）又不过于复杂。"
                    + "关键：没有 Harbor 的「棋盘状态→权重调整」反馈，概率配置更简单，运营成本低。"
                    + "对我们的启示：生成器概率池在 2~4 个档位，高级占比 10~20%，是比较平衡的休闲设计。",
                "E8F5FF"
            ),
            (
                "发现③  「Bag存储」系统降低棋盘压力",
                "GoodsState=11 的元素存放在 bag（存储格）而非棋盘上，共有10个bag位置（ID: 100503, 100502等）。"
                    + "Harbor 和 TT 都没有独立的 bag 存储区，棋盘格子一满就无法继续生成。"
                    + "MC 的 bag 系统允许玩家暂存高级元素，减少棋盘整理焦虑——是提升体验的重要细节。"
                    + "MC 还有 BAG_GENERATOR_STORAGE_KEY（生成器自己的专属仓储），设计层次更丰富。",
                "F0FFF0"
            ),
        };

        foreach (var (title, desc, hex) in findings)
        {
            var lc = ws.Cells[row, 1, row, 3];
            lc.Merge = true;
            lc.Value = title;
            lc.Style.Font.Bold = true;
            lc.Style.Font.Size = 10;
            lc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lc.Style.Fill.BackgroundColor.SetColor(HexColor("7B3F00"));
            lc.Style.Font.Color.SetColor(Color.White);
            lc.Style.WrapText = true;
            lc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Border(lc);

            var rc = ws.Cells[row, 4, row, 12];
            rc.Merge = true;
            rc.Value = desc;
            rc.Style.Font.Size = 10;
            rc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rc.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
            rc.Style.WrapText = true;
            rc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Border(rc);
            ws.Row(row++).Height = 58;
        }
        row++;

        // ── 数据来源 & 置信度 ──
        SectionTitle("■ 数据来源 & 置信度", row++, "595959");

        var sources = new[]
        {
            (
                "高置信（直接运行时数据）",
                "playerprefs.xml 所有字段（GameOrdersModel, GAME_LEVEL_MAP_DATA, GENERATOR_MACHINE_UUID_DATA等）：直接读取游戏运行时SharedPreferences，100%准确。\n"
                    + "订单系统：mDynamicOrdersOpen=true，mainOrderID=-1，waveGroupID/totalWave/inWaveOrderIndex：直接字段，高置信。\n"
                    + "生成器surplus、InitiativeProduceList历史记录：直接读取，准确反映玩家当前游戏状态。\n"
                    + "活动配置：USER_CONSTANT_CONFIG_DATA2（180条配置）ActivityTime_1xxx系列、ActivityHardSwitch/SoftSwitch：服务端下发实时配置，高置信。\n"
                    + "付费体系（语言包确认）：Diamonds/Energy/Coins三币种，Pass产品=Chef Pass(TaskPass)+Diamond Carnival(DiamondPass)+Energy Weekly Pass+Video Bonus Pass，有First Purchase/Elegant/Silver/Golden/Master Chef Offer等6档礼包。",
                "E8F5E8"
            ),
            (
                "高置信（UnityPy资产解包）",
                "元素英文名：configlanguagefolder.unity3d（APK内，Bundle头前缀34字节→剥离后UnityPy解析）→ lang_ui_en.json（19806条）。\n"
                    + "食材系列：C_leaf(10级:Lettuce→Artichoke), C_root(15级:Tomato→Corn), C_meat(12级:Bacon→Whole Lamb), "
                    + "C_egg(10级:Egg→Turkey), C_milk(9级:Frothed Milk→Parmesan), C_softdrink(17级:Water→Energy Drink), "
                    + "C_coffee(17级:Coffee Bean→Siphon Coffee), C_glass(7级:Broken Glass→Glass Dessert Stand), "
                    + "C_coconutmilk(9级:Glass of Coconut Water→Coconut Candy), C_coconutshell(7级:Coconut Shell→Coconut Teacup)。\n"
                    + "机器系列：M_pan/M_grill/M_pot/M_juicer/M_table 各10级，名称不随升级改变（均为Pan/Grill/Pot/Juicer/Chef's Counter）。\n"
                    + "主题机：G_america=Tourist Baggage(Lv1~7), c_americagift1~10=Bald Eagle Stamp→Desk Decoration。\n"
                    + "菜谱名：DS_american2_1~55 = Baby Back Ribs/Lamb Kebab/Burnt Ends等BBQ系列共55道菜。\n"
                    + "活动：HF=Happy Fishing(Bluegill Sunfish等10种鱼), IML_egypt=Noble by the Nile(埃及主题)。",
                "E8F5E8"
            ),
            (
                "高置信（metadata解密验证）",
                "global-metadata.dat：4字节循环XOR key=0x75 0xE5 0xE6 0xE9，解密后magic=0xAF1BB1FA验证通过，IL2CPP metadata确认。\n"
                    + "但字符串池本身二次混淆（非标准明文字符串池，熵值仅4.51但内容为乱码），类名无法直接提取。\n"
                    + "Assembly-CSharp.dll：CDPH自定义header格式，内部DLL加密（非MZ标准），但从raw strings提取到60191个字符串，"
                    + "确认存在类名：AddOrderScore, AddRaceExtraOrderToken, GetGoodsListBySeriesID, PlayInitiativeCDAnim等。\n"
                    + "2.dat：base64+AES加密（熵值8.00，多256字节倍数），非单字节XOR，密钥未获取。",
                "E8F5FF"
            ),
            (
                "中置信（推断）",
                "主题机分级名称（Tourist Baggage Lv1~5）：按surplus递增顺序命名，g_america语言包名固定为Tourist Baggage不带level后缀。\n"
                    + "合并链内「订单消耗」标记：仅基于5个动态订单样本，不代表全部订单需求分布。\n"
                    + "生成器产出概率：样本量4~18，统计可靠但非精确概率表。",
                "FFF8E8"
            ),
            (
                "未获取 / 无法分析",
                "完整配置表（CD秒数、合并规则、生成器解锁等级条件）：2.dat为AES加密（base64+AES，密钥未破解）。\n"
                    + "il2cpp类/字段名：global-metadata.dat字符串池二次混淆；Assembly-CSharp.dll CDPH格式内部加密。\n"
                    + "configstaticdatafolder.unity3d + configbase_ab：Bundle前缀后以'JIAMI'（加密）标记，UnityPy对象数=0，静态配置不可读。\n"
                    + "元素中文名：语言包仅提供英文；付费精确价格：无IAP SKU价格数据（需Google Play API）。",
                "FFF0E8"
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
            lc.Style.Fill.BackgroundColor.SetColor(HexColor("595959"));
            lc.Style.Font.Color.SetColor(Color.White);
            lc.Style.WrapText = true;
            lc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Border(lc);

            var rc = ws.Cells[row, 4, row, 12];
            rc.Merge = true;
            rc.Value = desc;
            rc.Style.Font.Size = 10;
            rc.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rc.Style.Fill.BackgroundColor.SetColor(HexColor(hex));
            rc.Style.WrapText = true;
            rc.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            Border(rc);
            ws.Row(row++).Height = 52;
        }

        ws.Column(1).Width = 12;
        for (var j = 2; j <= 12; j++)
            ws.Column(j).Width = 16;
    }
}
