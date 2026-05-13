using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// LTE 限时地图通用写入器。
/// 用法：LteMapWriter.Run(xlsxPath, mapN)  写入指定图编号（1~10）
///       LteMapWriter.RunAll(xlsxPath)      写入全部10张图
/// 每张图60行：有效节点行写业务数据+地图网格色块，剩余填"通用"。
/// 列宽13.0字符，行高14.25磅（chelizi标准）。
/// </summary>
public static class LteMapWriter
{
    private const long PeriodId = 764101;
    private const long ResBase = 762101;
    private const int Target = 60;
    private const int GridCols = 10;

    // ── 数据模型 ─────────────────────────────────────────────────────────────

    private sealed record Consume(string Item, int Qty);

    private sealed record Produce(string Item, int Qty);

    private sealed class NodeDef
    {
        public string Code { get; init; } = "";
        public string Pkg { get; init; } = "";
        public int? Level { get; init; }
        public string Typ { get; init; } = "";
        public int Count { get; init; }
        public int? Res4d { get; init; }
        public List<Consume> ConsumeList { get; init; } = [];
        public List<Produce> ProduceList { get; init; } = [];
    }

    private sealed class MapData
    {
        public int MapN { get; init; }
        public Dictionary<(int R, int C), string> Grid { get; init; } = [];
        public int GridRows { get; init; } = 8;
        public List<NodeDef> Nodes { get; init; } = [];
    }

    // ── 颜色 ─────────────────────────────────────────────────────────────────

    private static Color BgColor(string typ) =>
        typ switch
        {
            "链" => Hex("1E3A5F"),
            "链-3合" => Hex("1E3A5F"),
            "链-终-兑材料" => Hex("3B0A0A"),
            "链-建" => Hex("1C1C1C"),
            "门钥匙-固定" => Hex("3B1515"),
            "门-固定" => Hex("3B2200"),
            "兑-材料" => Hex("2D1B69"),
            "矿" => Hex("1A3800"),
            "兑" => Hex("3B1229"),
            "采" => Hex("063B3B"),
            "地标-体力" => Hex("063B1E"),
            "特殊" => Hex("1A1A3B"),
            "爆发" => Hex("2D0A1A"),
            "防线" => Hex("3B1515"),
            _ => Hex("1E293B"),
        };

    private static Color FgColor(string typ) =>
        typ switch
        {
            "链" => Hex("93C5FD"),
            "链-3合" => Hex("BFDBFE"),
            "链-终-兑材料" => Hex("FCA5A5"),
            "链-建" => Hex("9CA3AF"),
            "门钥匙-固定" => Hex("FCA5A5"),
            "门-固定" => Hex("FDBA74"),
            "兑-材料" => Hex("C4B5FD"),
            "矿" => Hex("86EFAC"),
            "兑" => Hex("F9A8D4"),
            "采" => Hex("5EEAD4"),
            "地标-体力" => Hex("6EE7B7"),
            "特殊" => Hex("DDD6FE"),
            "爆发" => Hex("F43F5E"),
            "防线" => Hex("FCA5A5"),
            _ => Hex("E5E7EB"),
        };

    private static readonly Color EmptyBg = Hex("1E293B");

    private static Color Hex(string h)
    {
        var v = Convert.ToInt32(h, 16);
        return Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
    }

    // ── 主链公共节点 ─────────────────────────────────────────────────────────

    private static List<NodeDef> MainChainNodes(string p) =>
        [
            new()
            {
                Code = $"{p}链A-1",
                Pkg = "锚链",
                Level = 1,
                Typ = "链",
                Res4d = 1
            },
            new()
            {
                Code = $"{p}链A-2",
                Pkg = "锚链",
                Level = 2,
                Typ = "链",
                Res4d = 2,
                ConsumeList = [new($"{p}链A-1", 1)]
            },
            new()
            {
                Code = $"{p}链A-3",
                Pkg = "锚链",
                Level = 3,
                Typ = "链",
                Res4d = 3,
                ConsumeList = [new($"{p}链A-2", 1)]
            },
            new()
            {
                Code = $"{p}链A-4",
                Pkg = "锚链",
                Level = 4,
                Typ = "链",
                ConsumeList = [new($"{p}链A-3", 1)]
            },
            new()
            {
                Code = $"{p}链A-5",
                Pkg = "锚链",
                Level = 5,
                Typ = "链",
                ConsumeList = [new($"{p}链A-4", 1)]
            },
            new()
            {
                Code = $"{p}链A-6",
                Pkg = "锚链",
                Level = 6,
                Typ = "链-3合",
                ConsumeList = [new($"{p}链A-5", 1), new($"{p}B-A6", 1)]
            },
            new()
            {
                Code = $"{p}B-A6",
                Pkg = "链-建",
                Typ = "链-建",
                Count = 1,
                ConsumeList = [new("体力", 30)]
            },
            new()
            {
                Code = $"{p}链A-7",
                Pkg = "锚链",
                Level = 7,
                Typ = "链-终-兑材料",
                ConsumeList = [new($"{p}链A-6", 1), new($"{p}B-A7", 1)]
            },
            new()
            {
                Code = $"{p}B-A7",
                Pkg = "链-建",
                Typ = "链-建",
                Count = 1,
                ConsumeList = [new("体力", 40)]
            },
        ];

    // ══════════════════════════════════════════════════════════════════════════
    // 图1 — 渔村启程
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap1()
    {
        var p = "图1";
        return new MapData
        {
            MapN = 1,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 5), $"{p}S-网A" },
                { (4, 5), $"{p}I-鱼A" },
                { (5, 7), $"{p}S-网B" },
                { (4, 7), $"{p}I-鱼B" },
                { (2, 6), $"{p}D-汇" },
                { (1, 6), $"{p}C-启" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}S-网A",
                    Pkg = "渔网A",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-鱼A", 3)]
                },
                new()
                {
                    Code = $"{p}S-网B",
                    Pkg = "渔网B",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-鱼B", 3)]
                },
                new()
                {
                    Code = $"{p}I-鱼A",
                    Pkg = "鱼获A",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}I-鱼B",
                    Pkg = "鱼获B",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "组合兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-鱼A", 9), new($"{p}I-鱼B", 9)]
                },
                new()
                {
                    Code = $"{p}C-启",
                    Pkg = "启航采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 10),
                        new($"{p}链A-4", 10)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "渔村宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 8)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图2 — 沉船湾（已审阅，与 LteMap2Writer 保持一致）
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap2()
    {
        var p = "图2";
        return new MapData
        {
            MapN = 2,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 0), $"{p}Key-绳" },
                { (4, 0), $"{p}M-绳" },
                { (3, 0), $"{p}I-绳" },
                { (5, 1), $"{p}Key-铁" },
                { (4, 1), $"{p}M-铁" },
                { (3, 1), $"{p}I-铁" },
                { (4, 3), $"{p}K-罗" },
                { (3, 3), $"{p}I-罗" },
                { (4, 4), $"{p}K-图" },
                { (3, 4), $"{p}I-图" },
                { (2, 1), $"{p}D-汇" },
                { (1, 0), $"{p}C-启" },
                { (0, 8), $"{p}宝箱" },
                { (0, 9), $"{p}K-打" },
                { (1, 9), $"{p}I-碎" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}Key-绳",
                    Pkg = "舱门钥匙",
                    Typ = "门钥匙-固定",
                    Count = 3,
                    ConsumeList = [new("体力", 20)]
                },
                new()
                {
                    Code = $"{p}M-绳",
                    Pkg = "绳结舱门",
                    Typ = "门-固定",
                    Count = 3,
                    ConsumeList = [new($"{p}Key-绳", 1)],
                    ProduceList = [new($"{p}I-绳", 3)]
                },
                new()
                {
                    Code = $"{p}I-绳",
                    Pkg = "绳结",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}Key-铁",
                    Pkg = "舱门钥匙",
                    Typ = "门钥匙-固定",
                    Count = 3,
                    ConsumeList = [new("体力", 20)]
                },
                new()
                {
                    Code = $"{p}M-铁",
                    Pkg = "铁锚舱门",
                    Typ = "门-固定",
                    Count = 3,
                    ConsumeList = [new($"{p}Key-铁", 1)],
                    ProduceList = [new($"{p}I-铁", 3)]
                },
                new()
                {
                    Code = $"{p}I-铁",
                    Pkg = "铁锚",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-罗",
                    Pkg = "罗盘矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-罗", 4)]
                },
                new()
                {
                    Code = $"{p}I-罗",
                    Pkg = "罗盘",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-图",
                    Pkg = "海图矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-图", 4)]
                },
                new()
                {
                    Code = $"{p}I-图",
                    Pkg = "海图",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "组合兑换",
                    Typ = "兑",
                    ConsumeList =
                    [
                        new($"{p}I-绳", 9),
                        new($"{p}I-铁", 9),
                        new($"{p}I-罗", 12),
                        new($"{p}I-图", 12)
                    ]
                },
                new()
                {
                    Code = $"{p}C-启",
                    Pkg = "启航采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 20)
                    ]
                },
                new()
                {
                    Code = $"{p}K-打",
                    Pkg = "沉船打捞",
                    Typ = "矿",
                    Count = 1,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-碎", 3)]
                },
                new()
                {
                    Code = $"{p}I-碎",
                    Pkg = "船锚碎片",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "航海日志",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图3 — 珊瑚礁市集
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap3()
    {
        var p = "图3";
        return new MapData
        {
            MapN = 3,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 4), $"{p}S-贝" },
                { (4, 4), $"{p}I-贝" },
                { (5, 5), $"{p}S-草" },
                { (4, 5), $"{p}I-草" },
                { (5, 6), $"{p}S-珠" },
                { (4, 6), $"{p}I-珠" },
                { (5, 7), $"{p}S-鱼" },
                { (4, 7), $"{p}I-鱼" },
                { (2, 5), $"{p}D-集" },
                { (1, 5), $"{p}C-开" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}S-贝",
                    Pkg = "贝壳摊",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-贝", 4)]
                },
                new()
                {
                    Code = $"{p}I-贝",
                    Pkg = "贝壳币",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}S-草",
                    Pkg = "海草摊",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-草", 4)]
                },
                new()
                {
                    Code = $"{p}I-草",
                    Pkg = "海草币",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}S-珠",
                    Pkg = "珍珠摊",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-珠", 4)]
                },
                new()
                {
                    Code = $"{p}I-珠",
                    Pkg = "珍珠币",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}S-鱼",
                    Pkg = "鱼干摊",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-鱼", 4)]
                },
                new()
                {
                    Code = $"{p}I-鱼",
                    Pkg = "鱼干币",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-集",
                    Pkg = "集市兑换",
                    Typ = "兑",
                    ConsumeList =
                    [
                        new($"{p}I-贝", 12),
                        new($"{p}I-草", 12),
                        new($"{p}I-珠", 12),
                        new($"{p}I-鱼", 12)
                    ]
                },
                new()
                {
                    Code = $"{p}C-开",
                    Pkg = "开市采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "市集宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 8)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图4 — 灯塔要塞
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap4()
    {
        var p = "图4";
        return new MapData
        {
            MapN = 4,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 1), $"{p}链A-1" },
                { (7, 2), $"{p}Gate-A" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (2, 2), $"{p}Gate-B" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}Gate-C" },
                { (0, 4), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 5), $"{p}K-信" },
                { (4, 5), $"{p}I-信" },
                { (5, 7), $"{p}K-油" },
                { (4, 7), $"{p}I-油" },
                { (2, 6), $"{p}D-汇" },
                { (1, 6), $"{p}C-灯" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}Gate-A",
                    Pkg = "防线A",
                    Typ = "防线",
                    Count = 1,
                    ConsumeList = [new("体力", 25)]
                },
                new()
                {
                    Code = $"{p}Gate-B",
                    Pkg = "防线B",
                    Typ = "防线",
                    Count = 1,
                    ConsumeList = [new("体力", 35)]
                },
                new()
                {
                    Code = $"{p}Gate-C",
                    Pkg = "防线C",
                    Typ = "防线",
                    Count = 1,
                    ConsumeList = [new("体力", 45)]
                },
                new()
                {
                    Code = $"{p}K-信",
                    Pkg = "信号弹矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-信", 4)]
                },
                new()
                {
                    Code = $"{p}I-信",
                    Pkg = "信号材料",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-油",
                    Pkg = "燃油矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-油", 4)]
                },
                new()
                {
                    Code = $"{p}I-油",
                    Pkg = "燃油",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "组合兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-信", 12), new($"{p}I-油", 12)]
                },
                new()
                {
                    Code = $"{p}C-灯",
                    Pkg = "灯塔采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "要塞宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图5 — 深海遗迹
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap5()
    {
        var p = "图5";
        return new MapData
        {
            MapN = 5,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 4), $"{p}K-古A" },
                { (5, 5), $"{p}K-古B" },
                { (4, 4), $"{p}I-古" },
                { (3, 4), $"{p}Lock-宝" },
                { (2, 4), $"{p}K-宝" },
                { (1, 4), $"{p}I-宝" },
                { (5, 7), $"{p}K-晶" },
                { (4, 7), $"{p}I-晶" },
                { (2, 6), $"{p}D-汇" },
                { (1, 6), $"{p}C-遗" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}K-古A",
                    Pkg = "古文矿A",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-古", 3)]
                },
                new()
                {
                    Code = $"{p}K-古B",
                    Pkg = "古文矿B",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-古", 3)]
                },
                new()
                {
                    Code = $"{p}I-古",
                    Pkg = "古代铭文",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}Lock-宝",
                    Pkg = "宝库封印",
                    Typ = "防线",
                    Count = 1,
                    ConsumeList = [new($"{p}I-古", 6)]
                },
                new()
                {
                    Code = $"{p}K-宝",
                    Pkg = "遗迹宝库",
                    Typ = "矿",
                    Count = 4,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-宝", 4)]
                },
                new()
                {
                    Code = $"{p}I-宝",
                    Pkg = "遗迹宝石",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-晶",
                    Pkg = "水晶矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-晶", 4)]
                },
                new()
                {
                    Code = $"{p}I-晶",
                    Pkg = "水晶",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "深海兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-宝", 16), new($"{p}I-晶", 12)]
                },
                new()
                {
                    Code = $"{p}C-遗",
                    Pkg = "遗迹采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "遗迹宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图6 — 海盗巢穴（双主链）
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap6()
    {
        var p = "图6";
        return new MapData
        {
            MapN = 6,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (7, 4), $"{p}B-1" },
                { (6, 4), $"{p}B-2" },
                { (5, 4), $"{p}B-3" },
                { (4, 4), $"{p}B-4" },
                { (7, 6), $"{p}C-1" },
                { (6, 6), $"{p}C-2" },
                { (5, 6), $"{p}C-3" },
                { (4, 6), $"{p}C-4" },
                { (3, 5), $"{p}D-1" },
                { (5, 8), $"{p}K-刀" },
                { (4, 8), $"{p}I-刀" },
                { (5, 9), $"{p}K-金" },
                { (4, 9), $"{p}I-金" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}B-1",
                    Pkg = "掠夺链1",
                    Level = 1,
                    Typ = "链"
                },
                new()
                {
                    Code = $"{p}B-2",
                    Pkg = "掠夺链2",
                    Level = 2,
                    Typ = "链",
                    ConsumeList = [new($"{p}B-1", 1)]
                },
                new()
                {
                    Code = $"{p}B-3",
                    Pkg = "掠夺链3",
                    Level = 3,
                    Typ = "链",
                    ConsumeList = [new($"{p}B-2", 1)]
                },
                new()
                {
                    Code = $"{p}B-4",
                    Pkg = "掠夺链4",
                    Level = 4,
                    Typ = "链",
                    ConsumeList = [new($"{p}B-3", 1), new($"{p}I-刀", 1)]
                },
                new()
                {
                    Code = $"{p}C-1",
                    Pkg = "藏宝链1",
                    Level = 1,
                    Typ = "链"
                },
                new()
                {
                    Code = $"{p}C-2",
                    Pkg = "藏宝链2",
                    Level = 2,
                    Typ = "链",
                    ConsumeList = [new($"{p}C-1", 1)]
                },
                new()
                {
                    Code = $"{p}C-3",
                    Pkg = "藏宝链3",
                    Level = 3,
                    Typ = "链",
                    ConsumeList = [new($"{p}C-2", 1)]
                },
                new()
                {
                    Code = $"{p}C-4",
                    Pkg = "藏宝链4",
                    Level = 4,
                    Typ = "链",
                    ConsumeList = [new($"{p}C-3", 1), new($"{p}I-金", 1)]
                },
                new()
                {
                    Code = $"{p}D-1",
                    Pkg = "宝藏终点",
                    Typ = "特殊",
                    ConsumeList = [new($"{p}B-4", 1), new($"{p}C-4", 1)],
                    ProduceList =
                    [
                        new($"{p}链A-1", 15),
                        new($"{p}链A-2", 15),
                        new($"{p}链A-3", 20),
                        new($"{p}链A-4", 20)
                    ]
                },
                new()
                {
                    Code = $"{p}K-刀",
                    Pkg = "刀具矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-刀", 4)]
                },
                new()
                {
                    Code = $"{p}I-刀",
                    Pkg = "刀具",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-金",
                    Pkg = "金币矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-金", 4)]
                },
                new()
                {
                    Code = $"{p}I-金",
                    Pkg = "金币",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "海盗宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 12)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图7 — 风暴峡谷
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap7()
    {
        var p = "图7";
        return new MapData
        {
            MapN = 7,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 4), $"{p}K-晶A" },
                { (5, 5), $"{p}K-晶B" },
                { (5, 6), $"{p}K-晶C" },
                { (5, 7), $"{p}K-晶D" },
                { (4, 5), $"{p}I-晶" },
                { (3, 5), $"{p}D-蓄" },
                { (2, 5), $"{p}I-能" },
                { (5, 9), $"{p}K-闪" },
                { (4, 9), $"{p}I-闪" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}K-晶A",
                    Pkg = "晶体矿A",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-晶", 4)]
                },
                new()
                {
                    Code = $"{p}K-晶B",
                    Pkg = "晶体矿B",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-晶", 4)]
                },
                new()
                {
                    Code = $"{p}K-晶C",
                    Pkg = "晶体矿C",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-晶", 4)]
                },
                new()
                {
                    Code = $"{p}K-晶D",
                    Pkg = "晶体矿D",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 15)],
                    ProduceList = [new($"{p}I-晶", 4)]
                },
                new()
                {
                    Code = $"{p}I-晶",
                    Pkg = "风暴晶体",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-蓄",
                    Pkg = "蓄力兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-晶", 48)]
                },
                new()
                {
                    Code = $"{p}I-能",
                    Pkg = "风暴能量",
                    Typ = "特殊",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}K-闪",
                    Pkg = "闪电爆发矿",
                    Typ = "爆发",
                    Count = 1,
                    ConsumeList = [new("体力", 70)],
                    ProduceList = [new($"{p}I-闪", 20)]
                },
                new()
                {
                    Code = $"{p}I-闪",
                    Pkg = "闪电材料",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "峡谷宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图8 — 冰封港湾
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap8()
    {
        var p = "图8";
        return new MapData
        {
            MapN = 8,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (7, 4), $"{p}K-解" },
                { (6, 4), $"{p}K-冰A" },
                { (5, 4), $"{p}I-冰A" },
                { (6, 6), $"{p}K-冰B" },
                { (5, 6), $"{p}I-冰B" },
                { (6, 8), $"{p}K-冰C" },
                { (5, 8), $"{p}I-冰C" },
                { (3, 6), $"{p}D-汇" },
                { (2, 6), $"{p}C-港" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}K-解",
                    Pkg = "解冻炉",
                    Typ = "防线",
                    Count = 1,
                    ConsumeList = [new("体力", 30)]
                },
                new()
                {
                    Code = $"{p}K-冰A",
                    Pkg = "冰矿A",
                    Typ = "矿",
                    Count = 5,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-冰A", 3)]
                },
                new()
                {
                    Code = $"{p}I-冰A",
                    Pkg = "冰材A",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-冰B",
                    Pkg = "冰矿B",
                    Typ = "矿",
                    Count = 5,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-冰B", 3)]
                },
                new()
                {
                    Code = $"{p}I-冰B",
                    Pkg = "冰材B",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-冰C",
                    Pkg = "冰矿C",
                    Typ = "矿",
                    Count = 5,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-冰C", 3)]
                },
                new()
                {
                    Code = $"{p}I-冰C",
                    Pkg = "冰材C",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "融合兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-冰A", 15), new($"{p}I-冰B", 15), new($"{p}I-冰C", 15)]
                },
                new()
                {
                    Code = $"{p}C-港",
                    Pkg = "港湾采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 20),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "港湾宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图9 — 火山群岛
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap9()
    {
        var p = "图9";
        return new MapData
        {
            MapN = 9,
            GridRows = 8,
            Grid = new Dictionary<(int, int), string>
            {
                { (7, 2), $"{p}链A-1" },
                { (6, 2), $"{p}链A-2" },
                { (5, 2), $"{p}链A-3" },
                { (3, 2), $"{p}链A-4" },
                { (1, 2), $"{p}链A-5" },
                { (0, 2), $"{p}链A-6" },
                { (0, 3), $"{p}B-A6" },
                { (0, 0), $"{p}链A-7" },
                { (0, 1), $"{p}B-A7" },
                { (5, 4), $"{p}Burst-A" },
                { (5, 5), $"{p}Burst-B" },
                { (4, 4), $"{p}I-岩" },
                { (5, 7), $"{p}K-硫" },
                { (4, 7), $"{p}I-硫" },
                { (5, 9), $"{p}K-黑" },
                { (4, 9), $"{p}I-黑" },
                { (2, 6), $"{p}D-汇" },
                { (1, 6), $"{p}C-岛" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                .. MainChainNodes(p),
                new()
                {
                    Code = $"{p}Burst-A",
                    Pkg = "火山A",
                    Typ = "爆发",
                    Count = 1,
                    ConsumeList = [new("体力", 40)],
                    ProduceList = [new($"{p}I-岩", 15)]
                },
                new()
                {
                    Code = $"{p}Burst-B",
                    Pkg = "火山B",
                    Typ = "爆发",
                    Count = 1,
                    ConsumeList = [new("体力", 40)],
                    ProduceList = [new($"{p}I-岩", 15)]
                },
                new()
                {
                    Code = $"{p}I-岩",
                    Pkg = "岩浆石",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-硫",
                    Pkg = "硫磺矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-硫", 4)]
                },
                new()
                {
                    Code = $"{p}I-硫",
                    Pkg = "硫磺",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-黑",
                    Pkg = "黑曜矿",
                    Typ = "矿",
                    Count = 3,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-黑", 4)]
                },
                new()
                {
                    Code = $"{p}I-黑",
                    Pkg = "黑曜石",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}D-汇",
                    Pkg = "熔炉兑换",
                    Typ = "兑",
                    ConsumeList = [new($"{p}I-岩", 30), new($"{p}I-硫", 12), new($"{p}I-黑", 12)]
                },
                new()
                {
                    Code = $"{p}C-岛",
                    Pkg = "群岛采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 10),
                        new($"{p}链A-2", 10),
                        new($"{p}链A-3", 15),
                        new($"{p}链A-4", 15)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "火山宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 10)]
                },
            ],
        };
    }

    // ══════════════════════════════════════════════════════════════════════════
    // 图10 — 龙王宫殿（终章，主链延至A-8）
    // ══════════════════════════════════════════════════════════════════════════
    private static MapData BuildMap10()
    {
        var p = "图10";
        return new MapData
        {
            MapN = 10,
            GridRows = 9,
            Grid = new Dictionary<(int, int), string>
            {
                { (8, 2), $"{p}链A-1" },
                { (7, 2), $"{p}链A-2" },
                { (6, 2), $"{p}链A-3" },
                { (4, 2), $"{p}链A-4" },
                { (2, 2), $"{p}链A-5" },
                { (1, 2), $"{p}链A-6" },
                { (1, 3), $"{p}B-A6" },
                { (0, 2), $"{p}链A-7" },
                { (0, 3), $"{p}B-A7" },
                { (0, 0), $"{p}链A-8" },
                { (0, 1), $"{p}B-A8" },
                { (6, 4), $"{p}K-龙A" },
                { (5, 4), $"{p}I-龙A" },
                { (6, 6), $"{p}K-龙B" },
                { (5, 6), $"{p}I-龙B" },
                { (6, 8), $"{p}K-龙C" },
                { (5, 8), $"{p}I-龙C" },
                { (8, 5), $"{p}Burst-X" },
                { (7, 5), $"{p}I-爆X" },
                { (8, 7), $"{p}Burst-Y" },
                { (7, 7), $"{p}I-爆Y" },
                { (3, 5), $"{p}D-龙" },
                { (2, 5), $"{p}C-宫" },
                { (0, 9), $"{p}宝箱" },
            },
            Nodes =
            [
                // 主链（含A-8，覆盖基础的A-7→终章A-8）
                new()
                {
                    Code = $"{p}链A-1",
                    Pkg = "锚链",
                    Level = 1,
                    Typ = "链",
                    Res4d = 1
                },
                new()
                {
                    Code = $"{p}链A-2",
                    Pkg = "锚链",
                    Level = 2,
                    Typ = "链",
                    Res4d = 2,
                    ConsumeList = [new($"{p}链A-1", 1)]
                },
                new()
                {
                    Code = $"{p}链A-3",
                    Pkg = "锚链",
                    Level = 3,
                    Typ = "链",
                    Res4d = 3,
                    ConsumeList = [new($"{p}链A-2", 1)]
                },
                new()
                {
                    Code = $"{p}链A-4",
                    Pkg = "锚链",
                    Level = 4,
                    Typ = "链",
                    ConsumeList = [new($"{p}链A-3", 1)]
                },
                new()
                {
                    Code = $"{p}链A-5",
                    Pkg = "锚链",
                    Level = 5,
                    Typ = "链",
                    ConsumeList = [new($"{p}链A-4", 1)]
                },
                new()
                {
                    Code = $"{p}链A-6",
                    Pkg = "锚链",
                    Level = 6,
                    Typ = "链-3合",
                    ConsumeList = [new($"{p}链A-5", 1), new($"{p}B-A6", 1)]
                },
                new()
                {
                    Code = $"{p}B-A6",
                    Pkg = "链-建",
                    Typ = "链-建",
                    Count = 1,
                    ConsumeList = [new("体力", 30)]
                },
                new()
                {
                    Code = $"{p}链A-7",
                    Pkg = "锚链",
                    Level = 7,
                    Typ = "链",
                    ConsumeList = [new($"{p}链A-6", 1), new($"{p}B-A7", 1)]
                },
                new()
                {
                    Code = $"{p}B-A7",
                    Pkg = "链-建",
                    Typ = "链-建",
                    Count = 1,
                    ConsumeList = [new("体力", 40)]
                },
                new()
                {
                    Code = $"{p}链A-8",
                    Pkg = "锚链",
                    Level = 8,
                    Typ = "链-终-兑材料",
                    ConsumeList = [new($"{p}链A-7", 1), new($"{p}B-A8", 1)]
                },
                new()
                {
                    Code = $"{p}B-A8",
                    Pkg = "链-建",
                    Typ = "链-建",
                    Count = 1,
                    ConsumeList = [new("体力", 60)]
                },
                // 3条主材料线
                new()
                {
                    Code = $"{p}K-龙A",
                    Pkg = "龙鳞矿",
                    Typ = "矿",
                    Count = 4,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-龙A", 4)]
                },
                new()
                {
                    Code = $"{p}I-龙A",
                    Pkg = "龙鳞",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-龙B",
                    Pkg = "龙晶矿",
                    Typ = "矿",
                    Count = 4,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-龙B", 4)]
                },
                new()
                {
                    Code = $"{p}I-龙B",
                    Pkg = "龙晶",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}K-龙C",
                    Pkg = "龙焰矿",
                    Typ = "矿",
                    Count = 4,
                    ConsumeList = [new("体力", 20)],
                    ProduceList = [new($"{p}I-龙C", 4)]
                },
                new()
                {
                    Code = $"{p}I-龙C",
                    Pkg = "龙焰",
                    Typ = "兑-材料"
                },
                // 2条爆发线
                new()
                {
                    Code = $"{p}Burst-X",
                    Pkg = "爆发X",
                    Typ = "爆发",
                    Count = 1,
                    ConsumeList = [new("体力", 50)],
                    ProduceList = [new($"{p}I-爆X", 20)]
                },
                new()
                {
                    Code = $"{p}I-爆X",
                    Pkg = "爆材X",
                    Typ = "兑-材料"
                },
                new()
                {
                    Code = $"{p}Burst-Y",
                    Pkg = "爆发Y",
                    Typ = "爆发",
                    Count = 1,
                    ConsumeList = [new("体力", 50)],
                    ProduceList = [new($"{p}I-爆Y", 20)]
                },
                new()
                {
                    Code = $"{p}I-爆Y",
                    Pkg = "爆材Y",
                    Typ = "兑-材料"
                },
                // 龙心石
                new()
                {
                    Code = $"{p}D-龙",
                    Pkg = "龙心石",
                    Typ = "特殊",
                    ConsumeList =
                    [
                        new($"{p}I-龙A", 16),
                        new($"{p}I-龙B", 16),
                        new($"{p}I-龙C", 16),
                        new($"{p}I-爆X", 20)
                    ]
                },
                new()
                {
                    Code = $"{p}C-宫",
                    Pkg = "龙宫采集",
                    Typ = "采",
                    ProduceList =
                    [
                        new($"{p}链A-1", 15),
                        new($"{p}链A-2", 15),
                        new($"{p}链A-3", 20),
                        new($"{p}链A-4", 30)
                    ]
                },
                new()
                {
                    Code = $"{p}宝箱",
                    Pkg = "龙藏宝箱",
                    Typ = "地标-体力",
                    ProduceList = [new("活1", 20)]
                },
            ],
        };
    }

    // ── 公共入口 ─────────────────────────────────────────────────────────────

    public static void RunAll(string xlsxPath)
    {
        for (int i = 1; i <= 10; i++)
            Run(xlsxPath, i);
    }

    public static void Run(string xlsxPath, int mapN)
    {
        var data = mapN switch
        {
            1 => BuildMap1(),
            2 => BuildMap2(),
            3 => BuildMap3(),
            4 => BuildMap4(),
            5 => BuildMap5(),
            6 => BuildMap6(),
            7 => BuildMap7(),
            8 => BuildMap8(),
            9 => BuildMap9(),
            10 => BuildMap10(),
            _ => throw new ArgumentOutOfRangeException(nameof(mapN), "仅支持1~10")
        };
        WriteMap(xlsxPath, data);
    }

    // ── 写入核心 ─────────────────────────────────────────────────────────────

    private static void WriteMap(string xlsxPath, MapData data)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var pkg = new ExcelPackage(new FileInfo(xlsxPath));
        var ws = pkg.Workbook.Worksheets[0];

        var mapLabel = $"图{data.MapN}";

        // 找起止行
        int startRow = 0,
            endRow = 0;
        int totalRows = ws.Dimension?.Rows ?? 0;
        for (int r = 1; r <= totalRows; r++)
        {
            if (ws.Cells[r, 1].Value?.ToString() != mapLabel)
                continue;
            if (startRow == 0)
                startRow = r;
            endRow = r;
        }
        if (startRow == 0)
            throw new InvalidOperationException($"找不到{mapLabel}行");

        int oldCount = endRow - startRow + 1;
        Console.WriteLine($"{mapLabel}: 原{oldCount}行 -> 新{Target}行, 起始行={startRow}");

        if (Target > oldCount)
            ws.InsertRow(endRow + 1, Target - oldCount);
        else if (Target < oldCount)
            ws.DeleteRow(startRow + Target, oldCount - Target);

        // 清空区域
        for (int ri = startRow; ri < startRow + Target; ri++)
        for (int ci = 1; ci <= 70; ci++)
        {
            var cell = ws.Cells[ri, ci];
            cell.Value = null;
            cell.StyleID = 0;
        }

        // 读图N-1最大gid（col52）用于续号
        long lastGid = 0;
        if (data.MapN > 1)
        {
            var prevLabel = $"图{data.MapN - 1}";
            for (int r = 1; r <= ws.Dimension?.Rows; r++)
            {
                if (ws.Cells[r, 1].Value?.ToString() != prevLabel)
                    continue;
                if (long.TryParse(ws.Cells[r, 52].Value?.ToString(), out var v) && v > lastGid)
                    lastGid = v;
            }
        }
        long gidCounter = lastGid;
        long NextGid() => ++gidCounter;

        var nodeMap = data.Nodes.ToDictionary(n => n.Code);
        var rows = new List<NodeDef?>(data.Nodes.Cast<NodeDef?>());
        while (rows.Count < Target)
            rows.Add(null);

        for (int ri = 0; ri < Target; ri++)
        {
            int excelRow = startRow + ri;
            var nd = rows[ri];

            ws.Cells[excelRow, 1].Value = mapLabel;

            if (nd is null)
            {
                ws.Cells[excelRow, 50].Value = "通用";
                continue;
            }

            var gid = NextGid();
            var fullId = PeriodId * 10000 + gid;

            ws.Cells[excelRow, 39].Value = fullId;
            ws.Cells[excelRow, 40].Value = fullId;
            ws.Cells[excelRow, 41].Value = nd.Count;
            ws.Cells[excelRow, 42].Value = fullId;

            if (nd.Res4d.HasValue)
            {
                var rid = ResBase * 10000 + nd.Res4d.Value;
                ws.Cells[excelRow, 43].Value = rid;
                ws.Cells[excelRow, 44].Value = rid;
            }

            ws.Cells[excelRow, 45].Value = mapLabel;
            ws.Cells[excelRow, 46].Value = nd.Code;
            ws.Cells[excelRow, 47].Value = nd.Code;
            ws.Cells[excelRow, 48].Value = nd.Pkg;
            if (nd.Level.HasValue)
                ws.Cells[excelRow, 49].Value = nd.Level.Value;
            ws.Cells[excelRow, 50].Value = nd.Typ;
            ws.Cells[excelRow, 51].Value = $"{nd.Typ}-{nd.Code}";
            ws.Cells[excelRow, 52].Value = gid;

            for (int i = 0; i < Math.Min(nd.ConsumeList.Count, 4); i++)
            {
                ws.Cells[excelRow, 53 + i * 2].Value = nd.ConsumeList[i].Item;
                ws.Cells[excelRow, 54 + i * 2].Value = nd.ConsumeList[i].Qty;
            }
            for (int i = 0; i < Math.Min(nd.ProduceList.Count, 4); i++)
            {
                ws.Cells[excelRow, 61 + i * 2].Value = nd.ProduceList[i].Item;
                ws.Cells[excelRow, 62 + i * 2].Value = nd.ProduceList[i].Qty;
            }

            // 地编格（col7~col16）
            var pos = data.Grid.FirstOrDefault(kv => kv.Value == nd.Code);
            if (pos.Value == nd.Code)
            {
                int gridRow = pos.Key.R;
                for (int gc = 0; gc < GridCols; gc++)
                {
                    int excelCol = 7 + gc;
                    var cell = ws.Cells[excelRow, excelCol];
                    if (data.Grid.TryGetValue((gridRow, gc), out var cellCode))
                    {
                        var typ2 = nodeMap.TryGetValue(cellCode, out var nd2) ? nd2.Typ : "兑-材料";
                        var shortName = cellCode.StartsWith(mapLabel)
                            ? cellCode[mapLabel.Length..]
                            : cellCode;
                        cell.Value = shortName;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(BgColor(typ2));
                        cell.Style.Font.Color.SetColor(FgColor(typ2));
                        cell.Style.Font.Size = 8;
                        cell.Style.Font.Bold = true;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    else
                    {
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(EmptyBg);
                    }
                }
            }
        }

        // 列宽行高
        for (int col = 7; col <= 16; col++)
            ws.Column(col).Width = 13.0;
        for (int r = startRow; r < startRow + Target; r++)
            ws.Row(r).Height = 14.25;

        pkg.Save();

        int validCount = rows.Count(n => n is not null);
        int genericCount = rows.Count(n => n is null);
        Console.WriteLine($"  OK: 有效节点={validCount}, 通用格={genericCount}");
    }
}
