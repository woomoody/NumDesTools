using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// 通用云雾地图写入器，适用于所有以 mapCloud*.json 为数据源的活动。
/// 每个活动生成一个 xlsx，含「总览」sheet + 每个 Cloud 一个 sheet。
/// </summary>
public static class CloudMapWriter
{

    private const string DecryptedJsonDir = @"C:\tmp\mftown\decrypted_json";

    private const double CellColWidth = 4.0;
    private const double RowNumColWidth = 5.7;
    private const float RowHeight = 15f;
    private const float FontSize = 8f;

    // ── 通用物品前缀颜色（覆盖所有活动）
    private static readonly (string Prefix, string Color)[] ItemColors =
    [
        // 限时活动通用（EventSh/Mayqueen/Mom 等 eh4xx 系列）
        ("limitEventEnergy", "FF9999"), // 能量限制块
        ("eventletter", "99CCFF"), // 信件
        ("inreh4", "CCFFCC"), // 初始解锁
        ("ehgo4", "FFFF99"), // 可收获
        ("tbeh4", "FFD9AA"), // 任务物品
        ("obeh4", "DDDDDD"), // 障碍
        ("obfc4", "DDDDDD"), // 障碍(obfc413等)
        ("eh4", "FFFF99"), // 可收获主体
        // 主线 / 剧情通用
        ("inr", "CCFFCC"), // 初始解锁
        ("tbocopt", "FFE0CC"), // 海岛可选任务
        ("tboc", "FFD9AA"), // 海岛任务
        ("tb", "FFD9AA"), // 任务物品
        ("sdoc", "FF9999"), // 海岛特殊障碍
        ("ob", "DDDDDD"), // 障碍
        ("ocean", "99CCFF"), // 海岛主资源
        // 邮差剧情
        ("st1", "B3D9FF"), // 邮差1主元素
        ("st2", "B3D9FF"), // 邮差2主元素
        ("st3", "B3D9FF"), // 邮差3主元素
        ("st4", "B3D9FF"), // 邮差4主元素
        ("ruby", "FF99CC"), // 红宝石
        ("mainenergy", "FF9999"), // 能量限制
        // 主线地图
        ("gem", "D4AAFF"), // 宝石
        ("hs", "AAE6FF"), // 特殊物品
        ("supergnome", "FFD966"), // 超级地精
        ("everyth", "FFB3B3"), // 万物
        ("remn", "C8F5F5"), // 残留物
        ("magicstar", "FFD9AA"), // 魔法星
        ("pearl", "F5F5DC"), // 珍珠
        // 通用
        ("gold", "FFE566"), // 金币
        ("null", "FFFFFF"), // 空格
    ];

    private static readonly (string Label, string Color)[] LegendItems =
    [
        ("初始解锁(inr/inreh4)", "CCFFCC"),
        ("可收获(eh4/ehgo4)", "FFFF99"),
        ("任务物品(tb/tboc/tbeh4)", "FFD9AA"),
        ("障碍(ob/obeh4/obfc)", "DDDDDD"),
        ("能量限制(limitEvent/mainenergy)", "FF9999"),
        ("信件(eventletter)", "99CCFF"),
        ("主资源(ocean/st*)", "B3D9FF"),
        ("宝石(gem)", "D4AAFF"),
        ("金币(gold)", "FFE566"),
        ("空格(null)", "FFFFFF"),
    ];

    // ── 活动配置 ─────────────────────────────────────────────────────────────

    public sealed class ActivityConfig
    {
        public string FileName { get; init; } = "";
        public string OutputName { get; init; } = "";
        public string? FirstAreaFile { get; init; }

        /// <summary>
        /// true = 按 eh4xx 物品字母分 Area sheet（EventSh/Mayqueen/Mom 等）
        /// false = 仅生成一张总览 sheet（Ocean/主线/邮差等）
        /// </summary>
        public bool AreaSheets { get; init; }

        /// <summary>
        /// 可选：元素关系 JSON 文件（producer chains / merge elements / obstacles 等）
        /// 存在时生成「元素关系」sheet
        /// </summary>
        public string? ElementDataFile { get; init; }
    }

    public static readonly ActivityConfig[] KnownActivities =
    [
        new()
        {
            FileName = "mapCloudOceanIsland.json",
            OutputName = "CC收集活动-海岛地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaOceanIsland.json",
        },
        new()
        {
            FileName = "mapCloudEpisodePostman1.json",
            OutputName = "CC收集活动-邮差第1章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEpisodePostman1.json",
        },
        new()
        {
            FileName = "mapCloudEpisodePostman2.json",
            OutputName = "CC收集活动-邮差第2章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEpisodePostman2.json",
        },
        new()
        {
            FileName = "mapCloudEpisodePostman3.json",
            OutputName = "CC收集活动-邮差第3章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEpisodePostman3.json",
        },
        new()
        {
            FileName = "mapCloudEpisodePostman4.json",
            OutputName = "CC收集活动-邮差第4章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEpisodePostman4.json",
        },
        new()
        {
            FileName = "mapCloudMain.json",
            OutputName = "CC收集活动-主线第1章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaMain.json",
        },
        new()
        {
            FileName = "mapCloudMain2.json",
            OutputName = "CC收集活动-主线第2章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaMain2.json",
        },
        new()
        {
            FileName = "mapCloudMain3.json",
            OutputName = "CC收集活动-主线第3章地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaMain3.json",
        },
        new()
        {
            FileName = "mapCloudEventSh.json",
            OutputName = "CC收集活动-岛屿疗愈(Island Retreat)地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEventSh.json",
            AreaSheets = true,
            ElementDataFile = "eventSh_element_data.json",
        },
        new()
        {
            FileName = "mapCloudEventMom.json",
            OutputName = "CC收集活动-妈妈的心愿(Mom's Wishes)地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEventMom.json",
            AreaSheets = true,
        },
        new()
        {
            FileName = "mapCloudEventMayqueen.json",
            OutputName = "CC收集活动-五月女王(The May Queen)地编信息.xlsx",
            FirstAreaFile = "mapFirstAreaEventMayqueen.json",
            AreaSheets = true,
        },
        new()
        {
            FileName = "mapCloudEventOnePiece.json",
            OutputName = "CC收集活动-寻宝海盗(A Pirate's Treasure)地编信息.xlsx",
            FirstAreaFile = null,
        },
    ];

    // ── 数据模型 ─────────────────────────────────────────────────────────────

    private sealed class CloudInfo
    {
        public string Name { get; set; } = "";
        public int Level { get; set; }
        public string Need { get; set; } = "";
        public string Reward { get; set; } = "";
        public HashSet<string> Coords { get; set; } = [];
        public int TileCount => Coords.Count;
        public int MinR,
            MaxR,
            MinC,
            MaxC;
    }

    private sealed class MapData
    {
        public Dictionary<string, string> AllTiles { get; set; } = [];
        public HashSet<string> FirstArea { get; set; } = [];
        public List<CloudInfo> Clouds { get; set; } = [];

        // AreaSheets 模式专用
        public Dictionary<string, char> CoordToArea { get; set; } = [];
        public Dictionary<char, (int MinR, int MaxR, int MinC, int MaxC)> AreaBounds { get; set; } =
            [];

        public int MinR,
            MaxR,
            MinC,
            MaxC;
    }

    // ── 公共入口：运行所有已知活动 ────────────────────────────────────────────

    public static void RunAll(string? outputDir = null, string? dataDir = null)
    {
        var dir = dataDir ?? DecryptedJsonDir;
        foreach (var cfg in KnownActivities)
        {
            var dataPath = Path.Combine(dir, cfg.FileName);
            if (!File.Exists(dataPath))
            {
                Console.WriteLine($"[CloudMap] 跳过（文件不存在）：{dataPath}");
                continue;
            }
            Run(cfg, outputDir, dataDir);
        }
    }

    // ── 公共入口：运行单个活动 ────────────────────────────────────────────────

    public static void Run(ActivityConfig cfg, string? outputDir = null, string? dataDir = null)
    {
        var dir = dataDir ?? DecryptedJsonDir;
        var outDir = outputDir ?? OutputPaths.Reports;
        var dataPath = Path.Combine(dir, cfg.FileName);
        var faPath = cfg.FirstAreaFile is not null ? Path.Combine(dir, cfg.FirstAreaFile) : null;
        var outputPath = Path.Combine(outDir, cfg.OutputName);

        var data = LoadData(dataPath, faPath);
        Console.WriteLine(
            $"[CloudMap] {cfg.OutputName}: 全图 {data.MaxR - data.MinR + 1}行×{data.MaxC - data.MinC + 1}列  "
                + $"({data.MinR}~{data.MaxR}, {data.MinC}~{data.MaxC})  Clouds={data.Clouds.Count}"
        );

        using var pkg = new ExcelPackage();

        if (cfg.AreaSheets && data.AreaBounds.Count > 0)
        {
            foreach (var letter in data.AreaBounds.Keys.Order())
            {
                var (minR, maxR, minC, maxC) = data.AreaBounds[letter];
                BuildAreaSheet(pkg, $"Area{letter}", data, minR, maxR, minC, maxC, letter);
            }
        }
        else
        {
            BuildFullMapSheet(pkg, data);
        }

        BuildCloudsSheet(pkg, data);
        BuildItemDictSheet(pkg, data);

        if (cfg.ElementDataFile is not null)
        {
            var elPath = Path.Combine(dir, cfg.ElementDataFile);
            if (File.Exists(elPath))
                BuildElementRelationsSheet(pkg, elPath);
        }

        pkg.SaveAs(new FileInfo(outputPath));
        Console.WriteLine($"[CloudMap] 已写入：{outputPath}");
        OutputPaths.GitCommit($"[CloudMap] 更新云雾地图报告 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── 加载数据 ─────────────────────────────────────────────────────────────

    private static MapData LoadData(string dataPath, string? firstAreaPath)
    {
        using var f = File.OpenRead(dataPath);
        var rawClouds =
            JsonSerializer.Deserialize<JsonElement[]>(f)
            ?? throw new InvalidDataException($"无法解析 JSON：{dataPath}");

        var allTiles = new Dictionary<string, string>();
        var clouds = new List<CloudInfo>();

        foreach (var raw in rawClouds)
        {
            var name = raw.TryGetProperty("name", out var nv) ? nv.GetString() ?? "" : "";
            var level = raw.TryGetProperty("level", out var lv) ? lv.GetInt32() : 0;
            var need = raw.TryGetProperty("need", out var needProp)
                ? (
                    needProp.ValueKind == JsonValueKind.Array
                        ? string.Join(
                            ", ",
                            needProp.EnumerateArray().Select(x => x.GetString() ?? "")
                        )
                        : needProp.GetString() ?? ""
                )
                : "";
            var reward = raw.TryGetProperty("reward", out var rewardProp)
                ? (
                    rewardProp.ValueKind == JsonValueKind.Array
                        ? string.Join(
                            ", ",
                            rewardProp.EnumerateArray().Select(x => x.GetString() ?? "")
                        )
                        : rewardProp.GetString() ?? ""
                )
                : "";

            var coords = new HashSet<string>();
            if (raw.TryGetProperty("area", out var areaProp))
                foreach (var tile in areaProp.EnumerateArray())
                {
                    var parts = tile.GetString()?.Split('#');
                    if (parts?.Length == 2)
                    {
                        allTiles[parts[0]] = parts[1];
                        coords.Add(parts[0]);
                    }
                }

            if (coords.Count == 0)
                continue;

            var rs = coords.Select(c => int.Parse(c.Split('-')[0])).ToList();
            var cs = coords.Select(c => int.Parse(c.Split('-')[1])).ToList();
            clouds.Add(
                new CloudInfo
                {
                    Name = name,
                    Level = level,
                    Need = need,
                    Reward = reward,
                    Coords = coords,
                    MinR = rs.Min(),
                    MaxR = rs.Max(),
                    MinC = cs.Min(),
                    MaxC = cs.Max(),
                }
            );
        }

        // firstArea：coord#item 列表（根可能是 array 包裹的 object，也可能直接是 object）
        var firstArea = new HashSet<string>();
        if (firstAreaPath is not null && File.Exists(firstAreaPath))
        {
            using var ff = File.OpenRead(firstAreaPath);
            var faDoc = JsonSerializer.Deserialize<JsonElement>(ff);
            JsonElement faObj =
                faDoc.ValueKind == JsonValueKind.Array && faDoc.GetArrayLength() > 0
                    ? faDoc[0]
                    : faDoc;
            if (faObj.TryGetProperty("firstArea", out var faProp))
            {
                if (faProp.ValueKind == JsonValueKind.Array)
                    foreach (var item in faProp.EnumerateArray())
                    {
                        var s = item.GetString();
                        if (s is not null)
                            firstArea.Add(s.Split('#')[0]);
                    }
                else if (faProp.ValueKind == JsonValueKind.Object)
                    foreach (var kv in faProp.EnumerateObject())
                        firstArea.Add(kv.Name);
            }
        }

        if (allTiles.Count == 0)
            return new MapData { Clouds = clouds, FirstArea = firstArea };

        var allR = allTiles.Keys.Select(k => int.Parse(k.Split('-')[0])).ToList();
        var allC = allTiles.Keys.Select(k => int.Parse(k.Split('-')[1])).ToList();

        // Area 检测：从 eh4xx 物品名中提取字母
        var coordToArea = BuildAreaMap(rawClouds);
        var areaBounds = new Dictionary<char, (int, int, int, int)>();
        foreach (var (coord, letter) in coordToArea)
        {
            var seg = coord.Split('-');
            int r = int.Parse(seg[0]),
                c = int.Parse(seg[1]);
            if (!areaBounds.TryGetValue(letter, out var b))
                areaBounds[letter] = (r, r, c, c);
            else
                areaBounds[letter] = (
                    Math.Min(b.Item1, r),
                    Math.Max(b.Item2, r),
                    Math.Min(b.Item3, c),
                    Math.Max(b.Item4, c)
                );
        }

        return new MapData
        {
            AllTiles = allTiles,
            FirstArea = firstArea,
            Clouds = clouds,
            CoordToArea = coordToArea,
            AreaBounds = areaBounds,
            MinR = allR.Min(),
            MaxR = allR.Max(),
            MinC = allC.Min(),
            MaxC = allC.Max(),
        };
    }

    private static Dictionary<string, char> BuildAreaMap(JsonElement[] rawClouds)
    {
        var coordToArea = new Dictionary<string, char>();
        foreach (var raw in rawClouds)
        {
            if (!raw.TryGetProperty("area", out var areaProp))
                continue;
            var letters = new List<char>();
            var tileCoords = new List<string>();
            foreach (var tile in areaProp.EnumerateArray())
            {
                var parts = tile.GetString()?.Split('#');
                if (parts?.Length != 2)
                    continue;
                tileCoords.Add(parts[0]);
                var m = Regex.Match(
                    parts[1],
                    @"^(?:inr|ob|ehgo|tbeh|eh)?eh4\d+_(?:plant_|bomb_|crop_|limit_)?([a-z])"
                );
                if (m.Success)
                    letters.Add(char.ToUpper(m.Groups[1].Value[0]));
            }
            if (letters.Count == 0)
                continue;
            var dominant = letters.GroupBy(x => x).MaxBy(g => g.Count())!.Key;
            foreach (var coord in tileCoords)
                coordToArea[coord] = dominant;
        }
        return coordToArea;
    }

    // ── Area 分 sheet（AreaSheets 模式）─────────────────────────────────────────

    private static readonly Dictionary<char, string> AreaColors =
        new()
        {
            { 'A', "FFD9D9" },
            { 'B', "FFEDD9" },
            { 'C', "FFFFD9" },
            { 'D', "D9F2D9" },
            { 'E', "D9EDFF" },
            { 'F', "EAD9FF" },
            { 'G', "FFD4EF" },
            { 'H', "C8F5F5" },
            { 'I', "F5E3C8" },
            { 'J', "E8FFE8" },
            { 'P', "E0E0E0" },
        };

    private static void BuildAreaSheet(
        ExcelPackage pkg,
        string sheetName,
        MapData data,
        int minR,
        int maxR,
        int minC,
        int maxC,
        char areaLetter
    )
    {
        int nRows = maxR - minR + 1;
        int nCols = maxC - minC + 1;
        int lgCol = nCols + 3;

        Console.WriteLine($"  [{sheetName}] {nRows}×{nCols}");
        var ws = pkg.Workbook.Worksheets.Add(sheetName);

        ws.Cells[1, 1].Value = "行\\列";
        for (int ci = 0; ci < nCols; ci++)
            ws.Cells[1, ci + 2].Value = minC + ci;
        ws.Cells[1, lgCol].Value = "图例";
        ws.Cells[1, lgCol + 1].Value = "说明";

        for (int ri = 0; ri < nRows; ri++)
        {
            int gameR = minR + ri;
            int sheetRow = ri + 2;
            ws.Cells[sheetRow, 1].Value = gameR;

            for (int ci = 0; ci < nCols; ci++)
            {
                int gameC = minC + ci;
                var coord = $"{gameR}-{gameC}";
                var item = data.AllTiles.GetValueOrDefault(coord, "");
                var cell = ws.Cells[sheetRow, ci + 2];

                if (!string.IsNullOrEmpty(item))
                {
                    cell.Value = ItemLabel(item);
                    cell.Style.Font.Size = FontSize;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    var color = ItemColor(item);
                    if (color is not null)
                        SetBg(cell, color);
                    if (data.FirstArea.Contains(coord))
                        cell.Style.Font.Bold = true;
                }
                else if (data.CoordToArea.GetValueOrDefault(coord) == areaLetter)
                {
                    SetBg(cell, AreaColors.GetValueOrDefault(areaLetter, "EEEEEE"));
                }
            }

            if (ri < LegendItems.Length)
            {
                var (lgLabel, lgColor) = LegendItems[ri];
                var lgCell = ws.Cells[sheetRow, lgCol];
                lgCell.Value = "■";
                lgCell.Style.Font.Bold = true;
                SetBg(lgCell, lgColor);
                ws.Cells[sheetRow, lgCol + 1].Value = lgLabel;
            }
        }

        ws.Column(1).Width = RowNumColWidth;
        for (int ci = 2; ci <= nCols + 1; ci++)
            ws.Column(ci).Width = CellColWidth;
        for (int r = 1; r <= nRows + 1; r++)
            ws.Row(r).Height = RowHeight;
    }

    // ── 整张地图（所有 Cloud 合并，用区域色区分）────────────────────────────────

    private static readonly string[] CloudPalette =
    [
        "FFD9D9",
        "FFEDD9",
        "FFFFD9",
        "D9F2D9",
        "D9EDFF",
        "EAD9FF",
        "FFD4EF",
        "C8F5F5",
        "F5E3C8",
        "E0E0E0",
    ];

    private static void BuildFullMapSheet(ExcelPackage pkg, MapData data)
    {
        int nRows = data.MaxR - data.MinR + 1;
        int nCols = data.MaxC - data.MinC + 1;
        int lgCol = nCols + 3;

        Console.WriteLine($"  [地图] {nRows}×{nCols}");
        var ws = pkg.Workbook.Worksheets.Add("地图");

        // 坐标 → Cloud，Cloud → 背景色
        var coordToCloud = new Dictionary<string, CloudInfo>();
        foreach (var cloud in data.Clouds)
        foreach (var coord in cloud.Coords)
            coordToCloud[coord] = cloud;

        var cloudColors = data
            .Clouds.OrderBy(c => c.Level)
            .ThenBy(c => c.Name)
            .Select((c, i) => (c.Name, CloudPalette[i % CloudPalette.Length]))
            .ToDictionary(x => x.Name, x => x.Item2);

        // 标题行
        ws.Cells[1, 1].Value = "行\\列";
        for (int ci = 0; ci < nCols; ci++)
            ws.Cells[1, ci + 2].Value = data.MinC + ci;
        ws.Cells[1, lgCol].Value = "图例";
        ws.Cells[1, lgCol + 1].Value = "说明";

        for (int ri = 0; ri < nRows; ri++)
        {
            int gameR = data.MinR + ri;
            int sheetRow = ri + 2;
            ws.Cells[sheetRow, 1].Value = gameR;

            for (int ci = 0; ci < nCols; ci++)
            {
                var coord = $"{gameR}-{data.MinC + ci}";
                var item = data.AllTiles.GetValueOrDefault(coord, "");
                var cell = ws.Cells[sheetRow, ci + 2];

                if (!string.IsNullOrEmpty(item))
                {
                    cell.Value = ItemLabel(item);
                    cell.Style.Font.Size = FontSize;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    var itemColor = ItemColor(item);
                    if (itemColor is not null)
                        SetBg(cell, itemColor);

                    if (data.FirstArea.Contains(coord))
                        cell.Style.Font.Bold = true;
                }
                else if (coordToCloud.ContainsKey(coord))
                {
                    // 空格但属于某 Cloud → 叠加该 Cloud 的背景色
                    var bgColor = cloudColors.GetValueOrDefault(coordToCloud[coord].Name, "EEEEEE");
                    SetBg(cell, bgColor);
                }
            }

            if (ri < LegendItems.Length)
            {
                var (lgLabel, lgColor) = LegendItems[ri];
                var lgCell = ws.Cells[sheetRow, lgCol];
                lgCell.Value = "■";
                lgCell.Style.Font.Bold = true;
                SetBg(lgCell, lgColor);
                ws.Cells[sheetRow, lgCol + 1].Value = lgLabel;
            }
        }

        ws.Column(1).Width = RowNumColWidth;
        for (int ci = 2; ci <= nCols + 1; ci++)
            ws.Column(ci).Width = CellColWidth;
        for (int r = 1; r <= nRows + 1; r++)
            ws.Row(r).Height = RowHeight;
    }

    // ── 云雾区域汇总 sheet ────────────────────────────────────────────────────

    private static void BuildCloudsSheet(ExcelPackage pkg, MapData data)
    {
        Console.WriteLine($"  [云雾区域] {data.Clouds.Count} 条");
        var ws = pkg.Workbook.Worksheets.Add("云雾区域");

        string[] headers = ["云名", "等级", "格子数", "解锁条件(need)", "奖励(reward)"];
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cells[1, i + 1];
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            SetBg(cell, "D9D9D9");
        }

        var sorted = data.Clouds.OrderBy(c => c.Level).ThenBy(c => c.Name).ToList();
        for (int ri = 0; ri < sorted.Count; ri++)
        {
            var c = sorted[ri];
            int row = ri + 2;
            ws.Cells[row, 1].Value = c.Name;
            ws.Cells[row, 2].Value = c.Level;
            ws.Cells[row, 3].Value = c.TileCount;
            ws.Cells[row, 4].Value = c.Need;
            ws.Cells[row, 5].Value = c.Reward;
        }

        ws.Column(1).Width = 22;
        ws.Column(2).Width = 6;
        ws.Column(3).Width = 8;
        ws.Column(4).Width = 28;
        ws.Column(5).Width = 28;
        ws.Cells[ws.Dimension.Address].Style.Font.Size = FontSize;
    }

    // ── 元素字典 sheet（缩写 ↔ 全名对照）────────────────────────────────────────

    // ── 元素关系 sheet ────────────────────────────────────────────────────────

    private static void BuildElementRelationsSheet(ExcelPackage pkg, string jsonPath)
    {
        using var f = File.OpenRead(jsonPath);
        var root = JsonSerializer.Deserialize<JsonElement>(f);
        Console.WriteLine($"  [元素关系] {Path.GetFileName(jsonPath)}");

        var ws = pkg.Workbook.Worksheets.Add("元素关系");
        int row = 1;

        // — 生产者升级链
        if (
            root.TryGetProperty("producer_chains", out var chains)
            && chains.ValueKind == JsonValueKind.Object
        )
        {
            WriteRelHeader(ws, row++, "生产者升级链(producer_chains)", "CCFFCC");
            WriteRelBold(ws, row++, ["基础元素", "升级链(Lv1→Lv4)"]);
            foreach (var kv in chains.EnumerateObject())
            {
                var chain = kv.Value.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
                ws.Cells[row, 1].Value = kv.Name;
                ws.Cells[row, 2].Value = string.Join(" → ", chain);
                row++;
            }
            row++;
        }

        // — 合并元素（按区域字母）
        if (root.TryGetProperty("merge_elements", out var mergeEls))
        {
            WriteRelHeader(ws, row++, "合并元素(merge_elements)", "FFFF99");
            var els = mergeEls.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
            foreach (var el in els)
            {
                ws.Cells[row, 1].Value = el;
                ws.Cells[row, 1].Style.Font.Size = FontSize;
                row++;
            }
            row++;
        }

        // — 初始解锁元素
        if (root.TryGetProperty("inr_elements", out var inrEls))
        {
            WriteRelHeader(ws, row++, "初始解锁元素(inr_elements)", "CCFFCC");
            foreach (var el in inrEls.EnumerateArray())
            {
                ws.Cells[row, 1].Value = el.GetString();
                ws.Cells[row, 1].Style.Font.Size = FontSize;
                row++;
            }
            row++;
        }

        // — 信件元素
        if (root.TryGetProperty("event_letters", out var letters) && letters.GetArrayLength() > 0)
        {
            WriteRelHeader(ws, row++, "信件元素(event_letters)", "99CCFF");
            foreach (var el in letters.EnumerateArray())
            {
                ws.Cells[row, 1].Value = el.GetString();
                row++;
            }
            row++;
        }

        // — 能量限制块
        if (root.TryGetProperty("energy", out var energy) && energy.GetArrayLength() > 0)
        {
            WriteRelHeader(ws, row++, "能量限制块(limitEventEnergy)", "FF9999");
            foreach (var el in energy.EnumerateArray())
            {
                ws.Cells[row, 1].Value = el.GetString();
                row++;
            }
            row++;
        }

        // — 障碍元素
        if (root.TryGetProperty("obstacles", out var obs))
        {
            WriteRelHeader(ws, row++, "障碍元素(obstacles)", "DDDDDD");
            foreach (var el in obs.EnumerateArray())
            {
                ws.Cells[row, 1].Value = el.GetString();
                ws.Cells[row, 1].Style.Font.Size = FontSize;
                row++;
            }
            row++;
        }

        ws.Column(1).Width = 36;
        ws.Column(2).Width = 80;
        if (ws.Dimension is not null)
            ws.Cells[ws.Dimension.Address].Style.Font.Size = FontSize;
    }

    private static void WriteRelHeader(ExcelWorksheet ws, int row, string title, string color)
    {
        var cell = ws.Cells[row, 1];
        cell.Value = title;
        cell.Style.Font.Bold = true;
        SetBg(cell, color);
    }

    private static void WriteRelBold(ExcelWorksheet ws, int row, string[] cols)
    {
        for (int i = 0; i < cols.Length; i++)
        {
            ws.Cells[row, i + 1].Value = cols[i];
            ws.Cells[row, i + 1].Style.Font.Bold = true;
        }
    }

    private static void BuildItemDictSheet(ExcelPackage pkg, MapData data)
    {
        // 统计所有出现过的 tile 类型及次数
        var typeCounts = new Dictionary<string, int>();
        foreach (var (coord, typ) in data.AllTiles)
        {
            if (string.IsNullOrEmpty(typ) || typ == "null")
                continue;
            typeCounts[typ] = typeCounts.GetValueOrDefault(typ) + 1;
        }
        // firstArea 里的类型也计入
        foreach (var coord in data.FirstArea)
        {
            if (data.AllTiles.TryGetValue(coord, out var typ) && !string.IsNullOrEmpty(typ))
                typeCounts[typ] = typeCounts.GetValueOrDefault(typ); // already counted above
        }

        if (typeCounts.Count == 0)
            return;

        Console.WriteLine($"  [元素字典] {typeCounts.Count} 种");
        var ws = pkg.Workbook.Worksheets.Add("元素字典");

        // 按类型前缀分组排序
        string TypeCategory(string t) =>
            t switch
            {
                _ when t.StartsWith("limitEventEnergy") => "0_energy",
                _ when t.StartsWith("eventletter") => "1_letter",
                _ when t.StartsWith("inreh4") => "2_inr",
                _ when t.StartsWith("ehgo4") => "3_ehgo",
                _ when t.StartsWith("tbeh4") => "4_tb",
                _ when t.StartsWith("eh4") => "5_eh",
                _ when t.StartsWith("obeh4") => "6_obeh",
                _ when t.StartsWith("obfc4") => "7_obfc4",
                _ when t.StartsWith("obfc") => "7_obfc",
                _ when t.StartsWith("ob") => "7_ob",
                _ when t.StartsWith("gold") => "8_gold",
                _ when t.StartsWith("replace") => "9_replace",
                _ => "z_" + t[..Math.Min(3, t.Length)],
            };

        string[] headers = ["缩写(地图中显示)", "完整名称", "类型分类", "出现次数", "颜色"];
        for (int i = 0; i < headers.Length; i++)
        {
            var hc = ws.Cells[1, i + 1];
            hc.Value = headers[i];
            hc.Style.Font.Bold = true;
            SetBg(hc, "D9D9D9");
        }

        var sorted = typeCounts.OrderBy(kv => TypeCategory(kv.Key)).ThenBy(kv => kv.Key).ToList();

        for (int ri = 0; ri < sorted.Count; ri++)
        {
            var (typ, cnt) = sorted[ri];
            int row = ri + 2;

            var label = ItemLabel(typ);
            var color = ItemColor(typ) ?? "F5F5F5";
            var category = TypeCategory(typ).TrimStart("0123456789_".ToCharArray());

            var labelCell = ws.Cells[row, 1];
            labelCell.Value = label;
            labelCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            SetBg(labelCell, color);

            ws.Cells[row, 2].Value = typ;
            ws.Cells[row, 3].Value = category;
            ws.Cells[row, 4].Value = cnt;

            var colorCell = ws.Cells[row, 5];
            colorCell.Value = "■";
            colorCell.Style.Font.Bold = true;
            colorCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            SetBg(colorCell, color);
        }

        ws.Column(1).Width = 10;
        ws.Column(2).Width = 36;
        ws.Column(3).Width = 14;
        ws.Column(4).Width = 10;
        ws.Column(5).Width = 6;
        ws.Cells[ws.Dimension.Address].Style.Font.Size = FontSize;
    }

    // ── 辅助 ─────────────────────────────────────────────────────────────────

    private static string? ItemColor(string item)
    {
        foreach (var (prefix, color) in ItemColors)
            if (item.StartsWith(prefix))
                return color;
        return "F5F5F5";
    }

    private static string ItemLabel(string item)
    {
        if (string.IsNullOrEmpty(item) || item == "null")
            return "·";
        if (item.StartsWith("limitEventEnergy"))
            return "E";
        if (item.StartsWith("eventletter"))
            return "L";
        if (item.StartsWith("mainenergy"))
            return "E";
        if (item.StartsWith("gold"))
            return "G";
        if (item.StartsWith("ruby"))
            return "R";
        if (item.StartsWith("sdoc"))
            return "S";
        // eh4xx 系列（inreh413_j1_1 → j1，obeh413_g21 → g21，eh413_plant_i3 → pi3）
        var m = Regex.Match(
            item,
            @"^(?:inr|ob|ehgo|tbeh|eh)?eh4\d+_(?:plant_|bomb_|crop_|limit_)?([a-z])(\w+?)(?:_\d+)?$"
        );
        if (m.Success)
            return m.Groups[1].Value + m.Groups[2].Value[..Math.Min(2, m.Groups[2].Value.Length)];
        // obfc63_1 → fc1
        m = Regex.Match(item, @"^obfc(\d+)_(\d+)$");
        if (m.Success)
            return "f" + m.Groups[2].Value;
        // ocean_N_M → N
        m = Regex.Match(item, @"^ocean_(\d+)_(\d+)$");
        if (m.Success)
            return m.Groups[1].Value;
        // tboc/tbocopt → 前缀简称
        m = Regex.Match(item, @"^tboc(?:opt)?_([a-z]+)_(\d+)$");
        if (m.Success)
            return m.Groups[1].Value[..Math.Min(2, m.Groups[1].Value.Length)] + m.Groups[2].Value;
        // inrXX_area_N / stN_area_N 类通用：取下划线分段
        m = Regex.Match(item, @"^(?:inr|tb|ob)?(\w+?)_(\d+)(?:_\d+)?$");
        if (m.Success)
            return m.Groups[1].Value[..Math.Min(2, m.Groups[1].Value.Length)] + m.Groups[2].Value;
        return item[..Math.Min(3, item.Length)];
    }

    private static string SanitizeSheetName(string name)
    {
        // Excel sheet名不能超过31字符，不能含 : \ / ? * [ ]
        var clean = Regex.Replace(name, @"[:\\/?*\[\]]", "_");
        return clean.Length > 31 ? clean[..31] : clean;
    }

    private static void SetBg(ExcelRange cell, string hex6)
    {
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(HexColor(hex6));
    }

    private static Color HexColor(string h)
    {
        var v = Convert.ToInt32(h.TrimStart('#'), 16);
        return Color.FromArgb((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
    }
}
