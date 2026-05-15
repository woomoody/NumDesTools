using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// 果酱节地图可视化写入器。
/// 数据源：C:/tmp/mftown/analysis/jam_config_full.json
/// 输出：每次调用生成一个 xlsx，含「总览」sheet + 每个 Area 一个 sheet。
/// 格式：列宽 28px≈4字符，行高 15pt，字体 8pt，居中，物品染色，firstArea 加粗，图例。
/// </summary>
public static class JamMapWriter
{
    private const string DefaultDataPath = @"C:\tmp\mftown\analysis\jam_config_full.json";
    private const string DefaultGameTextPath = @"C:\tmp\mftown\game_text_en.json";
    private const string DefaultElementPath = @"C:\tmp\mftown\jam_element_analysis.json";
    private const string DefaultProduceConsumePath = @"C:\tmp\mftown\jam_produce_consume.json";

    // ── 尺寸常量（openpyxl 到 EPPlus 换算：px/7 ≈ 字符宽，1pt=1磅）
    private const double CellColWidth = 4.0; // 28px ≈ 4 字符
    private const double RowNumColWidth = 5.7; // 40px ≈ 5.7 字符
    private const float RowHeight = 15f; // 20px ≈ 15pt
    private const float FontSize = 8f;

    // ── Area 背景色
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
            { 'P', "E0E0E0" },
        };

    // ── 物品前缀颜色
    private static readonly (string Prefix, string Color)[] ItemColors =
    [
        ("limitEventEnergy", "FF9999"),
        ("eventletter", "99CCFF"),
        ("inreh418", "CCFFCC"),
        ("ehgo418", "FFFF99"),
        ("eh418", "FFFF99"),
        ("obeh418", "DDDDDD"),
        ("obfc", "DDDDDD"),
        ("tbeh418", "FFD9AA"),
        ("gold", "FFE566"),
        ("null", "FFFFFF"),
    ];

    // ── 图例
    private static readonly (string Label, string Color)[] LegendItems =
    [
        ("初始解锁(inr)", "CCFFCC"),
        ("可收获(eh/go)", "FFFF99"),
        ("任务物品(tb)", "FFD9AA"),
        ("能量限制块", "FF9999"),
        ("信件(letter)", "99CCFF"),
        ("障碍(ob/obfc)", "DDDDDD"),
        ("金币(gold)", "FFE566"),
        ("空格(null)", "FFFFFF"),
    ];

    // ── 数据模型 ─────────────────────────────────────────────────────────────

    private sealed class CloudRecord
    {
        public string Name { get; set; } = "";
        public char AreaLetter { get; set; }
        public int TileCount { get; set; }
        public string Need { get; set; } = "";
        public string Reward { get; set; } = "";
    }

    private sealed class ElementData
    {
        public List<string> MergeElements { get; set; } = [];
        public List<string> Producers { get; set; } = [];
        public List<string> Ingredients { get; set; } = [];
        public List<string> Obstacles { get; set; } = [];
    }

    private sealed class ProduceConsumeData
    {
        public Dictionary<string, List<string>> ProducerChains { get; set; } = [];
        public List<RecipeRecord> Recipes { get; set; } = [];
        public List<StoryTask> StoryTasks { get; set; } = [];
        public List<string> UnlockEvents { get; set; } = [];
        public Dictionary<string, string> AreaRewards { get; set; } = [];
    }

    private sealed class RecipeRecord
    {
        public int Id { get; set; }
        public string Needs { get; set; } = "";
        public string Produce { get; set; } = "";
    }

    private sealed class StoryTask
    {
        public string Id { get; set; } = "";
        public string Entity { get; set; } = "";
        public string Target { get; set; } = "";
    }

    private sealed class JamData
    {
        public Dictionary<string, string> AllTiles { get; set; } = [];
        public HashSet<string> FirstArea { get; set; } = [];
        public Dictionary<string, char> CoordToArea { get; set; } = [];
        public Dictionary<char, (int MinR, int MaxR, int MinC, int MaxC)> AreaBounds { get; set; } =
            [];
        public Dictionary<char, string> AreaNames { get; set; } = [];
        public List<CloudRecord> Clouds { get; set; } = [];
        public ElementData Elements { get; set; } = new();
        public ProduceConsumeData ProduceConsume { get; set; } = new();
        public int MinR,
            MaxR,
            MinC,
            MaxC;
    }

    // ── 公共入口 ─────────────────────────────────────────────────────────────

    public static void Run(
        string? outputDir = null,
        string dataPath = DefaultDataPath,
        string gameTextPath = DefaultGameTextPath,
        string elementPath = DefaultElementPath,
        string produceConsumePath = DefaultProduceConsumePath
    )
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var dir = outputDir ?? OutputPaths.Reports;
        var outputPath = Path.Combine(dir, "CC收集活动-果酱节地编信息.xlsx");

        var data = LoadData(dataPath, gameTextPath, elementPath, produceConsumePath);
        Console.WriteLine(
            $"[JamMap] 全图 {data.MaxR - data.MinR + 1}行×{data.MaxC - data.MinC + 1}列  "
                + $"({data.MinR}~{data.MaxR}, {data.MinC}~{data.MaxC})"
        );
        Console.WriteLine($"[JamMap] Areas: {string.Join(' ', data.AreaBounds.Keys.Order())}");

        using var pkg = new ExcelPackage();

        foreach (var letter in data.AreaBounds.Keys.Order())
        {
            var (minR, maxR, minC, maxC) = data.AreaBounds[letter];
            var name = data.AreaNames.GetValueOrDefault(letter, $"Area {letter}");
            BuildMapSheet(pkg, $"Area{letter}-{name}", data, minR, maxR, minC, maxC, letter);
        }

        BuildCloudsSheet(pkg, data);
        BuildItemStatsSheet(pkg, data);
        BuildElementRelationsSheet(pkg, data);

        pkg.SaveAs(new FileInfo(outputPath));
        Console.WriteLine($"[JamMap] 已写入：{outputPath}");
        OutputPaths.GitCommit($"[JamMap] 更新果酱节地编报告 {DateTime.Today:yyyy-MM-dd}");
    }

    // ── 加载数据 ─────────────────────────────────────────────────────────────

    private static JamData LoadData(
        string dataPath,
        string gameTextPath,
        string elementPath,
        string produceConsumePath
    )
    {
        using var df = File.OpenRead(dataPath);
        var root = JsonSerializer.Deserialize<JsonElement>(df);

        var allTiles = new Dictionary<string, string>();
        foreach (var kv in root.GetProperty("allTiles").EnumerateObject())
            allTiles[kv.Name] = kv.Value.GetString() ?? "";

        var firstArea = new HashSet<string>();
        foreach (var kv in root.GetProperty("firstArea").EnumerateObject())
            firstArea.Add(kv.Name);

        // coord → area letter + cloud records
        var coordToArea = new Dictionary<string, char>();
        var cloudRecords = new List<CloudRecord>();
        foreach (var cloud in root.GetProperty("clouds").EnumerateArray())
        {
            var letters = new List<char>();
            var tileCoords = new List<string>();
            foreach (var tile in cloud.GetProperty("area").EnumerateArray())
            {
                var parts = tile.GetString()?.Split('#');
                if (parts?.Length == 2)
                {
                    tileCoords.Add(parts[0]);
                    var m = Regex.Match(
                        parts[1],
                        @"^(?:obeh418_|eh418_|ehgo418_|inreh418_|tbeh418_)([a-z])"
                    );
                    if (m.Success)
                        letters.Add(char.ToUpper(m.Groups[1].Value[0]));
                }
            }
            if (letters.Count == 0)
                continue;

            var dominant = letters.GroupBy(x => x).MaxBy(g => g.Count())!.Key;
            foreach (var coord in tileCoords)
                coordToArea[coord] = dominant;

            var cloudName = cloud.TryGetProperty("name", out var nv) ? nv.GetString() ?? "" : "";
            var need = cloud.TryGetProperty("need", out var needArr)
                ? string.Join(", ", needArr.EnumerateArray().Select(x => x.GetString() ?? ""))
                : "";
            var reward = cloud.TryGetProperty("reward", out var rewardArr)
                ? string.Join(", ", rewardArr.EnumerateArray().Select(x => x.GetString() ?? ""))
                : "";
            cloudRecords.Add(
                new CloudRecord
                {
                    Name = cloudName,
                    AreaLetter = dominant,
                    TileCount = tileCoords.Count,
                    Need = need,
                    Reward = reward,
                }
            );
        }

        // area bounds
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

        // area names from game_text
        var areaNames = new Dictionary<char, string>();
        if (File.Exists(gameTextPath))
        {
            using var gf = File.OpenRead(gameTextPath);
            var gt = JsonSerializer.Deserialize<Dictionary<string, string>>(gf) ?? [];
            for (int i = 1; i <= 9; i++)
            {
                var letter = (char)('A' + i - 1);
                if (gt.TryGetValue($"eh418_area{i}_name", out var n))
                    areaNames[letter] = n;
            }
            if (gt.TryGetValue("eh418_areaP_name", out var pn))
                areaNames['P'] = pn;
        }

        var allR = allTiles.Keys.Select(k => int.Parse(k.Split('-')[0])).ToList();
        var allC = allTiles.Keys.Select(k => int.Parse(k.Split('-')[1])).ToList();

        var elementData = LoadElementData(elementPath);
        var produceConsumeData = LoadProduceConsumeData(produceConsumePath);

        return new JamData
        {
            AllTiles = allTiles,
            FirstArea = firstArea,
            CoordToArea = coordToArea,
            AreaBounds = areaBounds,
            AreaNames = areaNames,
            Clouds = cloudRecords,
            Elements = elementData,
            ProduceConsume = produceConsumeData,
            MinR = allR.Min(),
            MaxR = allR.Max(),
            MinC = allC.Min(),
            MaxC = allC.Max(),
        };
    }

    private static ElementData LoadElementData(string path)
    {
        if (!File.Exists(path))
            return new ElementData();
        using var f = File.OpenRead(path);
        var root = JsonSerializer.Deserialize<JsonElement>(f);
        return new ElementData
        {
            MergeElements = ReadStringList(root, "merge_elements"),
            Producers = ReadStringList(root, "producers"),
            Ingredients = ReadStringList(root, "ingredients"),
            Obstacles = ReadStringList(root, "obstacles"),
        };
    }

    private static ProduceConsumeData LoadProduceConsumeData(string path)
    {
        if (!File.Exists(path))
            return new ProduceConsumeData();
        using var f = File.OpenRead(path);
        var root = JsonSerializer.Deserialize<JsonElement>(f);

        var chains = new Dictionary<string, List<string>>();
        if (root.TryGetProperty("producer_chains", out var pc))
            foreach (var kv in pc.EnumerateObject())
                chains[kv.Name] = kv
                    .Value.EnumerateArray()
                    .Select(x => x.GetString() ?? "")
                    .ToList();

        var recipes = new List<RecipeRecord>();
        if (root.TryGetProperty("jam_recipes", out var jr))
            foreach (var r in jr.EnumerateArray())
            {
                var needs = r.TryGetProperty("needs", out var nArr)
                    ? string.Join(
                        ", ",
                        nArr.EnumerateArray()
                            .Select(n =>
                                $"{n.GetProperty("element").GetString()}×{n.GetProperty("qty").GetInt32()}"
                            )
                    )
                    : "";
                var produce = r.TryGetProperty("produce", out var pArr)
                    ? string.Join(
                        ", ",
                        pArr.EnumerateArray()
                            .Select(p =>
                                $"{p.GetProperty("element").GetString()}×{p.GetProperty("qty").GetInt32()}"
                            )
                    )
                    : "";
                recipes.Add(
                    new RecipeRecord
                    {
                        Id = r.TryGetProperty("id", out var idv) ? idv.GetInt32() : 0,
                        Needs = needs,
                        Produce = produce,
                    }
                );
            }

        var tasks = new List<StoryTask>();
        if (root.TryGetProperty("story_tasks", out var st))
            foreach (var t in st.EnumerateArray())
                tasks.Add(
                    new StoryTask
                    {
                        Id = t.TryGetProperty("id", out var tv) ? tv.GetString() ?? "" : "",
                        Entity = t.TryGetProperty("entity", out var ev) ? ev.GetString() ?? "" : "",
                        Target = t.TryGetProperty("target", out var tgv)
                            ? tgv.GetString() ?? ""
                            : "",
                    }
                );

        var unlockEvents = root.TryGetProperty("unlock_events", out var ue)
            ? ue.EnumerateArray().Select(x => x.GetString() ?? "").ToList()
            : [];

        var areaRewards = new Dictionary<string, string>();
        if (root.TryGetProperty("area_rewards", out var ar))
            foreach (var kv in ar.EnumerateObject())
                areaRewards[kv.Name] = kv.Value.GetString() ?? "";

        return new ProduceConsumeData
        {
            ProducerChains = chains,
            Recipes = recipes,
            StoryTasks = tasks,
            UnlockEvents = unlockEvents,
            AreaRewards = areaRewards,
        };
    }

    private static List<string> ReadStringList(JsonElement root, string key) =>
        root.TryGetProperty(key, out var arr)
            ? arr.EnumerateArray().Select(x => x.GetString() ?? "").ToList()
            : [];

    // ── 总览：只画各 Area 轮廓 + 动线 ────────────────────────────────────────

    private static void BuildOverviewSheet(ExcelPackage pkg, JamData data)
    {
        int nRows = data.MaxR - data.MinR + 1;
        int nCols = data.MaxC - data.MinC + 1;
        Console.WriteLine($"  [总览] {nRows}×{nCols}（动线图）");
        var ws = pkg.Workbook.Worksheets.Add("总览");

        // 标题行
        ws.Cells[1, 1].Value = "行\\列";
        for (int ci = 0; ci < nCols; ci++)
            ws.Cells[1, ci + 2].Value = data.MinC + ci;

        // 每格：若属于某 Area 填字母+区域色，否则空白
        for (int ri = 0; ri < nRows; ri++)
        {
            int gameR = data.MinR + ri;
            int sheetRow = ri + 2;
            ws.Cells[sheetRow, 1].Value = gameR;

            for (int ci = 0; ci < nCols; ci++)
            {
                var coord = $"{gameR}-{data.MinC + ci}";
                if (!data.CoordToArea.TryGetValue(coord, out var letter))
                    continue;

                var cell = ws.Cells[sheetRow, ci + 2];
                cell.Value = letter.ToString();
                cell.Style.Font.Size = FontSize;
                cell.Style.Font.Bold = true;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                var aColor = AreaColors.GetValueOrDefault(letter, "EEEEEE");
                SetBg(cell, aColor);
            }
        }

        // 在各 Area 中心格写 Area 名称（覆盖字母，方便识别）
        foreach (var (letter, (minR, maxR, minC, maxC)) in data.AreaBounds)
        {
            int centerR = (minR + maxR) / 2;
            int centerC = (minC + maxC) / 2;
            int sheetRow = centerR - data.MinR + 2;
            int sheetCol = centerC - data.MinC + 2;
            var name = data.AreaNames.GetValueOrDefault(letter, $"Area{letter}");
            var cell = ws.Cells[sheetRow, sheetCol];
            cell.Value = $"{letter}:{name}";
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 7f;
        }

        // 列宽 / 行高（总览缩小一半，整体看动线）
        ws.Column(1).Width = RowNumColWidth;
        for (int ci = 2; ci <= nCols + 1; ci++)
            ws.Column(ci).Width = 2.0;
        for (int r = 1; r <= nRows + 1; r++)
            ws.Row(r).Height = 8f;
    }

    // ── 构建一个 sheet ────────────────────────────────────────────────────────

    private static void BuildMapSheet(
        ExcelPackage pkg,
        string sheetName,
        JamData data,
        int minR,
        int maxR,
        int minC,
        int maxC,
        char? areaLetter
    )
    {
        int nRows = maxR - minR + 1;
        int nCols = maxC - minC + 1;
        // sheet 布局：col1=行号, col2..n+1=地图, col n+3=图例色块, col n+4=图例说明
        int lgCol = nCols + 3; // 1-based 图例色块列

        Console.WriteLine($"  [{sheetName}] {nRows}×{nCols}");
        var ws = pkg.Workbook.Worksheets.Add(sheetName);

        // ── 标题行
        ws.Cells[1, 1].Value = "行\\列";
        for (int ci = 0; ci < nCols; ci++)
            ws.Cells[1, ci + 2].Value = minC + ci;
        ws.Cells[1, nCols + 2].Value = "";
        ws.Cells[1, lgCol].Value = "图例";
        ws.Cells[1, lgCol + 1].Value = "说明";

        // ── 地图数据行
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
                    if (color != null)
                        SetBg(cell, color);

                    if (data.FirstArea.Contains(coord))
                        cell.Style.Font.Bold = true;
                }
                else if (
                    areaLetter.HasValue
                    && data.CoordToArea.GetValueOrDefault(coord) == areaLetter.Value
                )
                {
                    // 空格子但属于本 Area → 叠加区域底色
                    var ac = AreaColors.GetValueOrDefault(areaLetter.Value, "EEEEEE");
                    SetBg(cell, ac);
                }
            }

            // 图例行（只填前 LegendItems.Length 行）
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

        // ── 列宽
        ws.Column(1).Width = RowNumColWidth;
        for (int ci = 2; ci <= nCols + 1; ci++)
            ws.Column(ci).Width = CellColWidth;

        // ── 行高
        for (int r = 1; r <= nRows + 1; r++)
            ws.Row(r).Height = RowHeight;
    }

    // ── 辅助：物品颜色 ────────────────────────────────────────────────────────

    private static string? ItemColor(string item)
    {
        foreach (var (prefix, color) in ItemColors)
            if (item.StartsWith(prefix))
                return color;
        return "F5F5F5";
    }

    // ── 辅助：物品简称 ────────────────────────────────────────────────────────

    private static string ItemLabel(string item)
    {
        if (string.IsNullOrEmpty(item) || item == "null")
            return "·";
        if (item.StartsWith("limitEventEnergy"))
            return "E";
        if (item.StartsWith("eventletter"))
            return "L";
        var m = Regex.Match(item, @"^(?:inr|ob|ehgo|tbeh|eh)?eh418_([a-z])([0-9]+)(?:_\d+)?$");
        if (m.Success)
            return m.Groups[1].Value + m.Groups[2].Value;
        m = Regex.Match(item, @"^(?:inr|ob|ehgo|tbeh|eh)?eh418_([a-z0-9]{1,3})");
        if (m.Success)
            return m.Groups[1].Value;
        if (item.StartsWith("obfc"))
            return "▪";
        if (item.StartsWith("gold"))
            return "G";
        return item[..Math.Min(3, item.Length)];
    }

    // ── 辅助：设背景色 ────────────────────────────────────────────────────────

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

    // ── 云雾区域 sheet ────────────────────────────────────────────────────────

    private static void BuildCloudsSheet(ExcelPackage pkg, JamData data)
    {
        Console.WriteLine($"  [云雾区域] {data.Clouds.Count} 条");
        var ws = pkg.Workbook.Worksheets.Add("云雾区域");

        string[] headers = ["云名", "所属Area", "格子数", "解锁条件(need)", "奖励(reward)"];
        for (int i = 0; i < headers.Length; i++)
        {
            var cell = ws.Cells[1, i + 1];
            cell.Value = headers[i];
            cell.Style.Font.Bold = true;
            SetBg(cell, "D9D9D9");
        }

        var sorted = data.Clouds.OrderBy(c => c.AreaLetter).ThenBy(c => c.Name).ToList();
        for (int ri = 0; ri < sorted.Count; ri++)
        {
            var c = sorted[ri];
            int row = ri + 2;
            ws.Cells[row, 1].Value = c.Name;
            ws.Cells[row, 2].Value = c.AreaLetter.ToString();
            ws.Cells[row, 3].Value = c.TileCount;
            ws.Cells[row, 4].Value = c.Need;
            ws.Cells[row, 5].Value = c.Reward;

            var aColor = AreaColors.GetValueOrDefault(c.AreaLetter, "EEEEEE");
            SetBg(ws.Cells[row, 2], aColor);
        }

        ws.Column(1).Width = 18;
        ws.Column(2).Width = 8;
        ws.Column(3).Width = 8;
        ws.Column(4).Width = 28;
        ws.Column(5).Width = 28;
        ws.Cells[ws.Dimension.Address].Style.Font.Size = FontSize;
    }

    // ── 物品统计 sheet ────────────────────────────────────────────────────────

    private static void BuildItemStatsSheet(ExcelPackage pkg, JamData data)
    {
        var el = data.Elements;
        Console.WriteLine(
            $"  [物品统计] merge={el.MergeElements.Count} producers={el.Producers.Count} "
                + $"ingredients={el.Ingredients.Count} obstacles={el.Obstacles.Count}"
        );
        var ws = pkg.Workbook.Worksheets.Add("物品统计");

        int row = 1;
        row = WriteSection(ws, row, "合并元素(merge_elements)", "FFFF99", el.MergeElements);
        row = WriteSection(ws, row, "生产者(producers)", "CCFFCC", el.Producers);
        row = WriteSection(ws, row, "原料(ingredients)", "FFD9AA", el.Ingredients);
        row = WriteSection(ws, row, "障碍(obstacles)", "DDDDDD", el.Obstacles);

        ws.Column(1).Width = 35;
        ws.Cells[ws.Dimension?.Address ?? "A1"].Style.Font.Size = FontSize;
    }

    private static int WriteSection(
        ExcelWorksheet ws,
        int startRow,
        string title,
        string color,
        List<string> items
    )
    {
        var hdr = ws.Cells[startRow, 1];
        hdr.Value = title;
        hdr.Style.Font.Bold = true;
        SetBg(hdr, color);
        startRow++;

        foreach (var item in items)
        {
            ws.Cells[startRow, 1].Value = item;
            startRow++;
        }
        return startRow + 1; // blank separator row
    }

    // ── Jam元素关系 sheet ─────────────────────────────────────────────────────

    private static void BuildElementRelationsSheet(ExcelPackage pkg, JamData data)
    {
        var pc = data.ProduceConsume;
        Console.WriteLine(
            $"  [Jam元素关系] chains={pc.ProducerChains.Count} recipes={pc.Recipes.Count} "
                + $"tasks={pc.StoryTasks.Count} unlocks={pc.UnlockEvents.Count}"
        );
        var ws = pkg.Workbook.Worksheets.Add("Jam元素关系");

        int row = 1;

        // — 生产链
        WriteRowHeader(ws, row++, "生产链(producer_chains)", "CCFFCC");
        WriteRowBold(ws, row++, ["基础元素", "升级链(Lv2→5)"]);
        foreach (var (baseEl, chain) in pc.ProducerChains.OrderBy(kv => kv.Key))
        {
            ws.Cells[row, 1].Value = baseEl;
            ws.Cells[row, 2].Value = string.Join(" → ", chain);
            row++;
        }
        row++;

        // — 配方
        WriteRowHeader(ws, row++, "果酱配方(jam_recipes)", "FFFF99");
        WriteRowBold(ws, row++, ["配方ID", "原料(元素×数量)", "产出(元素×数量)"]);
        foreach (var r in pc.Recipes)
        {
            ws.Cells[row, 1].Value = r.Id;
            ws.Cells[row, 2].Value = r.Needs;
            ws.Cells[row, 3].Value = r.Produce;
            row++;
        }
        row++;

        // — 故事任务
        WriteRowHeader(ws, row++, "故事任务(story_tasks)", "99CCFF");
        WriteRowBold(ws, row++, ["任务ID", "操作类型", "目标元素"]);
        foreach (var t in pc.StoryTasks)
        {
            ws.Cells[row, 1].Value = t.Id;
            ws.Cells[row, 2].Value = t.Entity;
            ws.Cells[row, 3].Value = t.Target;
            row++;
        }
        row++;

        // — 解锁事件
        WriteRowHeader(ws, row++, "解锁事件(unlock_events)", "FFD9AA");
        foreach (var ev in pc.UnlockEvents)
        {
            ws.Cells[row, 1].Value = ev;
            row++;
        }
        row++;

        // — 区域奖励
        WriteRowHeader(ws, row++, "区域奖励(area_rewards)", "DDDDDD");
        WriteRowBold(ws, row++, ["元素ID", "奖励"]);
        foreach (var (el, reward) in pc.AreaRewards)
        {
            ws.Cells[row, 1].Value = el;
            ws.Cells[row, 2].Value = reward;
            row++;
        }

        ws.Column(1).Width = 32;
        ws.Column(2).Width = 45;
        ws.Column(3).Width = 32;
        ws.Cells[ws.Dimension?.Address ?? "A1"].Style.Font.Size = FontSize;
    }

    private static void WriteRowHeader(ExcelWorksheet ws, int row, string title, string color)
    {
        var cell = ws.Cells[row, 1];
        cell.Value = title;
        cell.Style.Font.Bold = true;
        SetBg(cell, color);
    }

    private static void WriteRowBold(ExcelWorksheet ws, int row, string[] cols)
    {
        for (int i = 0; i < cols.Length; i++)
        {
            ws.Cells[row, i + 1].Value = cols[i];
            ws.Cells[row, i + 1].Style.Font.Bold = true;
        }
    }
}
