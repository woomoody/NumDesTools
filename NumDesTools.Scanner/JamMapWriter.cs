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
    private const string DefaultDataPath =
        @"C:\tmp\mftown\analysis\jam_config_full.json";
    private const string DefaultGameTextPath =
        @"C:\tmp\mftown\game_text_en.json";

    // ── 尺寸常量（openpyxl 到 EPPlus 换算：px/7 ≈ 字符宽，1pt=1磅）
    private const double CellColWidth = 4.0;    // 28px ≈ 4 字符
    private const double RowNumColWidth = 5.7;  // 40px ≈ 5.7 字符
    private const float RowHeight = 15f;         // 20px ≈ 15pt
    private const float FontSize = 8f;

    // ── Area 背景色
    private static readonly Dictionary<char, string> AreaColors = new()
    {
        { 'A', "FFD9D9" }, { 'B', "FFEDD9" }, { 'C', "FFFFD9" },
        { 'D', "D9F2D9" }, { 'E', "D9EDFF" }, { 'F', "EAD9FF" },
        { 'G', "FFD4EF" }, { 'H', "C8F5F5" }, { 'I', "F5E3C8" }, { 'P', "E0E0E0" },
    };

    // ── 物品前缀颜色
    private static readonly (string Prefix, string Color)[] ItemColors =
    [
        ("limitEventEnergy", "FF9999"),
        ("eventletter",      "99CCFF"),
        ("inreh418",         "CCFFCC"),
        ("ehgo418",          "FFFF99"),
        ("eh418",            "FFFF99"),
        ("obeh418",          "DDDDDD"),
        ("obfc",             "DDDDDD"),
        ("tbeh418",          "FFD9AA"),
        ("gold",             "FFE566"),
        ("null",             "FFFFFF"),
    ];

    // ── 图例
    private static readonly (string Label, string Color)[] LegendItems =
    [
        ("初始解锁(inr)", "CCFFCC"),
        ("可收获(eh/go)", "FFFF99"),
        ("任务物品(tb)",  "FFD9AA"),
        ("能量限制块",    "FF9999"),
        ("信件(letter)",  "99CCFF"),
        ("障碍(ob/obfc)", "DDDDDD"),
        ("金币(gold)",    "FFE566"),
        ("空格(null)",    "FFFFFF"),
    ];

    // ── 数据模型 ─────────────────────────────────────────────────────────────

    private sealed class JamData
    {
        public Dictionary<string, string> AllTiles { get; set; } = [];
        public HashSet<string> FirstArea { get; set; } = [];
        public Dictionary<string, char> CoordToArea { get; set; } = [];
        public Dictionary<char, (int MinR, int MaxR, int MinC, int MaxC)> AreaBounds { get; set; } = [];
        public Dictionary<char, string> AreaNames { get; set; } = [];
        public int MinR, MaxR, MinC, MaxC;
    }

    // ── 公共入口 ─────────────────────────────────────────────────────────────

    public static void Run(
        string outputPath,
        string dataPath = DefaultDataPath,
        string gameTextPath = DefaultGameTextPath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var data = LoadData(dataPath, gameTextPath);
        Console.WriteLine(
            $"[JamMap] 全图 {data.MaxR - data.MinR + 1}行×{data.MaxC - data.MinC + 1}列  "
            + $"({data.MinR}~{data.MaxR}, {data.MinC}~{data.MaxC})");
        Console.WriteLine($"[JamMap] Areas: {string.Join(' ', data.AreaBounds.Keys.Order())}");

        using var pkg = new ExcelPackage();

        BuildMapSheet(pkg, "总览", data, data.MinR, data.MaxR, data.MinC, data.MaxC, null);

        foreach (var letter in data.AreaBounds.Keys.Order())
        {
            var (minR, maxR, minC, maxC) = data.AreaBounds[letter];
            var name = data.AreaNames.GetValueOrDefault(letter, $"Area {letter}");
            BuildMapSheet(pkg, $"Area{letter}-{name}", data, minR, maxR, minC, maxC, letter);
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        pkg.SaveAs(new FileInfo(outputPath));
        Console.WriteLine($"[JamMap] 已写入：{outputPath}");
    }

    // ── 加载数据 ─────────────────────────────────────────────────────────────

    private static JamData LoadData(string dataPath, string gameTextPath)
    {
        using var df = File.OpenRead(dataPath);
        var root = JsonSerializer.Deserialize<JsonElement>(df);

        var allTiles = new Dictionary<string, string>();
        foreach (var kv in root.GetProperty("allTiles").EnumerateObject())
            allTiles[kv.Name] = kv.Value.GetString() ?? "";

        var firstArea = new HashSet<string>();
        foreach (var kv in root.GetProperty("firstArea").EnumerateObject())
            firstArea.Add(kv.Name);

        // coord → area letter
        var coordToArea = new Dictionary<string, char>();
        foreach (var cloud in root.GetProperty("clouds").EnumerateArray())
        {
            var letters = new List<char>();
            foreach (var tile in cloud.GetProperty("area").EnumerateArray())
            {
                var parts = tile.GetString()?.Split('#');
                if (parts?.Length == 2)
                {
                    var m = Regex.Match(parts[1],
                        @"^(?:obeh418_|eh418_|ehgo418_|inreh418_|tbeh418_)([a-z])");
                    if (m.Success) letters.Add(char.ToUpper(m.Groups[1].Value[0]));
                }
            }
            if (letters.Count == 0) continue;

            var dominant = letters.GroupBy(x => x).MaxBy(g => g.Count())!.Key;
            foreach (var tile in cloud.GetProperty("area").EnumerateArray())
            {
                var parts = tile.GetString()?.Split('#');
                if (parts?.Length == 2)
                    coordToArea[parts[0]] = dominant;
            }
        }

        // area bounds
        var areaBounds = new Dictionary<char, (int, int, int, int)>();
        foreach (var (coord, letter) in coordToArea)
        {
            var seg = coord.Split('-');
            int r = int.Parse(seg[0]), c = int.Parse(seg[1]);
            if (!areaBounds.TryGetValue(letter, out var b))
                areaBounds[letter] = (r, r, c, c);
            else
                areaBounds[letter] = (Math.Min(b.Item1, r), Math.Max(b.Item2, r),
                                      Math.Min(b.Item3, c), Math.Max(b.Item4, c));
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
                if (gt.TryGetValue($"eh418_area{i}_name", out var n)) areaNames[letter] = n;
            }
            if (gt.TryGetValue("eh418_areaP_name", out var pn)) areaNames['P'] = pn;
        }

        var allR = allTiles.Keys.Select(k => int.Parse(k.Split('-')[0])).ToList();
        var allC = allTiles.Keys.Select(k => int.Parse(k.Split('-')[1])).ToList();

        return new JamData
        {
            AllTiles = allTiles,
            FirstArea = firstArea,
            CoordToArea = coordToArea,
            AreaBounds = areaBounds,
            AreaNames = areaNames,
            MinR = allR.Min(), MaxR = allR.Max(),
            MinC = allC.Min(), MaxC = allC.Max(),
        };
    }

    // ── 构建一个 sheet ────────────────────────────────────────────────────────

    private static void BuildMapSheet(
        ExcelPackage pkg,
        string sheetName,
        JamData data,
        int minR, int maxR, int minC, int maxC,
        char? areaLetter)
    {
        int nRows = maxR - minR + 1;
        int nCols = maxC - minC + 1;
        // sheet 布局：col1=行号, col2..n+1=地图, col n+3=图例色块, col n+4=图例说明
        int lgCol = nCols + 3;  // 1-based 图例色块列

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
                else if (areaLetter.HasValue
                    && data.CoordToArea.GetValueOrDefault(coord) == areaLetter.Value)
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
            if (item.StartsWith(prefix)) return color;
        return "F5F5F5";
    }

    // ── 辅助：物品简称 ────────────────────────────────────────────────────────

    private static string ItemLabel(string item)
    {
        if (string.IsNullOrEmpty(item) || item == "null") return "·";
        if (item.StartsWith("limitEventEnergy")) return "E";
        if (item.StartsWith("eventletter")) return "L";
        var m = Regex.Match(item, @"^(?:inr|ob|ehgo|tbeh|eh)?eh418_([a-z])([0-9]+)(?:_\d+)?$");
        if (m.Success) return m.Groups[1].Value + m.Groups[2].Value;
        m = Regex.Match(item, @"^(?:inr|ob|ehgo|tbeh|eh)?eh418_([a-z0-9]{1,3})");
        if (m.Success) return m.Groups[1].Value;
        if (item.StartsWith("obfc")) return "▪";
        if (item.StartsWith("gold")) return "G";
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
}
