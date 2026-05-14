using System.Drawing;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NumDesTools.Scanner;

/// <summary>
/// 海岛活动（OceanIsland）地图可视化写入器。
/// 数据源：C:/tmp/mftown/mapCloudOceanIsland.json
/// 输出：一个 xlsx，含「总览」sheet。
/// 格式与 JamMapWriter 一致：列宽 28px≈4字符，行高 15pt，字体 8pt，居中，物品染色，图例。
/// </summary>
public static class OceanIslandMapWriter
{
    private const string DefaultDataPath =
        @"C:\tmp\mftown\mapCloudOceanIsland.json";
    private static readonly string DefaultOutputDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "workspace");

    private const double CellColWidth  = 4.0;
    private const double RowNumColWidth = 5.7;
    private const float  RowHeight     = 15f;
    private const float  FontSize      = 8f;

    // ── 物品前缀颜色（OceanIsland 专属）
    private static readonly (string Prefix, string Color)[] ItemColors =
    [
        ("tbocopt", "FFD9AA"),   // 可选任务物品
        ("tboc",    "FFD9AA"),   // 任务物品
        ("sdoc",    "FF9999"),   // 特殊障碍
        ("ocean",   "99CCFF"),   // 主资源
        ("gold",    "FFE566"),   // 金币
        ("null",    "FFFFFF"),   // 空
    ];

    private static readonly (string Label, string Color)[] LegendItems =
    [
        ("主资源(ocean)", "99CCFF"),
        ("任务物品(tboc)", "FFD9AA"),
        ("特殊障碍(sdoc)", "FF9999"),
        ("金币(gold)",     "FFE566"),
        ("空格(null)",     "FFFFFF"),
    ];

    // ── 数据模型 ─────────────────────────────────────────────────────────────

    private sealed class CloudInfo
    {
        public string Name { get; set; } = "";
        public int Level { get; set; }
        public HashSet<string> Coords { get; set; } = [];
        public int MinR, MaxR, MinC, MaxC;
    }

    private sealed class OceanData
    {
        public Dictionary<string, string> AllTiles { get; set; } = [];
        public List<CloudInfo> Clouds { get; set; } = [];
        public int MinR, MaxR, MinC, MaxC;
    }

    // ── 公共入口 ─────────────────────────────────────────────────────────────

    public static void Run(string? outputDir = null, string dataPath = DefaultDataPath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var dir = outputDir ?? DefaultOutputDir;
        var outputPath = Path.Combine(dir, "CC收集活动-海岛地编信息.xlsx");

        var data = LoadData(dataPath);
        Console.WriteLine(
            $"[OceanIsland] 全图 {data.MaxR - data.MinR + 1}行×{data.MaxC - data.MinC + 1}列  "
            + $"({data.MinR}~{data.MaxR}, {data.MinC}~{data.MaxC})  Clouds={data.Clouds.Count}");

        using var pkg = new ExcelPackage();
        BuildMapSheet(pkg, "总览", data, data.MinR, data.MaxR, data.MinC, data.MaxC, null);

        foreach (var cloud in data.Clouds.OrderBy(c => c.Level))
            BuildMapSheet(pkg, $"Lv{cloud.Level}-{cloud.Name}", data,
                cloud.MinR, cloud.MaxR, cloud.MinC, cloud.MaxC, cloud);

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        pkg.SaveAs(new FileInfo(outputPath));
        Console.WriteLine($"[OceanIsland] 已写入：{outputPath}");
    }

    // ── 加载数据 ─────────────────────────────────────────────────────────────

    private static OceanData LoadData(string dataPath)
    {
        using var f = File.OpenRead(dataPath);
        var rawClouds = JsonSerializer.Deserialize<JsonElement[]>(f)
            ?? throw new InvalidDataException("无法解析 OceanIsland JSON");

        var allTiles = new Dictionary<string, string>();
        var clouds = new List<CloudInfo>();

        foreach (var raw in rawClouds)
        {
            var name  = raw.TryGetProperty("name",  out var nv) ? nv.GetString() ?? "" : "";
            var level = raw.TryGetProperty("level", out var lv) ? lv.GetInt32() : 0;
            var coords = new HashSet<string>();

            foreach (var tile in raw.GetProperty("area").EnumerateArray())
            {
                var parts = tile.GetString()?.Split('#');
                if (parts?.Length == 2)
                {
                    allTiles[parts[0]] = parts[1];
                    coords.Add(parts[0]);
                }
            }

            if (coords.Count == 0) continue;

            var rs = coords.Select(c => int.Parse(c.Split('-')[0])).ToList();
            var cs = coords.Select(c => int.Parse(c.Split('-')[1])).ToList();
            clouds.Add(new CloudInfo
            {
                Name = name, Level = level, Coords = coords,
                MinR = rs.Min(), MaxR = rs.Max(),
                MinC = cs.Min(), MaxC = cs.Max(),
            });
        }

        var allR = allTiles.Keys.Select(k => int.Parse(k.Split('-')[0])).ToList();
        var allC = allTiles.Keys.Select(k => int.Parse(k.Split('-')[1])).ToList();

        return new OceanData
        {
            AllTiles = allTiles,
            Clouds = clouds,
            MinR = allR.Min(), MaxR = allR.Max(),
            MinC = allC.Min(), MaxC = allC.Max(),
        };
    }

    // ── 构建 sheet ────────────────────────────────────────────────────────────

    private static void BuildMapSheet(
        ExcelPackage pkg,
        string sheetName,
        OceanData data,
        int minR, int maxR, int minC, int maxC,
        CloudInfo? cloud)
    {
        int nRows = maxR - minR + 1;
        int nCols = maxC - minC + 1;
        int lgCol  = nCols + 3;

        Console.WriteLine($"  [{sheetName}] {nRows}×{nCols}");
        var ws = pkg.Workbook.Worksheets.Add(sheetName);

        // 标题行
        ws.Cells[1, 1].Value = "行\\列";
        for (int ci = 0; ci < nCols; ci++)
            ws.Cells[1, ci + 2].Value = minC + ci;
        ws.Cells[1, lgCol].Value     = "图例";
        ws.Cells[1, lgCol + 1].Value = "说明";

        // 地图行
        for (int ri = 0; ri < nRows; ri++)
        {
            int gameR    = minR + ri;
            int sheetRow = ri + 2;

            ws.Cells[sheetRow, 1].Value = gameR;

            for (int ci = 0; ci < nCols; ci++)
            {
                int gameC = minC + ci;
                var coord = $"{gameR}-{gameC}";
                var item  = data.AllTiles.GetValueOrDefault(coord, "");
                if (string.IsNullOrEmpty(item))
                {
                    if (cloud?.Coords.Contains(coord) == true)
                        SetBg(ws.Cells[sheetRow, ci + 2], "D6EAF8");
                    continue;
                }

                var cell = ws.Cells[sheetRow, ci + 2];
                cell.Value = ItemLabel(item);
                cell.Style.Font.Size = FontSize;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.VerticalAlignment   = ExcelVerticalAlignment.Center;

                var color = ItemColor(item);
                if (color != null) SetBg(cell, color);
                if (cloud?.Coords.Contains(coord) == true)
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin,
                        HexColor("4A90D9"));
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

        // 列宽 / 行高
        ws.Column(1).Width = RowNumColWidth;
        for (int ci = 2; ci <= nCols + 1; ci++)
            ws.Column(ci).Width = CellColWidth;
        for (int r = 1; r <= nRows + 1; r++)
            ws.Row(r).Height = RowHeight;
    }

    // ── 辅助 ─────────────────────────────────────────────────────────────────

    private static string? ItemColor(string item)
    {
        foreach (var (prefix, color) in ItemColors)
            if (item.StartsWith(prefix)) return color;
        return "F5F5F5";
    }

    private static string ItemLabel(string item)
    {
        if (string.IsNullOrEmpty(item) || item == "null") return "·";
        if (item.StartsWith("sdoc")) return "S";
        if (item.StartsWith("gold")) return "G";
        // ocean_N_M → N
        var m = Regex.Match(item, @"^ocean_(\d+)_(\d+)$");
        if (m.Success) return m.Groups[1].Value;
        // tboc/tbocopt → 取前缀后的简短标识
        m = Regex.Match(item, @"^tboc(?:opt)?_([a-z]+)_(\d+)$");
        if (m.Success) return m.Groups[1].Value[..Math.Min(2, m.Groups[1].Value.Length)] + m.Groups[2].Value;
        return item[..Math.Min(3, item.Length)];
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
