using System.Text;
using OfficeOpenXml;

namespace NumDesTools.Tests;

// 回归测试：ApplyDelete 曾经用多次 sheet.DeleteRow 从下往上删,在 EPPlus 8.2.0~8.6.1（含最新版，实测过）
// 大量、跨度很大的区间上会触发 EPPlus 自身的行索引偏移 bug，删除后中间留出一大段完全空白的行
// （详见 2026-07-08-epplus-deleterow-bug-fix-task.md）。现在改成整行拷贝重建，这里锁定这个行为。
public class ActivityDataBackupToolApplyDeleteTests
{
    static ActivityDataBackupToolApplyDeleteTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private const int SeedRows = 176978;
    private const int Cols = 12;

    // 真实场景复现用的坐标：236 个互不重叠、大小跨度很大（几十行到上千行）、间隔也跨度很大
    // （几行到近万行）的区间，来自一次真实的 Icon.xlsx 删除操作。
    private static readonly (int Start, int End)[] LargeScatteredBlocks =
    {
        (4812, 4869),
        (4870, 4927),
        (4928, 4985),
        (4986, 5043),
        (5044, 5101),
        (5102, 5159),
        (5160, 5217),
        (5218, 5275),
        (5276, 5333),
        (5334, 5391),
        (5392, 5450),
        (5451, 5509),
        (5528, 5586),
        (5605, 5663),
        (5682, 5740),
        (5757, 5815),
        (6727, 6804),
        (6805, 6880),
        (6881, 6956),
        (6957, 7032),
        (7033, 7092),
        (7093, 7168),
        (7169, 7228),
        (7229, 7288),
        (7289, 7348),
        (7349, 7408),
        (7409, 7468),
        (7469, 7528),
        (7529, 7588),
        (7589, 7648),
        (7649, 7708),
        (7709, 7768),
        (7769, 7828),
        (7829, 7888),
        (7889, 7948),
        (7949, 8008),
        (8009, 8068),
        (8069, 8128),
        (8129, 8188),
        (8189, 8248),
        (8249, 8308),
        (9818, 9901),
        (9902, 9969),
        (9970, 10037),
        (10038, 10121),
        (10122, 10205),
        (10206, 10289),
        (10290, 10373),
        (10374, 10457),
        (10458, 10541),
        (10574, 10657),
        (10690, 10773),
        (10806, 10889),
        (10924, 11007),
        (11040, 11123),
        (11840, 11911),
        (11912, 11980),
        (11981, 12049),
        (12058, 12135),
        (12164, 12244),
        (12273, 12350),
        (12395, 12479),
        (12480, 12564),
        (12593, 12677),
        (12706, 12790),
        (12819, 12903),
        (12932, 13016),
        (13045, 13129),
        (13158, 13242),
        (13271, 13355),
        (13356, 13440),
        (13491, 13625),
        (13654, 13788),
        (13817, 13951),
        (15670, 15808),
        (15838, 15972),
        (16002, 16136),
        (16166, 16300),
        (16330, 16464),
        (16494, 16628),
        (16658, 16792),
        (16821, 16955),
        (16984, 17118),
        (17147, 17281),
        (17310, 17444),
        (17473, 17613),
        (17642, 17776),
        (17805, 17939),
        (18915, 19053),
        (19115, 19320),
        (19321, 19526),
        (19527, 19665),
        (19727, 19926),
        (19927, 20126),
        (20127, 20326),
        (20327, 20526),
        (20527, 20726),
        (20727, 20926),
        (20927, 21126),
        (21127, 21326),
        (21327, 21526),
        (21527, 21726),
        (21727, 21926),
        (21927, 22126),
        (22127, 22326),
        (22327, 22457),
        (24574, 24844),
        (25156, 25426),
        (25427, 25697),
        (25698, 25968),
        (25969, 26239),
        (26240, 26510),
        (26511, 26781),
        (26782, 27052),
        (27053, 27323),
        (27324, 27594),
        (27595, 27865),
        (27866, 28136),
        (28137, 28407),
        (28408, 28678),
        (28679, 28949),
        (28950, 29220),
        (30525, 30911),
        (30912, 31298),
        (31299, 31685),
        (31686, 32074),
        (32075, 32461),
        (32462, 32848),
        (32849, 33235),
        (33236, 33622),
        (33623, 34009),
        (34010, 34396),
        (34397, 34783),
        (34784, 35170),
        (35171, 35557),
        (35558, 35944),
        (35945, 36331),
        (36332, 36718),
        (36719, 37105),
        (37106, 37492),
        (37493, 37879),
        (37880, 38259),
        (39331, 39926),
        (39927, 40518),
        (40519, 41110),
        (41111, 41702),
        (41703, 42294),
        (42295, 42886),
        (42887, 43472),
        (43473, 44058),
        (44059, 44644),
        (45533, 46010),
        (46011, 46487),
        (46488, 46964),
        (46965, 47441),
        (47442, 47918),
        (47919, 48395),
        (48396, 48872),
        (48873, 49349),
        (49350, 49826),
        (49827, 50303),
        (50304, 50780),
        (50781, 51257),
        (51258, 51734),
        (51735, 52210),
        (52211, 52687),
        (52688, 53164),
        (53165, 53641),
        (53642, 54125),
        (54380, 55201),
        (55202, 56023),
        (56024, 56916),
        (56917, 57738),
        (57739, 58714),
        (58715, 59536),
        (59537, 60358),
        (60359, 61180),
        (61181, 62002),
        (62003, 62824),
        (62825, 63646),
        (63647, 64468),
        (64469, 65290),
        (65291, 66112),
        (66113, 66934),
        (66935, 67756),
        (67757, 68577),
        (68578, 69398),
        (69399, 70219),
        (70220, 71041),
        (71042, 71863),
        (71864, 72684),
        (72685, 73505),
        (73506, 74324),
        (74325, 75143),
        (75144, 75962),
        (76783, 77601),
        (77602, 78430),
        (78623, 78996),
        (78997, 79370),
        (79371, 79744),
        (79745, 80118),
        (80119, 80492),
        (80493, 80866),
        (80867, 81240),
        (81241, 81614),
        (81615, 81988),
        (82160, 82695),
        (82696, 83233),
        (83234, 83769),
        (85664, 86369),
        (87076, 87781),
        (87782, 89193),
        (89194, 89899),
        (89900, 90605),
        (90606, 91311),
        (91312, 92017),
        (92018, 92723),
        (92724, 93429),
        (93430, 94135),
        (94136, 94841),
        (94842, 95547),
        (95548, 96249),
        (96250, 96951),
        (96952, 97653),
        (97654, 98355),
        (98356, 99057),
        (100907, 101792),
        (101793, 102329),
        (103295, 104180),
        (104182, 105067),
        (105069, 105954),
        (105956, 106841),
        (106843, 107728),
        (107731, 108616),
        (118372, 119335),
        (120322, 121285),
    };

    private static string BuildActivityTemplate(
        string dir,
        IReadOnlyList<(string ActivityId, string Start, string End)> ranges
    )
    {
        var templatePath = Path.Combine(dir, "template.xlsx");
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("大文件备份");
        sheet.Cells[1, 5].Value = "Icon.xlsx";
        sheet.Cells[2, 1].Value = "Ignore";
        sheet.Cells[2, 2].Value = "数据状态";
        sheet.Cells[2, 3].Value = "活动ID";
        sheet.Cells[2, 4].Value = "活动名称";
        sheet.Cells[2, 5].Value = "起始";
        sheet.Cells[2, 6].Value = "结束";

        var row = 3;
        foreach (var (id, start, end) in ranges)
        {
            sheet.Cells[row, 3].Value = id;
            sheet.Cells[row, 4].Value = $"活动{id}";
            sheet.Cells[row, 5].Value = start;
            sheet.Cells[row, 6].Value = end;
            row++;
        }

        pkg.SaveAs(new FileInfo(templatePath));
        return templatePath;
    }

    private static string BuildIconLikeSheet(
        string dir,
        int totalRows,
        IEnumerable<string> idsInOrder
    )
    {
        var path = Path.Combine(dir, "Icon.xlsx");
        using var pkg = new ExcelPackage();
        var sheet = pkg.Workbook.Worksheets.Add("Sheet1");
        sheet.Cells[2, 2].Value = "id";

        var row = XlsxCrossSync.DataStartRow;
        foreach (var id in idsInOrder)
        {
            for (var c = 1; c <= Cols; c++)
                sheet.Cells[row, c].Value = c == 2 ? id : $"r{row}c{c}";
            row++;
        }

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private static List<(int Start, int End)> FindInternalBlankRanges(string path)
    {
        using var pkg = new ExcelPackage(new FileInfo(path));
        var sheet = pkg.Workbook.Worksheets["Sheet1"]!;
        var idCol = PubMetToExcel.FindSourceCol(
            sheet,
            XlsxCrossSync.HeaderRow,
            XlsxCrossSync.KeyColumnName
        );
        var lastRow = sheet.Dimension?.End.Row ?? 0;
        var ranges = new List<(int Start, int End)>();
        var blankStart = -1;
        for (var r = XlsxCrossSync.DataStartRow; r <= lastRow; r++)
        {
            var hasId = !string.IsNullOrWhiteSpace(sheet.Cells[r, idCol].Text?.Trim());
            if (!hasId)
            {
                if (blankStart == -1)
                    blankStart = r;
            }
            else if (blankStart != -1)
            {
                ranges.Add((blankStart, r - 1));
                blankStart = -1;
            }
        }
        if (blankStart != -1)
            ranges.Add((blankStart, lastRow));
        return ranges;
    }

    [Fact]
    public void ApplyDelete_ManyLargeScatteredBlocks_DoesNotLeaveInternalBlankRows()
    {
        var workDir = Path.Combine(
            Path.GetTempPath(),
            "ApplyDeleteTests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(workDir);

        var ids = Enumerable.Range(1, SeedRows).Select(r => r.ToString()).ToList();
        var iconPath = BuildIconLikeSheet(workDir, SeedRows, ids);

        var deleteRowOffset = XlsxCrossSync.DataStartRow - 1;
        var ranges = LargeScatteredBlocks
            .Select(b =>
                (
                    Activity: new ActivityDataBackupTool.Activity(
                        0,
                        $"a{b.Start}",
                        "",
                        new Dictionary<string, (string Start, string End)>()
                    ),
                    Start: (b.Start + deleteRowOffset).ToString(),
                    End: (b.End + deleteRowOffset).ToString()
                )
            )
            .ToList();

        var expectedDeletedRows = LargeScatteredBlocks.Sum(b => b.End - b.Start + 1);

        var (summary, _) = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", iconPath, ranges);

        Assert.Contains($"删除 {expectedDeletedRows} 行", summary);

        var blanks = FindInternalBlankRanges(iconPath);
        Assert.Empty(blanks);

        using var pkg = new ExcelPackage(new FileInfo(iconPath));
        var sheet = pkg.Workbook.Worksheets["Sheet1"]!;
        var expectedLastRow = deleteRowOffset + SeedRows - expectedDeletedRows;
        Assert.Equal(expectedLastRow, sheet.Dimension!.End.Row);
    }

    [Fact]
    public void ApplyDelete_SmallNonOverlappingBlocks_StillWorks()
    {
        var workDir = Path.Combine(
            Path.GetTempPath(),
            "ApplyDeleteTests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(workDir);

        const int totalRows = 200;
        var ids = Enumerable.Range(1, totalRows).Select(r => r.ToString()).ToList();
        var iconPath = BuildIconLikeSheet(workDir, totalRows, ids);

        var deleteRowOffset = XlsxCrossSync.DataStartRow - 1;
        var blocks = new (int Start, int End)[] { (10, 15), (50, 60), (100, 120) };
        var ranges = blocks
            .Select(b =>
                (
                    Activity: new ActivityDataBackupTool.Activity(
                        0,
                        $"a{b.Start}",
                        "",
                        new Dictionary<string, (string Start, string End)>()
                    ),
                    Start: (b.Start + deleteRowOffset).ToString(),
                    End: (b.End + deleteRowOffset).ToString()
                )
            )
            .ToList();

        var (summary, _) = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", iconPath, ranges);
        var expectedDeletedRows = blocks.Sum(b => b.End - b.Start + 1);
        Assert.Contains($"删除 {expectedDeletedRows} 行", summary);
        Assert.Empty(FindInternalBlankRanges(iconPath));
    }

    [Fact]
    public void ApplyDelete_RowsWithComments_DoesNotThrow()
    {
        // 回归用例：sheet.Cells[...].Copy(...) 整行拷贝会连带处理单元格批注，实测在真实 Public 表上
        // 撞到 EPPlus 内部 "Text can't be null" 异常（ExcelCommentCollection.Add 里）。改成逐格搬 Value
        // 之后不再走 Copy 的批注拷贝路径，这里用真实带批注的单元格锁定这个行为。
        var workDir = Path.Combine(
            Path.GetTempPath(),
            "ApplyDeleteTests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(workDir);

        const int totalRows = 200;
        var ids = Enumerable.Range(1, totalRows).Select(r => r.ToString()).ToList();
        var iconPath = BuildIconLikeSheet(workDir, totalRows, ids);

        using (var pkg = new ExcelPackage(new FileInfo(iconPath)))
        {
            var sheet = pkg.Workbook.Worksheets["Sheet1"]!;
            // 批注要落在"要保留"的行上，才能验证 Copy 路径确实被绕开了（落在被删行上不会触发 Copy）。
            sheet.Cells[XlsxCrossSync.DataStartRow + 30, 3].AddComment("备注文本", "tester");
            sheet.Cells[XlsxCrossSync.DataStartRow + 150, 3].AddComment("另一条备注", "tester");
            pkg.Save();
        }

        var deleteRowOffset = XlsxCrossSync.DataStartRow - 1;
        var blocks = new (int Start, int End)[] { (10, 15), (50, 60), (100, 120) };
        var ranges = blocks
            .Select(b =>
                (
                    Activity: new ActivityDataBackupTool.Activity(
                        0,
                        $"a{b.Start}",
                        "",
                        new Dictionary<string, (string Start, string End)>()
                    ),
                    Start: (b.Start + deleteRowOffset).ToString(),
                    End: (b.End + deleteRowOffset).ToString()
                )
            )
            .ToList();

        var (summary, _) = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", iconPath, ranges);
        var expectedDeletedRows = blocks.Sum(b => b.End - b.Start + 1);
        Assert.Contains($"删除 {expectedDeletedRows} 行", summary);
        Assert.Empty(FindInternalBlankRanges(iconPath));
    }

    [Fact]
    public void BuildActivityTemplate_HelperProducesReadableTemplate()
    {
        // 仅验证测试自带的模板构造 helper 本身可用，避免以后改测试基础设施时静默出错。
        var workDir = Path.Combine(
            Path.GetTempPath(),
            "ApplyDeleteTests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(workDir);
        var templatePath = BuildActivityTemplate(workDir, new[] { ("2001", "10", "20") });
        var activities = ActivityDataBackupTool.LoadTemplateActivitiesFromFile(templatePath, null);
        Assert.NotNull(activities);
        var activity = Assert.Single(activities!);
        Assert.Equal("2001", activity.Id);
        Assert.Equal(("10", "20"), activity.RangeByTable["Icon.xlsx"]);
    }
}
