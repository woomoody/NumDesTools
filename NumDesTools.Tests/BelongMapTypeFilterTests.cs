using OfficeOpenXml;

namespace NumDesTools.Tests;

public class BelongMapTypeFilterTests : IDisposable
{
    static BelongMapTypeFilterTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly string _tempDir;

    public BelongMapTypeFilterTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private static ActivityDataBackupTool.Activity DummyAct(string id) =>
        new(0, id, "", new Dictionary<string, (string, string)>());

    private string MakeTypeXlsx(params (string id, string belongMapType)[] rows)
    {
        var path = Path.Combine(_tempDir, "Type.xlsx");
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "#";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 7].Value = "belongMapType";
        var r = 5;
        foreach (var (id, bmt) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 7].Value = bmt;
            r++;
        }
        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private string MakeTableXlsx(string fileName, params string[] ids)
    {
        var path = Path.Combine(_tempDir, fileName);
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "#";
        ws.Cells[2, 2].Value = "id";
        var r = 5;
        foreach (var id in ids)
        {
            ws.Cells[r, 2].Value = id;
            r++;
        }
        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private List<(ActivityDataBackupTool.Activity, string, string)> R(string start, string end) =>
        [(DummyAct("x"), start, end)];

    [Fact]
    public void LteOnlyIds_AreDeletable()
    {
        MakeTypeXlsx(("21010001", "[4]"), ("21010002", "[4]"), ("21010003", "[4]"));
        MakeTableXlsx("Item.xlsx", "21010001", "21010002", "21010003");
        var (d, p) = ActivityDataBackupTool.BuildDeletableIdSet(
            Path.Combine(_tempDir, "Item.xlsx"),
            Path.Combine(_tempDir, "Type.xlsx"),
            R("21010001", "21010003")
        );
        var flat = d.Values.SelectMany(ids => ids).ToHashSet();
        Assert.Equal(3, flat.Count);
        Assert.Empty(p);
    }

    [Fact]
    public void NonLteId_IsPreserved()
    {
        MakeTypeXlsx(("21010000", "[1]"), ("21010001", "[4]"), ("21010002", "[4]"));
        MakeTableXlsx("Item.xlsx", "21010000", "21010001", "21010002");
        var (d, p) = ActivityDataBackupTool.BuildDeletableIdSet(
            Path.Combine(_tempDir, "Item.xlsx"),
            Path.Combine(_tempDir, "Type.xlsx"),
            R("21010000", "21010002")
        );
        var flat = d.Values.SelectMany(ids => ids).ToHashSet();
        Assert.Equal(2, flat.Count);
        Assert.Single(p);
        Assert.DoesNotContain("21010000", flat);
        Assert.Contains(p, x => x.Contains("21010000") && x.Contains("[1]"));
    }

    [Fact]
    public void MixedBelongMapType_IsPreserved()
    {
        MakeTypeXlsx(("21010000", "[1,2,4]"));
        MakeTableXlsx("Item.xlsx", "21010000");
        var (d, p) = ActivityDataBackupTool.BuildDeletableIdSet(
            Path.Combine(_tempDir, "Item.xlsx"),
            Path.Combine(_tempDir, "Type.xlsx"),
            R("21010000", "21010000")
        );
        Assert.Empty(d.Values.SelectMany(ids => ids));
        Assert.Single(p);
    }

    [Fact]
    public void Item_PreservesNonDeletableIds()
    {
        MakeTableXlsx("Item.xlsx", "21010000", "21010001", "21010002");
        var ids = new HashSet<string> { "21010001", "21010002" };
        ActivityDataBackupTool.ApplyDeleteWithIdFilter(
            "Item.xlsx",
            Path.Combine(_tempDir, "Item.xlsx"),
            ids,
            []
        );
        using var pkg = new ExcelPackage(new FileInfo(Path.Combine(_tempDir, "Item.xlsx")));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var remaining = new List<string>();
        for (var r = 5; r <= ws.Dimension?.End.Row; r++)
        {
            var id = ws.Cells[r, 2].Text?.Trim();
            if (!string.IsNullOrEmpty(id))
                remaining.Add(id);
        }
        Assert.Contains("21010000", remaining);
        Assert.DoesNotContain("21010001", remaining);
    }

    [Fact]
    public void Type_PreservesNonDeletableIds()
    {
        MakeTypeXlsx(("21010000", "[1]"), ("21010001", "[4]"));
        var ids = new HashSet<string> { "21010001" };
        ActivityDataBackupTool.ApplyDeleteWithIdFilter(
            "Type.xlsx",
            Path.Combine(_tempDir, "Type.xlsx"),
            ids,
            []
        );
        using var pkg = new ExcelPackage(new FileInfo(Path.Combine(_tempDir, "Type.xlsx")));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var remaining = new List<string>();
        for (var r = 5; r <= ws.Dimension?.End.Row; r++)
        {
            var id = ws.Cells[r, 2].Text?.Trim();
            if (!string.IsNullOrEmpty(id))
                remaining.Add(id);
        }
        Assert.Contains("21010000", remaining);
        Assert.DoesNotContain("21010001", remaining);
    }

    [Fact]
    public void Icon_PreservesNonDeletableIds()
    {
        MakeTableXlsx("Icon.xlsx", "21010000", "21010001");
        var ids = new HashSet<string> { "21010001" };
        ActivityDataBackupTool.ApplyDeleteWithIdFilter(
            "Icon.xlsx",
            Path.Combine(_tempDir, "Icon.xlsx"),
            ids,
            []
        );
        using var pkg = new ExcelPackage(new FileInfo(Path.Combine(_tempDir, "Icon.xlsx")));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var remaining = new List<string>();
        for (var r = 5; r <= ws.Dimension?.End.Row; r++)
        {
            var id = ws.Cells[r, 2].Text?.Trim();
            if (!string.IsNullOrEmpty(id))
                remaining.Add(id);
        }
        Assert.Contains("21010000", remaining);
        Assert.DoesNotContain("21010001", remaining);
    }

    [Fact]
    public void AllThreeTables_UseSameDeletableIds_NoItemSpecialCase()
    {
        MakeTypeXlsx(("21010000", "[1]"), ("21010001", "[4]"));
        MakeTableXlsx("Item.xlsx", "21010000", "21010001");
        MakeTableXlsx("Icon.xlsx", "21010000", "21010001");

        var ids = new HashSet<string> { "21010001" };
        foreach (var table in new[] { "Item.xlsx", "Type.xlsx", "Icon.xlsx" })
        {
            ActivityDataBackupTool.ApplyDeleteFiltered(
                table,
                Path.Combine(_tempDir, table),
                [],
                ids,
                []
            );
        }

        // 三个表都应该保留 21010000，删除 21010001
        foreach (var table in new[] { "Item.xlsx", "Type.xlsx", "Icon.xlsx" })
        {
            using var pkg = new ExcelPackage(new FileInfo(Path.Combine(_tempDir, table)));
            var ws = pkg.Workbook.Worksheets["Sheet1"];
            var remaining = new List<string>();
            for (var r = 5; r <= ws.Dimension?.End.Row; r++)
            {
                var id = ws.Cells[r, 2].Text?.Trim();
                if (!string.IsNullOrEmpty(id))
                    remaining.Add(id);
            }
            Assert.Contains("21010000", remaining);
            Assert.DoesNotContain("21010001", remaining);
        }
    }
}
