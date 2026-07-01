using NumDesTools.ConflictResolver;
using OfficeOpenXml;

namespace NumDesTools.Tests.ConflictResolver;

public class ConflictApplierTests : IDisposable
{
    static ConflictApplierTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly List<string> _tmpFiles = [];

    private string MakeXlsx(params (string id, string name)[] rows)
    {
        var path = Path.GetTempFileName() + ".xlsx";
        _tmpFiles.Add(path);

        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");

        ws.Cells[2, 1].Value = "#注";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 3].Value = "name";

        ws.Cells[3, 2].Value = "int";
        ws.Cells[3, 3].Value = "string";

        ws.Cells[4, 2].Value = "编号";
        ws.Cells[4, 3].Value = "名称";

        int r = 5;
        foreach (var (id, name) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 3].Value = name;
            r++;
        }

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private string TmpOut()
    {
        var path = Path.GetTempFileName() + ".xlsx";
        _tmpFiles.Add(path);
        return path;
    }

    // ── 场景 A：Modified 行选 Theirs，单元格值写回 ──────────────────────────

    [Fact]
    public void Apply_ModifiedRow_ChoiceTheirs_CellValueUpdatedInOutput()
    {
        var oursPath = MakeXlsx(("1001", "旧名称"));
        var theirsPath = MakeXlsx(("1001", "新名称"));
        var outPath = TmpOut();

        var cell = new CellConflict
        {
            ColName = "name",
            OursValue = "旧名称",
            TheirsValue = "新名称",
            Choice = ConflictChoice.Theirs,
            IsExplicit = true,
        };
        var row = new RowConflict
        {
            RowKey = "1001",
            DiffType = RowDiffType.Modified,
            TheirsRowIndex = 0,
            OursRowIndex = 0,
        };
        row.Cells.Add(cell);

        var sheetDiff = new SheetDiff(
            "Sheet1",
            [row]
        )
        {
            AllColumns = ["#注", "id", "name"],
            TypeRow = new() { ["id"] = "int", ["name"] = "string" },
            LabelRow = new() { ["id"] = "编号", ["name"] = "名称" },
        };
        var diff = new FileDiff(oursPath, theirsPath, [sheetDiff]);

        ConflictApplier.Apply(diff, outPath, gitAdd: false);

        using var pkg = new ExcelPackage(new FileInfo(outPath));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        Assert.Equal("新名称", ws.Cells[5, 3].Value?.ToString());
    }

    // ── 场景 B：多行 Modified 全选 Theirs，每行都写对 ───────────────────────

    [Fact]
    public void Apply_MultipleModifiedRows_AllChoiceTheirs_AllCellsUpdated()
    {
        var oursPath = MakeXlsx(("1001", "旧A"), ("1002", "旧B"));
        var theirsPath = MakeXlsx(("1001", "新A"), ("1002", "新B"));
        var outPath = TmpOut();

        var row1 = new RowConflict
        {
            RowKey = "1001",
            DiffType = RowDiffType.Modified,
            TheirsRowIndex = 0,
            OursRowIndex = 0,
        };
        row1.Cells.Add(new CellConflict
        {
            ColName = "name",
            OursValue = "旧A",
            TheirsValue = "新A",
            Choice = ConflictChoice.Theirs,
            IsExplicit = true,
        });

        var row2 = new RowConflict
        {
            RowKey = "1002",
            DiffType = RowDiffType.Modified,
            TheirsRowIndex = 1,
            OursRowIndex = 1,
        };
        row2.Cells.Add(new CellConflict
        {
            ColName = "name",
            OursValue = "旧B",
            TheirsValue = "新B",
            Choice = ConflictChoice.Theirs,
            IsExplicit = true,
        });

        var sheetDiff = new SheetDiff("Sheet1", [row1, row2])
        {
            AllColumns = ["#注", "id", "name"],
            TypeRow = new() { ["id"] = "int", ["name"] = "string" },
            LabelRow = new() { ["id"] = "编号", ["name"] = "名称" },
        };
        var diff = new FileDiff(oursPath, theirsPath, [sheetDiff]);

        ConflictApplier.Apply(diff, outPath, gitAdd: false);

        using var pkg = new ExcelPackage(new FileInfo(outPath));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        Assert.Equal("新A", ws.Cells[5, 3].Value?.ToString());
        Assert.Equal("新B", ws.Cells[6, 3].Value?.ToString());
    }

    // ── 场景 C：OnlyTheirs 行选 Theirs，新行插入正确位置 ─────────────────────

    [Fact]
    public void Apply_OnlyTheirsRow_ChoiceTheirs_RowInsertedBetweenNeighbors()
    {
        var oursPath = MakeXlsx(("1001", "A"), ("1002", "B"));
        var theirsPath = MakeXlsx(("1001", "A"), ("8888", "新行"), ("1002", "B"));
        var outPath = TmpOut();

        var same1 = new RowConflict
        {
            RowKey = "1001",
            DiffType = RowDiffType.Same,
            TheirsRowIndex = 0,
            OursRowIndex = 0,
        };
        var inserted = new RowConflict
        {
            RowKey = "8888",
            DiffType = RowDiffType.OnlyTheirs,
            TheirsRowIndex = 1,
            OursRowIndex = -1,
            DefaultRowChoice = ConflictChoice.Theirs,
            TheirsFullRow = new Dictionary<string, object?>
            {
                ["id"] = "8888",
                ["name"] = "新行",
            },
        };
        var same2 = new RowConflict
        {
            RowKey = "1002",
            DiffType = RowDiffType.Same,
            TheirsRowIndex = 2,
            OursRowIndex = 1,
        };

        var sheetDiff = new SheetDiff("Sheet1", [same1, inserted, same2])
        {
            AllColumns = ["#注", "id", "name"],
            TypeRow = new() { ["id"] = "int", ["name"] = "string" },
            LabelRow = new() { ["id"] = "编号", ["name"] = "名称" },
        };
        var diff = new FileDiff(oursPath, theirsPath, [sheetDiff]);

        ConflictApplier.Apply(diff, outPath, gitAdd: false);

        using var pkg = new ExcelPackage(new FileInfo(outPath));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var keys = Enumerable
            .Range(5, ws.Dimension.End.Row - 4)
            .Select(r => ws.Cells[r, 2].Value?.ToString())
            .ToList();

        Assert.Equal(3, keys.Count);
        Assert.Equal("1001", keys[0]);
        Assert.Equal("8888", keys[1]);
        Assert.Equal("1002", keys[2]);
    }

    // ── 场景 D：OnlyOurs 行选 Theirs（接受删除），行从输出消失 ──────────────

    [Fact]
    public void Apply_OnlyOursRow_ChoiceTheirs_RowDeletedFromOutput()
    {
        var oursPath = MakeXlsx(("1001", "A"), ("9999", "要删除"));
        var theirsPath = MakeXlsx(("1001", "A"));
        var outPath = TmpOut();

        var same = new RowConflict
        {
            RowKey = "1001",
            DiffType = RowDiffType.Same,
            TheirsRowIndex = 0,
            OursRowIndex = 0,
        };
        var deleted = new RowConflict
        {
            RowKey = "9999",
            DiffType = RowDiffType.OnlyOurs,
            TheirsRowIndex = -1,
            OursRowIndex = 1,
            DefaultRowChoice = ConflictChoice.Theirs,
        };

        var sheetDiff = new SheetDiff("Sheet1", [same, deleted])
        {
            AllColumns = ["#注", "id", "name"],
            TypeRow = new() { ["id"] = "int", ["name"] = "string" },
            LabelRow = new() { ["id"] = "编号", ["name"] = "名称" },
        };
        var diff = new FileDiff(oursPath, theirsPath, [sheetDiff]);

        ConflictApplier.Apply(diff, outPath, gitAdd: false);

        using var pkg = new ExcelPackage(new FileInfo(outPath));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var keys = Enumerable
            .Range(5, ws.Dimension.End.Row - 4)
            .Select(r => ws.Cells[r, 2].Value?.ToString())
            .Where(k => !string.IsNullOrEmpty(k))
            .ToList();

        Assert.DoesNotContain("9999", keys);
        Assert.Contains("1001", keys);
    }

    public void Dispose()
    {
        foreach (var f in _tmpFiles)
            try
            {
                File.Delete(f);
            }
            catch { }
    }
}
