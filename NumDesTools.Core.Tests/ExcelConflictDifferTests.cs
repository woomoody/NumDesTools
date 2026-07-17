using NumDesTools.ConflictResolver;
using OfficeOpenXml;

namespace NumDesTools.Tests;

/// <summary>
/// 通过公共接口 ExcelConflictDiffer.Diff 测试 Diff 行为。
/// 测试不关心内部用 EPPlus 还是 MiniExcel 读取——只验证输出正确性。
/// </summary>
public class ExcelConflictDifferTests : IDisposable
{
    static ExcelConflictDifferTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly List<string> _tmpFiles = [];

    // ── helpers ──────────────────────────────────────────────────────────────

    /// 创建标准 4-header config xlsx：row2=列名, row3=类型, row4=标签, row5+=数据
    private string MakeXlsx(params (string id, string name, string? extra)[] rows)
    {
        var path = Path.GetTempFileName() + ".xlsx";
        _tmpFiles.Add(path);

        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");

        ws.Cells[2, 1].Value = "#注";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 3].Value = "name";
        ws.Cells[2, 4].Value = "extra";

        ws.Cells[3, 2].Value = "int";
        ws.Cells[3, 3].Value = "string";
        ws.Cells[3, 4].Value = "string";

        ws.Cells[4, 2].Value = "编号";
        ws.Cells[4, 3].Value = "名称";
        ws.Cells[4, 4].Value = "备注";

        int r = 5;
        foreach (var (id, name, extra) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 3].Value = name;
            if (extra != null)
                ws.Cells[r, 4].Value = extra;
            r++;
        }

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private string MakeXlsx(params (string id, string name)[] rows) =>
        MakeXlsx(rows.Select(r => (r.id, r.name, (string?)null)).ToArray());

    // ── Tracer bullet ─────────────────────────────────────────────────────────

    [Fact]
    public void Diff_IdenticalFiles_NoConflicts()
    {
        var ours = MakeXlsx(("1001", "活动A"), ("1002", "活动B"));
        var theirs = MakeXlsx(("1001", "活动A"), ("1002", "活动B"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs);

        Assert.Equal(0, diff.TotalConflictRows);
    }

    // ── Modified 行 ──────────────────────────────────────────────────────────

    [Fact]
    public void Diff_ModifiedCell_DetectsChange()
    {
        var ours = MakeXlsx(("1001", "旧名称"));
        var theirs = MakeXlsx(("1001", "新名称"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs);
        var sheet = diff.Sheets.Single();
        var row = sheet.Rows.Single(r => r.DiffType == RowDiffType.Modified);

        Assert.Equal("1001", row.RowKey);
        Assert.Single(row.Cells);
        Assert.Equal("name", row.Cells[0].ColName);
        Assert.Equal("旧名称", row.Cells[0].OursValue?.ToString());
        Assert.Equal("新名称", row.Cells[0].TheirsValue?.ToString());
    }

    [Fact]
    public void Diff_UnchangedRows_AreMarkedSame()
    {
        var ours = MakeXlsx(("1001", "活动A"), ("1002", "活动B"));
        var theirs = MakeXlsx(("1001", "活动A"), ("1002", "活动B_改"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs);
        var sheet = diff.Sheets.Single();

        Assert.Contains(sheet.Rows, r => r.RowKey == "1001" && r.DiffType == RowDiffType.Same);
        Assert.Contains(sheet.Rows, r => r.RowKey == "1002" && r.DiffType == RowDiffType.Modified);
    }

    // ── OnlyOurs / OnlyTheirs ─────────────────────────────────────────────────

    [Fact]
    public void Diff_OnlyOursRow_DetectsAddedByOurs()
    {
        var ours = MakeXlsx(("1001", "A"), ("9999", "仅我有"));
        var theirs = MakeXlsx(("1001", "A"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs);
        var sheet = diff.Sheets.Single();
        var row = sheet.Rows.Single(r => r.DiffType == RowDiffType.OnlyOurs);

        Assert.Equal("9999", row.RowKey);
    }

    [Fact]
    public void Diff_OnlyTheirsRow_DetectsAddedByTheirs()
    {
        var ours = MakeXlsx(("1001", "A"));
        var theirs = MakeXlsx(("1001", "A"), ("8888", "仅对方有"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs);
        var sheet = diff.Sheets.Single();
        var row = sheet.Rows.Single(r => r.DiffType == RowDiffType.OnlyTheirs);

        Assert.Equal("8888", row.RowKey);
    }

    // ── 三方预选 ──────────────────────────────────────────────────────────────

    [Fact]
    public void Diff_WithBase_PreselectedOursSideChange()
    {
        // base = "原值"，ours = "A 改了"，theirs = "原值" → 应预选 Ours
        var basePath = MakeXlsx(("1001", "原值"));
        var ours = MakeXlsx(("1001", "A 改了"));
        var theirs = MakeXlsx(("1001", "原值"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs, basePath);
        var sheet = diff.Sheets.Single();
        var row = sheet.Rows.Single(r => r.DiffType == RowDiffType.Modified);
        var cell = row.Cells.Single();

        Assert.True(cell.IsExplicit, "单方修改应被自动预选");
        Assert.Equal(ConflictChoice.Ours, cell.Choice);
    }

    [Fact]
    public void Diff_WithBase_BothChangedNotPreselected()
    {
        // base = "原值"，ours = "A 改了"，theirs = "B 也改了" → 不预选（人工选）
        var basePath = MakeXlsx(("1001", "原值"));
        var ours = MakeXlsx(("1001", "A 改了"));
        var theirs = MakeXlsx(("1001", "B 也改了"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs, basePath);
        var sheet = diff.Sheets.Single();
        var row = sheet.Rows.Single(r => r.DiffType == RowDiffType.Modified);
        var cell = row.Cells.Single();

        Assert.False(cell.IsExplicit, "双方都改了不应自动预选");
    }

    // ── RowOrigin 三方推断 ────────────────────────────────────────────────────

    [Fact]
    public void Diff_WithBase_OnlyOursNotInBase_IsAddedByOurs()
    {
        var basePath = MakeXlsx(("1001", "A"));
        var ours = MakeXlsx(("1001", "A"), ("2000", "A新增"));
        var theirs = MakeXlsx(("1001", "A"));

        var diff = ExcelConflictDiffer.Diff(ours, theirs, basePath);
        var row = diff.Sheets.Single().Rows.Single(r => r.RowKey == "2000");

        Assert.Equal(RowOrigin.AddedByOurs, row.Origin);
        Assert.Equal(ConflictChoice.Ours, row.RowChoice); // 默认保留
    }

    [Fact]
    public void Diff_WithBase_OnlyOursInBase_IsDeletedByTheirs()
    {
        var basePath = MakeXlsx(("1001", "A"), ("3000", "B 删了"));
        var ours = MakeXlsx(("1001", "A"), ("3000", "B 删了"));
        var theirs = MakeXlsx(("1001", "A")); // B 被 theirs 删了

        var diff = ExcelConflictDiffer.Diff(ours, theirs, basePath);
        var row = diff.Sheets.Single().Rows.Single(r => r.RowKey == "3000");

        Assert.Equal(RowOrigin.DeletedByTheirs, row.Origin);
        Assert.Equal(ConflictChoice.Theirs, row.RowChoice); // 默认接受删除
    }

    // ── type/label 行合并 ────────────────────────────────────────────────────

    [Fact]
    public void Diff_TheirsNewColumn_TypeRowContainsTheirsType()
    {
        // OURS 没有 extra 列，THEIRS 有 extra 列（string 类型）
        var ours = MakeXlsx(("1001", "A")); // 只有 id + name
        var theirs = MakeXlsx(("1001", "A", "extra值")); // 多了 extra 列

        var diff = ExcelConflictDiffer.Diff(ours, theirs);
        var sheet = diff.Sheets.Single();

        // TypeRow 应包含 THEIRS 新列的类型，以便 Apply 阶段 EnsureNewColsMeta 能填入
        Assert.True(
            sheet.TypeRow.ContainsKey("extra"),
            "THEIRS 新增列的 type 应在 TypeRow 中"
        );
        Assert.Equal("string", sheet.TypeRow["extra"]);
    }

    public void Dispose()
    {
        foreach (var f in _tmpFiles)
            try { File.Delete(f); } catch { }
    }
}
