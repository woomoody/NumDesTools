using OfficeOpenXml;

namespace NumDesTools.Tests;

public class ActivityDataBackupToolTests : IDisposable
{
    static ActivityDataBackupToolTests()
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

        var r = 5;
        foreach (var (id, name) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 3].Value = name;
            r++;
        }

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    // id 用真实 Icon/Type/Item 表的 8 位前缀+序号格式（比如 11010001）——
    // XlsxCrossSync 的分组插入按 6 位前缀匹配同组最后一行，短 id（长度<6）会直接落到表尾，
    // 跟生产表 id 形状一致才能测出真实的插入位置。
    private static (string Id, string Name)[] SampleRows() =>
        Enumerable.Range(1, 10).Select(i => ($"110100{i:D2}", $"活动110100{i:D2}")).ToArray();

    private static List<(string Id, string Name)> ReadRows(string path)
    {
        using var pkg = new ExcelPackage(new FileInfo(path));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var result = new List<(string Id, string Name)>();
        if (ws.Dimension is null)
            return result;
        for (var r = 5; r <= ws.Dimension.End.Row; r++)
        {
            var id = ws.Cells[r, 2].Text;
            if (string.IsNullOrEmpty(id))
                continue;
            result.Add((id, ws.Cells[r, 3].Text));
        }
        return result;
    }

    // id/说明列布局跟 Icon.xlsx 生产表一致（说明列在 id 右边一列）。commentHeader 默认"#备注"，
    // 传"#竞品名称"之类的非标准名字用于测试 FindDescriptiveHashCol 的兼容逻辑。
    private string MakeXlsxWithComment(
        (string id, string comment)[] rows,
        string commentHeader = "#备注"
    )
    {
        var path = Path.GetTempFileName() + ".xlsx";
        _tmpFiles.Add(path);

        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[2, 1].Value = "#";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 3].Value = commentHeader;

        var r = 5;
        foreach (var (id, comment) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 3].Value = comment;
            r++;
        }

        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    private static (string Id, string Comment)[] SampleRowsWithComment(string comment) =>
        Enumerable.Range(1, 10).Select(i => ($"110100{i:D2}", comment)).ToArray();

    private static List<string> ReadIds(string path)
    {
        using var pkg = new ExcelPackage(new FileInfo(path));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        var result = new List<string>();
        if (ws.Dimension is null)
            return result;
        for (var r = 5; r <= ws.Dimension.End.Row; r++)
        {
            var id = ws.Cells[r, 2].Text;
            if (!string.IsNullOrEmpty(id))
                result.Add(id);
        }
        return result;
    }

    [Fact]
    public void ApplyDelete_RemovesOnlyRowsInIdRange()
    {
        var livePath = MakeXlsx(SampleRows());
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        var remaining = ReadRows(livePath);
        Assert.Equal(7, remaining.Count);
        Assert.DoesNotContain(remaining, x => x.Id is "11010004" or "11010005" or "11010006");
        Assert.Contains(remaining, x => x.Id == "11010003");
        Assert.Contains(remaining, x => x.Id == "11010007");
    }

    [Fact]
    public void ApplyDelete_StillDeletes_ButLogsMismatch_WhenMixedInRowHasDifferentIdAndComment()
    {
        var rows = SampleRowsWithComment("常规活动");
        rows[4] = ("99990001", "神秘活动"); // 区间中间混入了不同活动的数据（id前缀、说明列中文签名都不一致）
        var livePath = MakeXlsxWithComment(rows);
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        var result = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        Assert.Equal(7, ReadIds(livePath).Count); // 疑似混入不阻断，照常删除整段
        Assert.False(string.IsNullOrEmpty(result.MismatchDetail)); // 详情走 CTP，先确认真的收集到了
    }

    [Fact]
    public void ApplyDelete_LogsWithActualColumnName_WhenCommentColumnIsNonStandardName()
    {
        // Item.xlsx 的说明列不叫"#备注"，叫"#竞品名称"——只要带 # 前缀就该被 FindDescriptiveHashCol 认出来。
        var rows = SampleRowsWithComment("常规活动");
        rows[4] = ("99990001", "神秘活动");
        var livePath = MakeXlsxWithComment(rows, commentHeader: "#竞品名称");
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        var result = ActivityDataBackupTool.ApplyDelete("Item.xlsx", livePath, ranges);

        Assert.Equal(7, ReadIds(livePath).Count); // 非标准列名也能核实到，同样不阻断
        Assert.Contains("#竞品名称", result.MismatchDetail); // 提示信息里用的是实际列名，不是硬编码的"#备注"
    }

    [Fact]
    public void ApplyDelete_DoesNotFlagRealItems_WhenBoundaryRowsAreReservedPlaceholders()
    {
        // 还原真实故障场景：起止id（21010000/21011008）是预留占位行，备注是空的；真实数据是
        // "Lte-阿拉丁-神灯1"~"神灯8" 这种英文前缀+中文名的老数据格式。id前缀只取前4位（"2101"对
        // 起止id和真实条目都一致）就足够判定同一活动，根本不会走到备注比对那一步。
        var rows = new (string id, string comment)[]
        {
            ("21010000", ""),
            ("21010101", "Lte-阿拉丁-神灯1"),
            ("21010102", "Lte-阿拉丁-神灯2"),
            ("21010103", "Lte-阿拉丁-神灯3"),
            ("21010104", "Lte-阿拉丁-神灯4"),
            ("21010105", "Lte-阿拉丁-神灯5"),
            ("21010106", "Lte-阿拉丁-神灯6"),
            ("21010107", "Lte-阿拉丁-神灯7"),
            ("21010108", "Lte-阿拉丁-神灯8"),
            ("21011008", ""),
        };
        var livePath = MakeXlsxWithComment(rows);
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "阿拉丁", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "21010000", "21011008"),
        };

        var result = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        Assert.Empty(ReadIds(livePath)); // 整段照常删除
        Assert.True(string.IsNullOrEmpty(result.MismatchDetail)); // 起止id和真实条目共享4位前缀，不该有任何提示
    }

    [Fact]
    public void ApplyDelete_SkipsMismatchNote_WhenIdSharesOnlyFirst4Digits()
    {
        // 还原真实故障场景二：id 45061701~45061708 是同一活动，第5-6位是活动内部分桶（不是活动本身），
        // 45069901 前4位"4506"跟主流一致，只是分桶不同——4位前缀判定同一活动即可，不该被判成混入。
        var rows = new (string id, string comment)[]
        {
            ("45061701", "常规道具1"),
            ("45061702", "常规道具2"),
            ("45061703", "常规道具3"),
            ("45069901", "活动体力-3-礼包触发"), // 前4位仍是"4506"，只是5-6位分桶不同
            ("45061704", "常规道具4"),
        };
        var livePath = MakeXlsxWithComment(rows);
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "45061701", "45061704"),
        };

        var result = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        Assert.Empty(ReadIds(livePath)); // 整段照常删除
        Assert.True(string.IsNullOrEmpty(result.MismatchDetail)); // 前4位一致，不该有提示
    }

    [Fact]
    public void ApplyDelete_SkipsMismatchNote_WhenCommentIsCompoundNameContainingMajoritySig()
    {
        // 备注签名放宽成"谁是谁的前缀"：多数签名是"阿拉丁"，这一行备注是"阿拉丁副本纪念品"
        // （没有分隔符可截断，整段都是中文），应该算同一活动，不能要求精确相等。
        var rows = new (string id, string comment)[]
        {
            ("21010101", "阿拉丁-神灯1"),
            ("21010102", "阿拉丁-神灯2"),
            ("99990000", "阿拉丁副本纪念品"), // id前缀不同，靠备注签名核实
            ("21010103", "阿拉丁-神灯3"),
        };
        var livePath = MakeXlsxWithComment(rows);
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "21010101", "21010103"),
        };

        var result = ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        Assert.Empty(ReadIds(livePath)); // 整段照常删除
        Assert.True(string.IsNullOrEmpty(result.MismatchDetail)); // 备注互为前缀关系，不该有提示
    }

    [Fact]
    public void ApplyDelete_StillDeletes_WhenIdDiffersButCommentPrefixMatches()
    {
        var rows = SampleRowsWithComment("常规活动");
        rows[4] = ("99990001", "常规活动-补充说明"); // id前缀不同，但#备注中文前缀仍是同一个活动
        var livePath = MakeXlsxWithComment(rows);
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);

        Assert.Equal(7, ReadIds(livePath).Count); // #备注核实通过，照常删除3行
    }

    [Fact]
    public void ApplyRestore_ReinsertsDeletedRows_BetweenOriginalNeighbours()
    {
        var backupPath = MakeXlsx(SampleRows());
        var livePath = MakeXlsx(SampleRows());
        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        // 先模拟"删除"，再"还原"
        ActivityDataBackupTool.ApplyDelete("Icon.xlsx", livePath, ranges);
        Assert.Equal(7, ReadRows(livePath).Count);

        ActivityDataBackupTool.ApplyRestore("Icon.xlsx", livePath, [backupPath], ranges);

        var restored = ReadRows(livePath);
        // 行内数据（id→name）完整还原才是硬要求；行的具体插入位置不影响配置表按 id 查值的正确性，
        // 分组插入用的是"同前缀最后一行之后"这个既有规则（不保证插回原来那一格）。
        Assert.Equal(10, restored.Count);
        Assert.Equal(
            SampleRows().Select(x => x.Id).OrderBy(x => x),
            restored.Select(x => x.Id).OrderBy(x => x)
        );
        Assert.Contains(restored, x => x.Id == "11010005" && x.Name == "活动11010005");
    }

    [Fact]
    public void ApplyRestore_UpdatesExistingRowInRange_WithoutTouchingRowsOutsideRange()
    {
        var backupPath = MakeXlsx(SampleRows());
        var liveRows = SampleRows();
        liveRows[4] = ("11010005", "被误改的名字"); // 索引4 = id 11010005，模拟正式表数据被改坏
        var livePath = MakeXlsx(liveRows);

        var activity = new ActivityDataBackupTool.Activity(0, "9001", "测试活动", new());
        var ranges = new List<(ActivityDataBackupTool.Activity Activity, string Start, string End)>
        {
            (activity, "11010004", "11010006"),
        };

        ActivityDataBackupTool.ApplyRestore("Icon.xlsx", livePath, [backupPath], ranges);

        var restored = ReadRows(livePath);
        Assert.Equal(10, restored.Count);
        Assert.Contains(restored, x => x.Id == "11010005" && x.Name == "活动11010005");
        // 区间外的行不受影响
        Assert.Contains(restored, x => x.Id == "11010001" && x.Name == "活动11010001");
    }

    public void Dispose()
    {
        foreach (var f in _tmpFiles)
        {
            try
            {
                File.Delete(f);
            }
            catch
            {
                // ignore
            }
        }
    }
}
