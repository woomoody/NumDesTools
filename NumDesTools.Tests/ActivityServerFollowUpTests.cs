using NumDesTools.AutoInsert;
using OfficeOpenXml;

namespace NumDesTools.Tests;

public class ActivityServerFollowUpTests : IDisposable
{
    static ActivityServerFollowUpTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly string _tempDir;

    public ActivityServerFollowUpTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private string MakeFollowUpXlsx(params (string id, string activityIds)[] rows)
    {
        var path = Path.Combine(_tempDir, "ActivityClientFollowUpData.xlsx");
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "#";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 6].Value = "activityIds";
        var r = 5;
        foreach (var (id, activityIds) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 6].Value = activityIds;
            r++;
        }
        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    [Fact]
    public void BuildFollowUpTargetMap_EmptyDir_ReturnsEmpty()
    {
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpTargetMap(_tempDir);
        Assert.Empty(map);
    }

    [Fact]
    public void BuildFollowUpTargetMap_SingleMapping()
    {
        MakeFollowUpXlsx(("80002", "[80001]"));
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpTargetMap(_tempDir);
        Assert.Single(map);
        Assert.Equal("80002", map["80001"].Single());
    }

    [Fact]
    public void BuildFollowUpTargetMap_MultipleTargets()
    {
        MakeFollowUpXlsx(("56002", "[56003,56004,56002]"));
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpTargetMap(_tempDir);
        Assert.Equal(3, map.Count);
    }

    [Fact]
    public void BuildFollowUpTargetMap_MultiplePredecessors()
    {
        MakeFollowUpXlsx(("80002", "[80001]"), ("80005", "[80001]"));
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpTargetMap(_tempDir);
        Assert.Equal(2, map["80001"].Count);
    }

    [Fact]
    public void BuildFollowUpPredecessorMap_SingleMapping()
    {
        MakeFollowUpXlsx(("80002", "[80001]"));
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        Assert.Single(map);
        Assert.Equal("80001", map["80002"].Single());
    }

    [Fact]
    public void BuildFollowUpPredecessorMap_MultipleTargets()
    {
        MakeFollowUpXlsx(("56002", "[56003,56004,56002]"));
        var map = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        Assert.Single(map);
        Assert.Equal(3, map["56002"].Count);
    }

    [Fact]
    public void FollowUpCheck_TargetInBatch_NoWarning()
    {
        // 80001 是 80002 的续开目标，两者都在 batch 里 → 无警告
        MakeFollowUpXlsx(("80002", "[80001]"));
        var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        var batch = new HashSet<string> { "80001", "80002" };

        var warnings = new List<string>();
        foreach (var id in batch)
        {
            if (predMap.TryGetValue(id, out var targets))
            {
                var missing = targets.Where(t => !batch.Contains(t)).Distinct().ToList();
                if (missing.Count > 0)
                    warnings.Add($"missing: {string.Join(",", missing)}");
            }
        }
        Assert.Empty(warnings);
    }

    [Fact]
    public void FollowUpCheck_TargetNotInBatch_Warning()
    {
        // 80002 的续开目标是 80001，但 80001 不在 batch 里 → 警告
        MakeFollowUpXlsx(("80002", "[80001]"));
        var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        var batch = new HashSet<string> { "80002" }; // 80001 not in batch

        var warnings = new List<string>();
        foreach (var id in batch)
        {
            if (predMap.TryGetValue(id, out var targets))
            {
                var missing = targets.Where(t => !batch.Contains(t)).Distinct().ToList();
                if (missing.Count > 0)
                    warnings.Add($"missing: {string.Join(",", missing)}");
            }
        }
        Assert.Single(warnings);
        Assert.Contains("80001", warnings[0]);
    }

    [Fact]
    public void FollowUpCheck_NoFollowUpData_NoWarning()
    {
        var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        Assert.Empty(predMap);
    }

    [Fact]
    public void FollowUpCheck_MultipleTargetsPartialBatch_Warning()
    {
        // 56002 → [56003, 56004, 56002], batch 只有 56002 和 56003 → 56004 缺失
        MakeFollowUpXlsx(("56002", "[56003,56004,56002]"));
        var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        var batch = new HashSet<string> { "56002", "56003" };

        var warnings = new List<string>();
        foreach (var id in batch)
        {
            if (predMap.TryGetValue(id, out var targets))
            {
                var missing = targets.Where(t => !batch.Contains(t)).Distinct().ToList();
                if (missing.Count > 0)
                    warnings.Add($"missing: {string.Join(",", missing)}");
            }
        }
        Assert.Single(warnings);
        Assert.Contains("56004", warnings[0]);
    }

    [Fact]
    public void FollowUpCheck_AllTargetsInBatch_NoWarning()
    {
        // 56002 → [56003, 56004], all in batch
        MakeFollowUpXlsx(("56002", "[56003,56004]"));
        var predMap = ExcelDataAutoInsertActivityServer.BuildFollowUpPredecessorMap(_tempDir);
        var batch = new HashSet<string> { "56002", "56003", "56004" };

        var warnings = new List<string>();
        foreach (var id in batch)
        {
            if (predMap.TryGetValue(id, out var targets))
            {
                var missing = targets.Where(t => !batch.Contains(t)).Distinct().ToList();
                if (missing.Count > 0)
                    warnings.Add("warning");
            }
        }
        Assert.Empty(warnings);
    }
}
