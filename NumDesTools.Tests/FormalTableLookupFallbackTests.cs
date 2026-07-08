using OfficeOpenXml;

namespace NumDesTools.Tests;

public class FormalTableLookupFallbackTests : IDisposable
{
    static FormalTableLookupFallbackTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly List<string> _roots = [];

    private string CreateRoot()
    {
        var root = Path.Combine(
            Path.GetTempPath(),
            "NumDesTools.Tests",
            Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(root);
        _roots.Add(root);
        return root;
    }

    private static string CreateWorkbook(string path, Action<ExcelWorksheet> build)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        build(ws);
        pkg.SaveAs(new FileInfo(path));
        return path;
    }

    [Fact]
    public void FindKeyRow_DoesNotTouchBackup_WhenLiveTableContainsValue()
    {
        FormalTableLookupFallback.ResetBackupCandidateCache();

        var liveRoot = CreateRoot();
        var backupRoot = CreateRoot();
        CreateWorkbook(
            Path.Combine(liveRoot, "Icon.xlsx"),
            ws =>
            {
                // NPOI row/col 是 0-based，所以写到 EPPlus 的第4行第3列。
                ws.Cells[4, 3].Value = "live-id";
            }
        );

        var ex = Record.Exception(() =>
            Assert.Equal(
                3,
                FormalTableLookupFallback.FindKeyRow(
                    liveRoot,
                    "Icon.xlsx",
                    2,
                    "live-id",
                    "Sheet1",
                    rootsProvider: () => (backupRoot, liveRoot),
                    backupCandidatesProvider: (_, _) =>
                        throw new InvalidOperationException("backup should not be touched")
                )
            )
        );

        Assert.Null(ex);
    }

    [Fact]
    public void FindKeyRow_FallsBackToBackup_WhenLiveTableMissingValue()
    {
        FormalTableLookupFallback.ResetBackupCandidateCache();

        var liveRoot = CreateRoot();
        var backupRoot = CreateRoot();
        var backupFile = Path.Combine(backupRoot, "Icon_backup_2026-07-08.xlsx");

        CreateWorkbook(
            Path.Combine(liveRoot, "Icon.xlsx"),
            ws =>
            {
                ws.Cells[4, 3].Value = "other-id";
            }
        );
        CreateWorkbook(
            backupFile,
            ws =>
            {
                ws.Cells[4, 3].Value = "backup-id";
            }
        );

        var backupCalls = 0;
        var row = FormalTableLookupFallback.FindKeyRow(
            liveRoot,
            "Icon.xlsx",
            2,
            "backup-id",
            "Sheet1",
            rootsProvider: () => (backupRoot, liveRoot),
            backupCandidatesProvider: (root, file) =>
            {
                backupCalls++;
                Assert.Equal(backupRoot, root);
                Assert.Equal("Icon.xlsx", file);
                return [backupFile];
            }
        );

        Assert.Equal(3, row);
        Assert.Equal(1, backupCalls);
    }

    public void Dispose()
    {
        foreach (var root in _roots)
        {
            try
            {
                if (Directory.Exists(root))
                    Directory.Delete(root, recursive: true);
            }
            catch
            {
                // ignore
            }
        }
    }
}
