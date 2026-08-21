using System.Diagnostics;
using LibGit2Sharp;
using NumDesTools.ConflictResolver;
using OfficeOpenXml;

namespace NumDesTools.Tests.ConflictResolver;

public sealed class SemanticNoOpTests : IDisposable
{
    static SemanticNoOpTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly List<string> _temporaryPaths = [];

    [Fact]
    public void IsSemanticNoOp_WhenSourceChangeIsAlreadyInTarget_IgnoresOoxmlDifferences()
    {
        var basePath = CreateWorkbook(("Sheet1", "1001", "base"));
        var sourcePath = CreateWorkbook(("Sheet1", "1001", "source"));
        var targetPath = CreateWorkbookWithOptions([("Sheet1", "1001", "source")], styleCell: true);

        Assert.True(ExcelConflictDiffer.IsSemanticNoOp(basePath, sourcePath, targetPath));
        Assert.NotEqual(File.ReadAllBytes(sourcePath), File.ReadAllBytes(targetPath));
    }

    [Fact]
    public void IsSemanticNoOp_WhenSourceOnlyRowIsMissingFromTarget_LeavesDifference()
    {
        var basePath = CreateWorkbook(("Sheet1", "1001", "base"));
        var sourcePath = CreateWorkbook(
            ("Sheet1", "1001", "base"),
            ("Sheet1", "2002", "source-only-row")
        );
        var targetPath = CreateWorkbook(("Sheet1", "1001", "base"));

        Assert.False(ExcelConflictDiffer.IsSemanticNoOp(basePath, sourcePath, targetPath));
    }

    [Fact]
    public void IsSemanticNoOp_WhenSourceOnlySheetIsMissingFromTarget_LeavesDifference()
    {
        var basePath = CreateWorkbook(("Sheet1", "1001", "base"));
        var sourcePath = CreateWorkbook(
            ("Sheet1", "1001", "base"),
            ("Extra", "2002", "source-only-sheet")
        );
        var targetPath = CreateWorkbook(("Sheet1", "1001", "base"));

        Assert.False(ExcelConflictDiffer.IsSemanticNoOp(basePath, sourcePath, targetPath));
    }

    [Fact]
    public void IsSemanticNoOp_WhenSourceMetadataChangeIsMissingFromTarget_LeavesDifference()
    {
        var basePath = CreateWorkbookWithOptions([("Sheet1", "1001", "base")], type: "string");
        var sourcePath = CreateWorkbookWithOptions([("Sheet1", "1001", "base")], type: "int");
        var targetPath = CreateWorkbookWithOptions([("Sheet1", "1001", "base")], type: "string");

        Assert.False(ExcelConflictDiffer.IsSemanticNoOp(basePath, sourcePath, targetPath));
    }

    [Fact]
    public void TryResolveSemanticNoOp_CopiesTargetStageBytesAndKeepsMergeBookkeeping()
    {
        var repositoryPath = Path.Combine(
            Path.GetTempPath(),
            "NumDesTools-semantic-no-op-" + Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(repositoryPath);
        _temporaryPaths.Add(repositoryPath);

        var basePath = CreateWorkbook(("Sheet1", "1001", "base"));
        var sourcePath = CreateWorkbook(
            ("Sheet1", "1001", "source"),
            ("Sheet1", "2002", "source-only")
        );
        var targetPath = CreateWorkbookWithOptions(
            [("Sheet1", "1001", "source"), ("Sheet1", "2002", "source-only")],
            styleCell: true
        );

        RunGit(repositoryPath, "init -b main");
        RunGit(repositoryPath, "config user.name test");
        RunGit(repositoryPath, "config user.email test@example.invalid");
        File.Copy(basePath, Path.Combine(repositoryPath, "config.xlsx"));
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m base");
        RunGit(repositoryPath, "switch -c source");
        File.Copy(sourcePath, Path.Combine(repositoryPath, "config.xlsx"), overwrite: true);
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m source");
        RunGit(repositoryPath, "switch main");
        File.Copy(targetPath, Path.Combine(repositoryPath, "config.xlsx"), overwrite: true);
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m target");
        RunGit(repositoryPath, "merge --no-commit source", expectSuccess: false);

        using var repo = new Repository(repositoryPath);
        var conflict = repo.Index.Conflicts["config.xlsx"]!;
        var expectedTargetBytes = ReadBlob(repo, conflict.Ours!);
        File.WriteAllBytes(Path.Combine(repositoryPath, "config.xlsx"), [1, 2, 3, 4]);

        Assert.True(ConflictApplier.TryResolveSemanticNoOp(repositoryPath, "config.xlsx"));

        Assert.Equal(
            expectedTargetBytes,
            File.ReadAllBytes(Path.Combine(repositoryPath, "config.xlsx"))
        );
        using var resolvedRepo = new Repository(repositoryPath);
        Assert.Null(resolvedRepo.Index.Conflicts["config.xlsx"]);
        Assert.Equal(
            expectedTargetBytes,
            ReadBlob(resolvedRepo, resolvedRepo.Index["config.xlsx"]!)
        );
        Assert.True(File.Exists(Path.Combine(resolvedRepo.Info.Path, "MERGE_HEAD")));
    }

    [Fact]
    public void TryResolveSemanticNoOp_WhenSourceDiffersFromTarget_LeavesConflict()
    {
        var repositoryPath = Path.Combine(
            Path.GetTempPath(),
            "NumDesTools-semantic-diff-" + Guid.NewGuid().ToString("N")
        );
        Directory.CreateDirectory(repositoryPath);
        _temporaryPaths.Add(repositoryPath);

        var basePath = CreateWorkbook(("Sheet1", "1001", "base"));
        var sourcePath = CreateWorkbook(
            ("Sheet1", "1001", "source"),
            ("Sheet1", "2002", "source-only")
        );
        var targetPath = CreateWorkbook(("Sheet1", "1001", "target"));

        RunGit(repositoryPath, "init -b main");
        RunGit(repositoryPath, "config user.name test");
        RunGit(repositoryPath, "config user.email test@example.invalid");
        File.Copy(basePath, Path.Combine(repositoryPath, "config.xlsx"));
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m base");
        RunGit(repositoryPath, "switch -c source");
        File.Copy(sourcePath, Path.Combine(repositoryPath, "config.xlsx"), overwrite: true);
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m source");
        RunGit(repositoryPath, "switch main");
        File.Copy(targetPath, Path.Combine(repositoryPath, "config.xlsx"), overwrite: true);
        RunGit(repositoryPath, "add config.xlsx");
        RunGit(repositoryPath, "commit -m target");
        RunGit(repositoryPath, "merge --no-commit source", expectSuccess: false);

        Assert.False(ConflictApplier.TryResolveSemanticNoOp(repositoryPath, "config.xlsx"));
        using var repo = new Repository(repositoryPath);
        Assert.NotNull(repo.Index.Conflicts["config.xlsx"]);
    }

    private string CreateWorkbook(params (string sheet, string id, string value)[] rows) =>
        CreateWorkbookWithOptions(rows);

    private string CreateWorkbookWithOptions(
        (string sheet, string id, string value)[] rows,
        string type = "string",
        bool styleCell = false
    ) => CreateWorkbookCore(rows, type, styleCell);

    private string CreateWorkbookCore(
        (string sheet, string id, string value)[] rows,
        string type,
        bool styleCell
    )
    {
        var path = Path.Combine(
            Path.GetTempPath(),
            "NumDesTools-semantic-" + Guid.NewGuid().ToString("N") + ".xlsx"
        );
        _temporaryPaths.Add(path);

        using var package = new ExcelPackage();
        foreach (var sheetRows in rows.GroupBy(row => row.sheet, StringComparer.Ordinal))
        {
            var worksheet = package.Workbook.Worksheets.Add(sheetRows.Key);
            worksheet.Cells[2, 1].Value = "#note";
            worksheet.Cells[2, 2].Value = "id";
            worksheet.Cells[2, 3].Value = "value";
            worksheet.Cells[3, 2].Value = type;
            worksheet.Cells[3, 3].Value = "string";
            worksheet.Cells[4, 2].Value = "Id";
            worksheet.Cells[4, 3].Value = "Value";

            var rowIndex = 5;
            foreach (var row in sheetRows)
            {
                worksheet.Cells[rowIndex, 2].Value = row.id;
                worksheet.Cells[rowIndex, 3].Value = row.value;
                rowIndex++;
            }

            if (styleCell)
                worksheet.Cells[5, 3].Style.Font.Bold = true;
        }

        package.SaveAs(new FileInfo(path));
        return path;
    }

    private static byte[] ReadBlob(Repository repository, IndexEntry entry)
    {
        using var stream = repository.Lookup<Blob>(entry.Id)!.GetContentStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void RunGit(string repositoryPath, string arguments, bool expectSuccess = true)
    {
        using var process = Process.Start(
            new ProcessStartInfo
            {
                FileName = "git",
                Arguments = arguments,
                WorkingDirectory = repositoryPath,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            }
        )!;
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (expectSuccess != (process.ExitCode == 0))
            throw new InvalidOperationException($"git {arguments} failed: {output}{error}");
    }

    public void Dispose()
    {
        foreach (var path in _temporaryPaths)
        {
            try
            {
                if (Directory.Exists(path))
                    Directory.Delete(path, recursive: true);
                else
                    File.Delete(path);
            }
            catch { }
        }
    }
}
