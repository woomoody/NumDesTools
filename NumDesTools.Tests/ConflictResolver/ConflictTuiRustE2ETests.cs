using System.Text.Json;
using System.Text.Json.Serialization;
using NumDesTools.ConflictResolver;
using OfficeOpenXml;

namespace NumDesTools.Tests.ConflictResolver;

/// <summary>端到端：C# Diff → JSON → Rust TUI → result.json → 合并 → Apply（跳过真实 TUI 交互，模拟 selections）。</summary>
public class ConflictTuiRustE2ETests : IDisposable
{
    static ConflictTuiRustE2ETests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    private readonly List<string> _tmpFiles = [];

    public void Dispose()
    {
        foreach (var f in _tmpFiles.Where(File.Exists))
            File.Delete(f);
    }

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

    [Fact]
    public void E2E_DiffToJson_ApplySelections_WritesCorrectValue()
    {
        // 1. 造冲突：ours 是"旧名称"，theirs 是"新名称"
        var oursPath = MakeXlsx(("1001", "旧名称"));
        var theirsPath = MakeXlsx(("1001", "新名称"));
        var outPath = TmpOut();

        // 2. Diff
        var diff = ExcelConflictDiffer.Diff(oursPath, theirsPath);
        Assert.Single(diff.Sheets[0].Rows, r => r.DiffType == RowDiffType.Modified);

        // 3. 序列化成 JSON（模拟 C# 端写 diff.json）
        var dto = diff.ToDto();
        var json = dto.ToJson();
        Assert.Contains("\"sheetName\":\"Sheet1\"", json);
        Assert.Contains("\"rowKey\":\"1001\"", json);

        // 4. 模拟 Rust TUI 返回的 selections（用户选了 Theirs）
        var result = new SelectionResultDto
        {
            Confirmed = true,
            Selections =
            [
                new SelectionDto
                {
                    SheetName = "Sheet1",
                    RowKey = "1001",
                    ColName = "name",
                    Choice = ConflictChoice.Theirs,
                },
            ],
        };

        // 5. 合并到 FileDiff
        diff.ApplySelections(result);
        var cell = diff.Sheets[0].Rows[0].Cells.First(c => c.ColName == "name");
        Assert.Equal(ConflictChoice.Theirs, cell.Choice);
        Assert.True(cell.IsExplicit);

        // 6. Apply 写回
        ConflictApplier.Apply(diff, outPath, gitAdd: false);

        // 7. 验证输出文件里是"新名称"
        using var pkg = new ExcelPackage(new FileInfo(outPath));
        var ws = pkg.Workbook.Worksheets["Sheet1"];
        Assert.Equal("新名称", ws.Cells[5, 3].Value?.ToString());
    }

    [Fact]
    public void E2E_JsonRoundtrip_PreservesAllFields()
    {
        var oursPath = MakeXlsx(("1001", "旧名称"));
        var theirsPath = MakeXlsx(("1001", "新名称"));
        var diff = ExcelConflictDiffer.Diff(oursPath, theirsPath);

        var dto = diff.ToDto();
        var json = dto.ToJson();
        var deserialized = FileDiffDto.FromJson(json);

        Assert.NotNull(deserialized);
        Assert.Equal(dto.Sheets.Count, deserialized.Sheets.Count);
        Assert.Equal(
            dto.Sheets[0].Rows[0].Cells[0].ColName,
            deserialized.Sheets[0].Rows[0].Cells[0].ColName
        );
        Assert.Equal(
            dto.Sheets[0].Rows[0].Cells[0].Choice,
            deserialized.Sheets[0].Rows[0].Cells[0].Choice
        );
    }

    [Fact]
    public void E2E_RustJsonFormat_DeserializesCorrectly()
    {
        // 模拟 Rust serde_json 输出的格式（camelCase + 枚举 string）
        var rustJson =
            @"{
            ""confirmed"": true,
            ""selections"": [
                { ""sheetName"": ""Sheet1"", ""rowKey"": ""1001"", ""colName"": ""name"", ""choice"": ""Theirs"" }
            ]
        }";
        var result = JsonSerializer.Deserialize<SelectionResultDto>(
            rustJson,
            new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                Converters = { new JsonStringEnumConverter() },
            }
        );

        Assert.NotNull(result);
        Assert.True(result.Confirmed);
        Assert.Single(result.Selections);
        Assert.Equal(ConflictChoice.Theirs, result.Selections[0].Choice);
    }
}
