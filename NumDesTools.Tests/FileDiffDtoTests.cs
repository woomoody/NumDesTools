using NumDesTools.ConflictResolver;

namespace NumDesTools.Tests;

/// <summary>FileDiffDto 序列化/反序列化对齐测试（C# ↔ Rust schema 一致性）。</summary>
public class FileDiffDtoTests
{
    [Fact]
    public void FileDiffDto_Roundtrip_PreservesData()
    {
        var dto = new FileDiffDto
        {
            OursPath = @"C:\a.xlsx",
            TheirsPath = @"C:\b.xlsx",
            Sheets =
            [
                new SheetDiffDto
                {
                    SheetName = "Sheet1",
                    AllColumns = ["id", "hp"],
                    TypeRow = new() { ["id"] = "int" },
                    LabelRow = new() { ["id"] = "编号" },
                    Rows =
                    [
                        new RowConflictDto
                        {
                            SheetName = "Sheet1",
                            RowKey = "1001",
                            DiffType = RowDiffType.Modified,
                            Origin = RowOrigin.Unknown,
                            OursRowIndex = 5,
                            TheirsRowIndex = 5,
                            AllColumns = ["id", "hp"],
                            OursFullRow = new() { ["id"] = "1001", ["hp"] = "100" },
                            TheirsFullRow = new() { ["id"] = "1001", ["hp"] = "120" },
                            Cells =
                            [
                                new CellConflictDto
                                {
                                    ColName = "hp",
                                    OursValue = "100",
                                    TheirsValue = "120",
                                    Choice = ConflictChoice.Ours,
                                    IsExplicit = false,
                                },
                            ],
                            RowChoice = ConflictChoice.Ours,
                            RowChoiceExplicit = false,
                            AiSuggestion = "",
                        },
                    ],
                },
            ],
        };

        var json = dto.ToJson();
        var deserialized = FileDiffDto.FromJson(json);

        Assert.NotNull(deserialized);
        Assert.Equal(dto.OursPath, deserialized.OursPath);
        Assert.Single(deserialized.Sheets);
        Assert.Equal("hp", deserialized.Sheets[0].Rows[0].Cells[0].ColName);
        Assert.Equal(ConflictChoice.Ours, deserialized.Sheets[0].Rows[0].Cells[0].Choice);
    }

    [Fact]
    public void FileDiffDto_JsonUsesCamelCase()
    {
        var dto = new FileDiffDto { OursPath = "a", TheirsPath = "b" };
        var json = dto.ToJson();
        Assert.Contains("\"oursPath\"", json);
        Assert.Contains("\"theirsPath\"", json);
        Assert.DoesNotContain("\"OursPath\"", json);
    }

    [Fact]
    public void SelectionResultDto_DeserializesFromRustFormat()
    {
        // Rust 端 serde_json 输出的格式（camelCase + 枚举按 string）
        var json =
            @"{
            ""confirmed"": true,
            ""selections"": [
                { ""sheetName"": ""Sheet1"", ""rowKey"": ""1001"", ""colName"": ""hp"", ""choice"": ""Theirs"" },
                { ""sheetName"": ""Sheet1"", ""rowKey"": ""1002"", ""colName"": null, ""choice"": ""Ours"" }
            ]
        }";
        var result = System.Text.Json.JsonSerializer.Deserialize<SelectionResultDto>(
            json,
            new System.Text.Json.JsonSerializerOptions
            {
                PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
                Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() },
            }
        );

        Assert.NotNull(result);
        Assert.True(result.Confirmed);
        Assert.Equal(2, result.Selections.Count);
        Assert.Equal(ConflictChoice.Theirs, result.Selections[0].Choice);
        Assert.Null(result.Selections[1].ColName);
    }

    [Fact]
    public void ApplySelections_UpdatesChoiceAndIsExplicit()
    {
        var diff = new FileDiff(
            "a.xlsx",
            "b.xlsx",
            [
                new SheetDiff(
                    "Sheet1",
                    [
                        new RowConflict
                        {
                            SheetName = "Sheet1",
                            RowKey = "1001",
                            DiffType = RowDiffType.Modified,
                            Cells = { new CellConflict { ColName = "hp", OursValue = "100", TheirsValue = "120" } },
                        },
                    ]
                ),
            ]
        );
        var result = new SelectionResultDto
        {
            Confirmed = true,
            Selections =
            [
                new SelectionDto
                {
                    SheetName = "Sheet1",
                    RowKey = "1001",
                    ColName = "hp",
                    Choice = ConflictChoice.Theirs,
                },
            ],
        };

        diff.ApplySelections(result);

        var cell = diff.Sheets[0].Rows[0].Cells[0];
        Assert.Equal(ConflictChoice.Theirs, cell.Choice);
        Assert.True(cell.IsExplicit);
    }
}
