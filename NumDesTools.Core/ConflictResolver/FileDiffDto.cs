using System.Text.Json;
using System.Text.Json.Serialization;

namespace NumDesTools.ConflictResolver;

/// <summary>
/// FileDiff 的 JSON 序列化 DTO（纯 POCO，不碰 INotifyPropertyChanged/ObservableCollection）。
/// 用于 C# 引擎 → Rust TUI 的跨语言数据交换（文件系统中转方案）。
/// 序列化配置：camelCase + JsonStringEnumConverter + UTF8 无 BOM。
/// </summary>
public record FileDiffDto
{
    public required string OursPath { get; init; }
    public required string TheirsPath { get; init; }

    /// <summary>我方/对方对应的分支或 commit 名（git-add 场景下 OursPath/TheirsPath 只是临时提取路径，没有可读性）。</summary>
    public string? OursLabel { get; init; }
    public string? TheirsLabel { get; init; }
    public List<SheetDiffDto> Sheets { get; init; } = [];

    private static readonly JsonSerializerOptions s_jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() },
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public string ToJson() => JsonSerializer.Serialize(this, s_jsonOptions);

    public static FileDiffDto? FromJson(string json) =>
        JsonSerializer.Deserialize<FileDiffDto>(json, s_jsonOptions);
}

public record SheetDiffDto
{
    public required string SheetName { get; init; }
    public List<string> AllColumns { get; init; } = [];
    public Dictionary<string, string> TypeRow { get; init; } = new();
    public Dictionary<string, string> LabelRow { get; init; } = new();
    public List<RowConflictDto> Rows { get; init; } = [];
}

public record RowConflictDto
{
    public required string SheetName { get; init; }
    public required string RowKey { get; init; }
    public RowDiffType DiffType { get; init; }
    public RowOrigin Origin { get; init; }
    public int OursRowIndex { get; init; }
    public int TheirsRowIndex { get; init; }
    public List<string> AllColumns { get; init; } = [];
    public Dictionary<string, string?>? OursFullRow { get; init; }
    public Dictionary<string, string?>? TheirsFullRow { get; init; }
    public List<CellConflictDto> Cells { get; init; } = [];
    public ConflictChoice RowChoice { get; init; }
    public bool RowChoiceExplicit { get; init; }
    public string AiSuggestion { get; init; } = string.Empty;
}

public record CellConflictDto
{
    public required string ColName { get; init; }
    public string? OursValue { get; init; }
    public string? TheirsValue { get; init; }
    public ConflictChoice Choice { get; init; }
    public bool IsExplicit { get; init; }
}

/// <summary>Rust TUI 回传的用户选择（只传变化量，不回传完整 FileDiff）。</summary>
public record SelectionResultDto
{
    public bool Confirmed { get; init; }
    public List<SelectionDto> Selections { get; init; } = [];
}

public record SelectionDto
{
    public required string SheetName { get; init; }
    public required string RowKey { get; init; }
    public string? ColName { get; init; } // null = 整行（OnlyOurs/OnlyTheirs）
    public ConflictChoice Choice { get; init; }
}
