using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace NumDesTools.ConflictResolver;

public enum ConflictChoice { Ours, Theirs }

public enum RowDiffType { Modified, OnlyOurs, OnlyTheirs, Same }

/// <summary>单元格级别的冲突</summary>
public class CellConflict : INotifyPropertyChanged
{
    public string ColName { get; init; } = string.Empty;
    public object? OursValue { get; init; }
    public object? TheirsValue { get; init; }

    private ConflictChoice _choice = ConflictChoice.Ours;
    private bool _isExplicit = false;

    public ConflictChoice Choice
    {
        get => _choice;
        set { _choice = value; OnPropertyChanged(); OnPropertyChanged(nameof(ChoiceOurs)); OnPropertyChanged(nameof(ChoiceTheirs)); }
    }

    /// <summary>用户已明确做出选择（true=已选，false=待选）</summary>
    public bool IsExplicit
    {
        get => _isExplicit;
        set { _isExplicit = value; OnPropertyChanged(); }
    }

    public void ClearChoice()
    {
        _isExplicit = false;
        OnPropertyChanged(nameof(IsExplicit));
    }

    public bool ChoiceOurs  { get => Choice == ConflictChoice.Ours;   set { if (value) Choice = ConflictChoice.Ours; } }
    public bool ChoiceTheirs{ get => Choice == ConflictChoice.Theirs; set { if (value) Choice = ConflictChoice.Theirs; } }

    public string OursDisplay   => OursValue?.ToString()   ?? "(空)";
    public string TheirsDisplay => TheirsValue?.ToString() ?? "(空)";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>一行内所有差异的聚合</summary>
public class RowConflict : INotifyPropertyChanged
{
    public string SheetName { get; init; } = string.Empty;
    public string RowKey    { get; init; } = string.Empty;
    public RowDiffType DiffType { get; init; }

    /// <summary>仅有差异的列（Modified时）</summary>
    public ObservableCollection<CellConflict> Cells { get; } = [];

    /// <summary>完整列顺序，用于渲染整行预览</summary>
    public List<string> AllColumns { get; init; } = [];
    public IDictionary<string, object?>? OursFullRow   { get; init; }
    public IDictionary<string, object?>? TheirsFullRow { get; init; }

    public string DiffTypeBadge => DiffType switch
    {
        RowDiffType.OnlyOurs   => "仅我有（对方删除）",
        RowDiffType.OnlyTheirs => "仅对方有（新增行）",
        RowDiffType.Same       => "相同",
        _                      => $"冲突 {Cells.Count} 列"
    };

    /// <summary>整行的选择（仅 OnlyOurs/OnlyTheirs 时有效）</summary>
    private ConflictChoice _rowChoice;
    public ConflictChoice RowChoice
    {
        get => _rowChoice;
        set { _rowChoice = value; OnPropertyChanged(); OnPropertyChanged(nameof(RowChoiceOurs)); OnPropertyChanged(nameof(RowChoiceTheirs)); }
    }
    public bool RowChoiceOurs   { get => RowChoice == ConflictChoice.Ours;   set { if (value) RowChoice = ConflictChoice.Ours; } }
    public bool RowChoiceTheirs { get => RowChoice == ConflictChoice.Theirs; set { if (value) RowChoice = ConflictChoice.Theirs; } }

    public bool IsResolved => DiffType switch
    {
        RowDiffType.Modified => true, // 格级选择总是有值（Ours/Theirs），不存在"未选"
        _                    => true
    };

    /// <summary>将该行所有单元格冲突批量设置为同一选择</summary>
    public void SetAllCells(ConflictChoice choice)
    {
        foreach (var c in Cells) { c.Choice = choice; c.IsExplicit = true; }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>一个 Sheet 的所有行差异</summary>
public record SheetDiff(string SheetName, List<RowConflict> Rows)
{
    public bool HasConflict => Rows.Any(r => r.DiffType != RowDiffType.Same);

    /// <summary>列名 → type 行值（第3行，如 "int", "string"）</summary>
    public Dictionary<string, string> TypeRow    { get; init; } = new();
    /// <summary>列名 → 中文说明行值（第4行）</summary>
    public Dictionary<string, string> LabelRow   { get; init; } = new();
    /// <summary>列顺序（与 RowConflict.AllColumns 一致）</summary>
    public List<string>               AllColumns { get; init; } = new();

    /// <summary>对所有行中指定列的冲突单元格批量设置选择</summary>
    public void SetColumnChoice(string colName, ConflictChoice choice)
    {
        foreach (var row in Rows)
        {
            if (row.DiffType == RowDiffType.Modified)
            {
                var cell = row.Cells.FirstOrDefault(c => c.ColName == colName);
                if (cell != null) { cell.Choice = choice; cell.IsExplicit = true; }
            }
            else
            {
                // OnlyOurs / OnlyTheirs：整行选择
                row.RowChoice = choice;
            }
        }
    }
}

/// <summary>整个文件的差异结果</summary>
public record FileDiff(string OursPath, string TheirsPath, List<SheetDiff> Sheets)
{
    public int TotalConflictRows  => Sheets.Sum(s => s.Rows.Count(r => r.DiffType != RowDiffType.Same));
    public int TotalConflictCells => Sheets.Sum(s => s.Rows.Sum(r => r.Cells.Count));
}
