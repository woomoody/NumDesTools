using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace NumDesTools.ConflictResolver;

public enum ConflictChoice
{
    Ours,
    Theirs,
}

public enum RowDiffType
{
    Modified,
    OnlyOurs,
    OnlyTheirs,
    Same,
}

/// <summary>
/// 三方推断的行来源（需要 merge-base 才能区分，无 base 时为 Unknown）。
/// </summary>
public enum RowOrigin
{
    Unknown,
    AddedByOurs, // 不在 base、只在 OURS：A 新增
    DeletedByTheirs, // 在 base、只在 OURS：B 删除
    AddedByTheirs, // 不在 base、只在 THEIRS：B 新增
    DeletedByOurs, // 在 base、只在 THEIRS：A 删除
}

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
        set
        {
            _choice = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ChoiceOurs));
            OnPropertyChanged(nameof(ChoiceTheirs));
        }
    }

    /// <summary>用户已明确做出选择（true=已选，false=待选）</summary>
    public bool IsExplicit
    {
        get => _isExplicit;
        set
        {
            _isExplicit = value;
            OnPropertyChanged();
        }
    }

    public void ClearChoice()
    {
        _isExplicit = false;
        OnPropertyChanged(nameof(IsExplicit));
    }

    public bool ChoiceOurs
    {
        get => Choice == ConflictChoice.Ours;
        set
        {
            if (value)
                Choice = ConflictChoice.Ours;
        }
    }
    public bool ChoiceTheirs
    {
        get => Choice == ConflictChoice.Theirs;
        set
        {
            if (value)
                Choice = ConflictChoice.Theirs;
        }
    }

    public string OursDisplay => OursValue?.ToString() ?? "(空)";
    public string TheirsDisplay => TheirsValue?.ToString() ?? "(空)";

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>一行内所有差异的聚合</summary>
public class RowConflict : INotifyPropertyChanged
{
    public RowConflict()
    {
        Cells.CollectionChanged += (_, e) =>
        {
            if (e.NewItems != null)
            {
                foreach (CellConflict cell in e.NewItems)
                    cell.PropertyChanged += OnCellPropertyChanged;
            }
        };
    }

    private void OnCellPropertyChanged(
        object? sender,
        System.ComponentModel.PropertyChangedEventArgs e
    )
    {
        // 冒泡：单元格选择变化时通知行刷新
        OnPropertyChanged(e.PropertyName);
    }

    public string SheetName { get; init; } = string.Empty;
    public string RowKey { get; init; } = string.Empty;
    public RowDiffType DiffType { get; init; }

    /// <summary>仅有差异的列（Modified时）</summary>
    public ObservableCollection<CellConflict> Cells { get; } = [];

    /// <summary>完整列顺序，用于渲染整行预览</summary>
    public List<string> AllColumns { get; init; } = [];
    public IDictionary<string, object?>? OursFullRow { get; init; }
    public IDictionary<string, object?>? TheirsFullRow { get; init; }

    /// <summary>所有 # 前缀列（备注类）的非空值，按列顺序，供行头显示</summary>
    public List<(string Col, string Val)> HashColValues
    {
        get
        {
            var src = OursFullRow ?? TheirsFullRow;
            if (src == null || AllColumns.Count == 0)
                return [];
            var result = new List<(string, string)>();
            foreach (var col in AllColumns)
            {
                if (!col.StartsWith('#'))
                    continue;
                if (!src.TryGetValue(col, out var v))
                    continue;
                var s = v?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(s))
                    result.Add((col, s));
            }
            return result;
        }
    }

    /// <summary>所有 # 列值拼成单行，供 ToolTip 等场景</summary>
    public string DisplayName => string.Join("  ", HashColValues.Select(x => x.Val));

    /// <summary>该行在 THEIRS 文件中的原始行索引（0-based data row）；-1 表示不来自 THEIRS。</summary>
    public int TheirsRowIndex { get; init; } = -1;

    /// <summary>该行在 OURS 文件中的原始行索引（0-based data row）；-1 表示 OURS 无此行。</summary>
    public int OursRowIndex { get; init; } = -1;

    /// <summary>三方推断的行来源（无 merge-base 时为 Unknown）。</summary>
    public RowOrigin Origin { get; init; } = RowOrigin.Unknown;

    public string DiffTypeBadge =>
        DiffType switch
        {
            RowDiffType.OnlyOurs => Origin switch
            {
                RowOrigin.AddedByOurs => "我方新增 ✓（已选保留）",
                RowOrigin.DeletedByTheirs => "对方删除 ✓（已选接受）",
                _ => "仅我有",
            },
            RowDiffType.OnlyTheirs => Origin switch
            {
                RowOrigin.AddedByTheirs => "对方新增 ✓（已选保留）",
                RowOrigin.DeletedByOurs => "我方删除 ✓（已选接受）",
                _ => "仅对方有",
            },
            RowDiffType.Same => "相同",
            _ => $"冲突 {Cells.Count} 列",
        };

    /// <summary>整行的选择（仅 OnlyOurs/OnlyTheirs 时有效）</summary>
    private ConflictChoice _rowChoice;

    /// <summary>
    /// 仅供 Differ 对象初始化器使用：设置默认选择但不标记"用户已明确处理"。
    /// UI 交互必须使用 <see cref="RowChoice"/>。
    /// </summary>
    public ConflictChoice DefaultRowChoice
    {
        init
        {
            _rowChoice = value;
            _rowChoiceExplicit = true; // OnlyOurs/OnlyTheirs 行默认即已解决，不须用户再点
        }
    }

    public ConflictChoice RowChoice
    {
        get => _rowChoice;
        set
        {
            _rowChoice = value;
            _rowChoiceExplicit = true;
            OnPropertyChanged();
            OnPropertyChanged(nameof(RowChoiceOurs));
            OnPropertyChanged(nameof(RowChoiceTheirs));
        }
    }
    public bool RowChoiceOurs
    {
        get => RowChoice == ConflictChoice.Ours;
        set
        {
            if (value)
                RowChoice = ConflictChoice.Ours;
        }
    }
    public bool RowChoiceTheirs
    {
        get => RowChoice == ConflictChoice.Theirs;
        set
        {
            if (value)
                RowChoice = ConflictChoice.Theirs;
        }
    }

    /// <summary>用户是否已明确处理过该行（Modified = 所有冲突格都做过选择；OnlyOurs/OnlyTheirs = 用户主动选过保留/放弃）</summary>
    public bool IsResolved =>
        DiffType switch
        {
            RowDiffType.Modified => Cells.All(c => c.IsExplicit),
            RowDiffType.OnlyOurs or RowDiffType.OnlyTheirs => _rowChoiceExplicit,
            _ => true,
        };

    private bool _rowChoiceExplicit;

    private string _aiSuggestion = string.Empty;
    public string AiSuggestion
    {
        get => _aiSuggestion;
        set
        {
            _aiSuggestion = value;
            OnPropertyChanged();
        }
    }

    /// <summary>拖拽多选时的行选中状态（纯视觉，不持久化）</summary>
    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value)
                return;
            _isSelected = value;
            OnPropertyChanged();
        }
    }

    /// <summary>将该行所有单元格冲突批量设置为同一选择</summary>
    public void SetAllCells(ConflictChoice choice)
    {
        foreach (var c in Cells)
        {
            c.Choice = choice;
            c.IsExplicit = true;
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>一个 Sheet 的所有行差异</summary>
public record SheetDiff(string SheetName, List<RowConflict> Rows)
{
    public bool HasConflict => Rows.Any(r => r.DiffType != RowDiffType.Same);

    /// <summary>存在"双方都改、选不出默认值"的行——需要人工判断，不是靠三方预选/新增删除默认就能了事的真冲突。</summary>
    public bool HasTrueConflict => Rows.Any(r => r.DiffType != RowDiffType.Same && !r.IsResolved);

    /// <summary>列名 → type 行值（第3行，如 "int", "string"）</summary>
    public Dictionary<string, string> TypeRow { get; init; } = new();

    /// <summary>列名 → 中文说明行值（第4行）</summary>
    public Dictionary<string, string> LabelRow { get; init; } = new();

    /// <summary>列顺序（与 RowConflict.AllColumns 一致）</summary>
    public List<string> AllColumns { get; init; } = new();

    /// <summary>对所有行中指定列的冲突单元格批量设置选择</summary>
    public void SetColumnChoice(string colName, ConflictChoice choice)
    {
        foreach (var row in Rows)
        {
            if (row.DiffType == RowDiffType.Modified)
            {
                var cell = row.Cells.FirstOrDefault(c => c.ColName == colName);
                if (cell != null)
                {
                    cell.Choice = choice;
                    cell.IsExplicit = true;
                }
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
    public int TotalConflictRows =>
        Sheets.Sum(s => s.Rows.Count(r => r.DiffType != RowDiffType.Same));
    public int TotalConflictCells => Sheets.Sum(s => s.Rows.Sum(r => r.Cells.Count));

    /// <summary>整个文件里有没有需要人工判断的行。false=所有差异都已被三方预选/新增删除默认值覆盖，可以一键接受。</summary>
    public bool HasTrueConflict => Sheets.Any(s => s.HasTrueConflict);
}
