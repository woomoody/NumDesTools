using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

// ── 数据行模型 ────────────────────────────────────────────────────────────────

public class CloneIdRow : INotifyPropertyChanged
{
    private string _sourceId  = "";
    private string _targetId  = "";
    private string _remark    = "";
    private string _boundTable = "";
    private System.Windows.Media.Color  _rowBg     = System.Windows.Media.Color.FromRgb(0x23, 0x23, 0x23);

    public string SourceId   { get => _sourceId;   set { _sourceId   = value; OnChanged(); } }
    public string TargetId   { get => _targetId;   set { _targetId   = value; OnChanged(); } }
    public string Remark     { get => _remark;     set { _remark     = value; OnChanged(); } }
    public string BoundTable { get => _boundTable; set { _boundTable = value; OnChanged(); } }
    public System.Windows.Media.Color  RowBg      { get => _rowBg;      set { _rowBg      = value; OnChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([System.Runtime.CompilerServices.CallerMemberName] string? p = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
}

// ── 疑似表决策模型 ────────────────────────────────────────────────────────────

public class SuspectDecision : INotifyPropertyChanged
{
    private string _tableName = "";
    private bool   _include   = true;

    public string TableName { get => _tableName; set { _tableName = value; OnChanged(); } }
    public bool   Include   { get => _include;   set { _include   = value; OnChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([System.Runtime.CompilerServices.CallerMemberName] string? p = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
}

// ── 目标表选择模型 ────────────────────────────────────────────────────────────

public class TableSelection : INotifyPropertyChanged
{
    private string _tableName = "";
    private bool   _selected  = true;

    public string TableName { get => _tableName; set { _tableName = value; OnChanged(); } }
    public bool   Selected  { get => _selected;  set { _selected  = value; OnChanged(); } }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnChanged([System.Runtime.CompilerServices.CallerMemberName] string? p = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));
}

// ── 历史条目模型 ──────────────────────────────────────────────────────────────

public class HistoryEntry
{
    public string Label   { get; set; } = "";
    public string Tooltip { get; set; } = "";
    public List<CloneIdRow> Rows          { get; set; } = new();
    public string           Remark        { get; set; } = "";
    public bool?            ReplaceArt    { get; set; } = null; // null=未记录
    public bool?            ReplaceSubTable { get; set; } = null;
    public List<SuspectDecision> SuspectDecisions { get; set; } = new();
    public List<TableSelection>  TableSelections  { get; set; } = new();
}

// ── 窗口 ──────────────────────────────────────────────────────────────────────

public partial class CloneActivityWindow : Window
{
    public bool Confirmed { get; private set; }
    public List<CloneIdRow>       ResultRows            { get; private set; } = new();
    public string                 ResultRemark          { get; private set; } = "";
    public bool?                  ResultReplaceArt      { get; private set; } = null;
    public bool?                  ResultReplaceSubTable { get; private set; } = null;
    public List<SuspectDecision>  ResultSuspectDecisions { get; private set; } = new();

    private readonly ObservableCollection<CloneIdRow>       _rows             = new();
    private readonly ObservableCollection<HistoryEntry>     _history          = new();
    private readonly ObservableCollection<SuspectDecision>  _suspectDecisions = new();

    private readonly List<CloneIdRow>? _prefillRows;

    public CloneActivityWindow(List<CloneIdRow>? prefillRows = null)
    {
        InitializeComponent();
        _prefillRows = prefillRows;
        IdRowList.ItemsSource       = _rows;
        HistoryList.ItemsSource     = _history;
        SuspectList.ItemsSource     = _suspectDecisions;
        LoadHistory();
        InitRows();
    }

    // ── 初始化行 ──────────────────────────────────────────────────────────────

    private void InitRows()
    {
        if (_prefillRows?.Count > 0)
        {
            foreach (var r in _prefillRows)
            {
                if (!string.IsNullOrEmpty(r.BoundTable))
                    r.RowBg = System.Windows.Media.Color.FromRgb(0x3A, 0x2A, 0x10);
                _rows.Add(r);
            }
        }
        else
        {
            _rows.Add(new CloneIdRow());
        }
    }

    // ── 外部注入疑似表列表（克隆前由 ActivityDataCloner 调用）────────────────

    public void SetSuspectTables(IEnumerable<string> tableNames,
                                  IEnumerable<SuspectDecision>? saved = null)
    {
        _suspectDecisions.Clear();
        var savedMap = saved?.ToDictionary(d => d.TableName, d => d.Include,
                           StringComparer.OrdinalIgnoreCase)
                       ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        foreach (var t in tableNames)
        {
            _suspectDecisions.Add(new SuspectDecision
            {
                TableName = t,
                Include   = savedMap.TryGetValue(t, out var v) ? v : true,
            });
        }
        SuspectPanel.Visibility = _suspectDecisions.Count > 0
            ? Visibility.Visible : Visibility.Collapsed;
    }

    // ── 历史记录 I/O ──────────────────────────────────────────────────────────

    private static string HistoryFile =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                     "NumDesTools", "Config", "clone_history_ui.json");

    private void LoadHistory()
    {
        _history.Clear();
        if (!File.Exists(HistoryFile)) return;
        try
        {
            var list = Newtonsoft.Json.JsonConvert.DeserializeObject<List<HistoryEntry>>(
                File.ReadAllText(HistoryFile)) ?? new();
            foreach (var e in list) _history.Add(e);
        }
        catch { }
    }

    private void SaveHistory()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(HistoryFile)!);
            File.WriteAllText(HistoryFile,
                Newtonsoft.Json.JsonConvert.SerializeObject(_history.ToList(),
                    Newtonsoft.Json.Formatting.Indented),
                System.Text.Encoding.UTF8);
        }
        catch { }
    }

    private void PushCurrentToHistory(List<TableSelection>? tableSelections = null)
    {
        var validRows = _rows.Where(r => !string.IsNullOrWhiteSpace(r.SourceId)
                                      && !string.IsNullOrWhiteSpace(r.TargetId)).ToList();
        if (validRows.Count == 0) return;

        var pairs = validRows
            .Where(r => string.IsNullOrEmpty(r.BoundTable))
            .Select(r => $"{r.SourceId}→{r.TargetId}");
        var label = string.Join("  ", pairs);
        if (!string.IsNullOrEmpty(RemarkBox.Text))
            label += "  " + RemarkBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(label)) label = validRows[0].SourceId + "→" + validRows[0].TargetId;

        var tooltip = string.Join("\n", validRows.Select(r =>
        {
            var s = $"{r.SourceId} → {r.TargetId}";
            if (!string.IsNullOrEmpty(r.Remark))     s += $"  ({r.Remark})";
            if (!string.IsNullOrEmpty(r.BoundTable)) s += $"  [表:{r.BoundTable}]";
            return s;
        }));
        tooltip += $"\n美术资源：{(ReplaceArtCheck.IsChecked == true ? "替换" : "保留")}";
        tooltip += $"\n子表ID：{(ReplaceSubCheck.IsChecked == true ? "替换" : "保留")}";
        if (tableSelections?.Count > 0)
        {
            var skipped = tableSelections.Where(t => !t.Selected).Select(t => t.TableName).ToList();
            if (skipped.Count > 0)
                tooltip += $"\n跳过表：{string.Join("  ", skipped)}";
        }

        var existing = _history.FirstOrDefault(h => h.Label == label);
        if (existing != null) _history.Remove(existing);

        var entry = new HistoryEntry
        {
            Label   = label,
            Tooltip = tooltip,
            Rows    = validRows.Select(r => new CloneIdRow
            {
                SourceId   = r.SourceId,
                TargetId   = r.TargetId,
                Remark     = r.Remark,
                BoundTable = r.BoundTable,
            }).ToList(),
            Remark           = RemarkBox.Text.Trim(),
            ReplaceArt       = ReplaceArtCheck.IsChecked,
            ReplaceSubTable  = ReplaceSubCheck.IsChecked,
            SuspectDecisions = _suspectDecisions.Select(d => new SuspectDecision
            {
                TableName = d.TableName,
                Include   = d.Include,
            }).ToList(),
            TableSelections = tableSelections?.Select(t => new TableSelection
            {
                TableName = t.TableName,
                Selected  = t.Selected,
            }).ToList() ?? new(),
        };
        _history.Insert(0, entry);
        if (_history.Count > 20) _history.RemoveAt(_history.Count - 1);
        SaveHistory();
    }

    // ── UI 事件 ───────────────────────────────────────────────────────────────

    private void AddRow_Click(object sender, RoutedEventArgs e)
        => _rows.Add(new CloneIdRow());

    private void RemoveRow_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is CloneIdRow row && _rows.Count > 1)
            _rows.Remove(row);
    }

    private void HistoryList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        // 单击只预览提示，不立即应用
    }

    private void ApplyHistory_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryList.SelectedItem is not HistoryEntry entry) return;
        ApplyHistoryEntry(entry);
    }

    private void ApplyHistoryEntry(HistoryEntry entry)
    {
        _rows.Clear();
        foreach (var r in entry.Rows)
        {
            var row = new CloneIdRow
            {
                SourceId   = r.SourceId,
                TargetId   = r.TargetId,
                Remark     = r.Remark,
                BoundTable = r.BoundTable,
            };
            if (!string.IsNullOrEmpty(r.BoundTable))
                row.RowBg = System.Windows.Media.Color.FromRgb(0x3A, 0x2A, 0x10);
            _rows.Add(row);
        }
        RemarkBox.Text = entry.Remark;

        // 恢复决策状态（有历史值则填充，null 时用默认值：美术=替换，子表=保留）
        ReplaceArtCheck.IsChecked = entry.ReplaceArt     ?? true;
        ReplaceSubCheck.IsChecked = entry.ReplaceSubTable ?? false;

        _suspectDecisions.Clear();
        foreach (var d in entry.SuspectDecisions)
            _suspectDecisions.Add(new SuspectDecision { TableName = d.TableName, Include = d.Include });
        SuspectPanel.Visibility = _suspectDecisions.Count > 0
            ? Visibility.Visible : Visibility.Collapsed;

        SavedTableSelections = entry.TableSelections.Select(t => new TableSelection
        {
            TableName = t.TableName,
            Selected  = t.Selected,
        }).ToList();

        StatusText.Text = $"已应用历史：{entry.Label}";
    }

    private void DeleteHistory_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryList.SelectedItem is not HistoryEntry entry) return;
        _history.Remove(entry);
        SaveHistory();
    }

    private void Run_Click(object sender, RoutedEventArgs e)
    {
        var valid = _rows.Where(r => !string.IsNullOrWhiteSpace(r.SourceId)
                                   && !string.IsNullOrWhiteSpace(r.TargetId)).ToList();
        if (valid.Count == 0)
        {
            StatusText.Text = "请至少填写一对源 ID 和目标 ID。";
            return;
        }

        foreach (var r in valid)
        {
            if (!long.TryParse(r.SourceId.Trim(), out _) && string.IsNullOrEmpty(r.BoundTable))
            {
                StatusText.Text = $"源 ID「{r.SourceId}」格式有误，请填纯数字。";
                return;
            }
            if (!long.TryParse(r.TargetId.Trim(), out _) && string.IsNullOrEmpty(r.BoundTable))
            {
                StatusText.Text = $"目标 ID「{r.TargetId}」格式有误，请填纯数字。";
                return;
            }
        }

        // 历史先不带表选择，RunInternal 完成后再追加 table selections
        PushCurrentToHistory(null);
        ResultRows             = valid;
        ResultRemark           = RemarkBox.Text.Trim();
        ResultReplaceArt       = ReplaceArtCheck.IsChecked;
        ResultReplaceSubTable  = ReplaceSubCheck.IsChecked;
        ResultSuspectDecisions = _suspectDecisions.ToList();
        Confirmed = true;
        Close();
    }

    /// <summary>
    /// RunInternal 完成表列表发现后回写，把 tableSelections 补存到最近一条历史记录里。
    /// 历史已由 Run_Click 写入，这里只更新最新一条的 TableSelections 字段。
    /// </summary>
    public void UpdateHistoryWithTableSelections(List<TableSelection> tableSelections)
    {
        if (_history.Count == 0) return;
        var latest = _history[0];
        latest.TableSelections = tableSelections.Select(t => new TableSelection
        {
            TableName = t.TableName,
            Selected  = t.Selected,
        }).ToList();

        // 更新 tooltip 里的跳过列表
        var skipped = tableSelections.Where(t => !t.Selected).Select(t => t.TableName).ToList();
        var tooltip = latest.Tooltip;
        // 移除旧的跳过行（如果有）
        var skipLine = tooltip.Split('\n').FirstOrDefault(l => l.StartsWith("跳过表："));
        if (skipLine != null) tooltip = tooltip.Replace("\n" + skipLine, "");
        if (skipped.Count > 0)
            tooltip += $"\n跳过表：{string.Join("  ", skipped)}";
        latest.Tooltip = tooltip;

        SaveHistory();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    // ── 外部调用辅助 ──────────────────────────────────────────────────────────

    /// <summary>Returns the saved table selections from the last applied history entry (may be empty).</summary>
    public List<TableSelection> SavedTableSelections { get; private set; } = new();

    public (List<(string src, string dst)> global,
            List<(string src, string dst, string table)> perTable) ParseResult()
    {
        var global   = new List<(string, string)>();
        var perTable = new List<(string, string, string)>();
        foreach (var r in ResultRows)
        {
            var src = r.SourceId.Trim();
            var dst = r.TargetId.Trim();
            if (string.IsNullOrEmpty(r.BoundTable))
                global.Add((src, dst));
            else
                perTable.Add((src, dst, r.BoundTable.Trim()));
        }
        return (global, perTable);
    }
}
