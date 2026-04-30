using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;
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
    public string Label        { get; set; } = "";
    public string Tooltip      { get; set; } = "";
    public string GroupId      { get; set; } = "";   // 活动组名，如"美梦玩法-迷梦甜品局"
    public string InstanceName { get; set; } = "";   // 期次名，如"第2期"
    public List<CloneIdRow> Rows          { get; set; } = new();
    public string           Remark        { get; set; } = "";
    public string           RemarkDst     { get; set; } = ""; // 备注替换目标文本
    public bool?            ReplaceArt    { get; set; } = null; // null=未记录
    public bool?            ReplaceSubTable { get; set; } = null;
    public bool?            IncrementRemark { get; set; } = null; // null=未记录，默认true
    public List<SuspectDecision>             SuspectDecisions     { get; set; } = new();
    public List<TableSelection>              TableSelections      { get; set; } = new();
    // key=表名, value=该表被屏蔽的异构前缀列表
    public Dictionary<string, List<string>> BlockedAlienPrefixes { get; set; } = new();
}

// ── 窗口 ──────────────────────────────────────────────────────────────────────

public partial class CloneActivityWindow : Window
{
    public bool Confirmed { get; private set; }
    public List<CloneIdRow>       ResultRows             { get; private set; } = new();
    public string                 ResultRemark           { get; private set; } = "";
    public string                 ResultRemarkDst        { get; private set; } = "";
    public bool?                  ResultReplaceArt       { get; private set; } = null;
    public bool?                  ResultReplaceSubTable  { get; private set; } = null;
    public bool                   ResultIncrementRemark  { get; private set; } = true;
    public bool                   ResultBlockAlienTable   { get; private set; } = false;
    public List<string>           ResultBlockAlienPrefixes { get; private set; } = new();
    public List<SuspectDecision>  ResultSuspectDecisions  { get; private set; } = new();

    private readonly ObservableCollection<CloneIdRow>       _rows             = new();
    private readonly ObservableCollection<HistoryEntry>     _history          = new();
    private readonly ObservableCollection<SuspectDecision>  _suspectDecisions = new();

    private readonly List<CloneIdRow>? _prefillRows;
    private readonly string            _excelPath;

    private readonly CollectionViewSource _historyView = new();

    public CloneActivityWindow(List<CloneIdRow>? prefillRows = null, string excelPath = "")
    {
        InitializeComponent();
        _prefillRows = prefillRows;
        _excelPath   = excelPath;
        IdRowList.ItemsSource   = _rows;
        SuspectList.ItemsSource = _suspectDecisions;

        _historyView.Source = _history;
        _historyView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(HistoryEntry.GroupId)));
        HistoryList.ItemsSource = _historyView.View;

        LoadHistory();
        InitRows();

        // 异构ID映射模式：预填行全部绑定了具体表 → 显示屏蔽按钮
        if (_prefillRows?.Count > 0 && _prefillRows.All(r => !string.IsNullOrEmpty(r.BoundTable)))
            BlockAlienBtn.Visibility = System.Windows.Visibility.Visible;
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

    private string HistoryFile =>
        Path.Combine(Path.GetDirectoryName(_excelPath)!, "TablesTools", "AliceConfig", "clone_history_ui.json");

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

    private void PushCurrentToHistory(List<TableSelection>? tableSelections = null,
                                       string groupId = "", string instanceName = "")
    {
        var validRows = _rows.Where(r => !string.IsNullOrWhiteSpace(r.SourceId)
                                      && !string.IsNullOrWhiteSpace(r.TargetId)).ToList();
        if (validRows.Count == 0) return;

        var pairs = validRows
            .Where(r => string.IsNullOrEmpty(r.BoundTable))
            .Select(r => $"{r.SourceId}→{r.TargetId}");
        var label = string.Join("  ", pairs);
        if (!string.IsNullOrEmpty(instanceName))
            label = string.IsNullOrEmpty(groupId) ? instanceName : $"{groupId} · {instanceName}";
        else if (!string.IsNullOrEmpty(groupId))
            label = groupId + "  " + label;
        else if (!string.IsNullOrEmpty(RemarkBox.Text))
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
        tooltip += $"\n备注序号自增：{(IncrementRemarkCheck.IsChecked == true ? "是" : "否")}";
        if (!string.IsNullOrEmpty(RemarkDstBox.Text.Trim()))
            tooltip += $"\n备注替换：{RemarkBox.Text.Trim()}→{RemarkDstBox.Text.Trim()}";
        if (tableSelections?.Count > 0)
        {
            var skipped = tableSelections.Where(t => !t.Selected).Select(t => t.TableName).ToList();
            if (skipped.Count > 0)
                tooltip += $"\n跳过表：{string.Join("  ", skipped)}";
        }

        // 已存在同 label → 原地更新所有字段，不移位不重写文件（等 UpdateHistoryWithTableSelections 时再保存）
        var existing = _history.FirstOrDefault(h => h.Label == label);
        if (existing != null)
        {
            existing.Tooltip         = tooltip;
            existing.Remark          = RemarkBox.Text.Trim();
            existing.RemarkDst       = RemarkDstBox.Text.Trim();
            existing.ReplaceArt      = ReplaceArtCheck.IsChecked;
            existing.ReplaceSubTable = ReplaceSubCheck.IsChecked;
            existing.IncrementRemark = IncrementRemarkCheck.IsChecked;
            existing.Rows = validRows.Select(r => new CloneIdRow
            {
                SourceId   = r.SourceId,
                TargetId   = r.TargetId,
                Remark     = r.Remark,
                BoundTable = r.BoundTable,
            }).ToList();
            return;
        }

        var entry = new HistoryEntry
        {
            Label        = label,
            Tooltip      = tooltip,
            GroupId      = groupId,
            InstanceName = instanceName,
            Rows    = validRows.Select(r => new CloneIdRow
            {
                SourceId   = r.SourceId,
                TargetId   = r.TargetId,
                Remark     = r.Remark,
                BoundTable = r.BoundTable,
            }).ToList(),
            Remark           = RemarkBox.Text.Trim(),
            RemarkDst        = RemarkDstBox.Text.Trim(),
            ReplaceArt       = ReplaceArtCheck.IsChecked,
            ReplaceSubTable  = ReplaceSubCheck.IsChecked,
            IncrementRemark  = IncrementRemarkCheck.IsChecked,
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

    private void SrcIncrement_Click(object sender, RoutedEventArgs e)
        => ShiftIds(srcDelta: +1, dstDelta: 0);

    private void SrcDecrement_Click(object sender, RoutedEventArgs e)
        => ShiftIds(srcDelta: -1, dstDelta: 0);

    private void DstIncrement_Click(object sender, RoutedEventArgs e)
        => ShiftIds(srcDelta: 0, dstDelta: +1);

    private void DstDecrement_Click(object sender, RoutedEventArgs e)
        => ShiftIds(srcDelta: 0, dstDelta: -1);

    // sign: +1 = 自增一步, -1 = 自减一步（步长 = |dst - src|）
    private void ShiftIds(int srcDelta, int dstDelta)
    {
        var valid = _rows.Where(r => long.TryParse(r.SourceId.Trim(), out _)
                                    && long.TryParse(r.TargetId.Trim(), out _)).ToList();
        if (valid.Count == 0) { StatusText.Text = "没有可操作的数字行。"; return; }

        foreach (var r in valid)
        {
            var src  = long.Parse(r.SourceId.Trim());
            var dst  = long.Parse(r.TargetId.Trim());
            var step = Math.Abs(dst - src);
            if (step == 0) step = 1;
            if (srcDelta != 0)
                r.SourceId = Math.Max(1, src + srcDelta * step).ToString();
            if (dstDelta != 0)
                r.TargetId = Math.Max(1, dst + dstDelta * step).ToString();
        }
        var which = (srcDelta != 0 ? "源ID" : "目标ID");
        var dir   = ((srcDelta + dstDelta) > 0 ? "＋" : "－");
        StatusText.Text = $"{which} {dir}1步 × {valid.Count} 行。";
    }

    private void RemoveRow_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is CloneIdRow row && _rows.Count > 1)
            _rows.Remove(row);
    }

    private void HistoryList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        // 单击只预览提示，不立即应用
    }

    private void HistoryList_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Enter)
        {
            if (HistoryList.SelectedItem is HistoryEntry entry) ApplyHistoryEntry(entry);
            e.Handled = true;
        }
        else if (e.Key == System.Windows.Input.Key.Escape)
        {
            Close();
            e.Handled = true;
        }
    }

    private void HistoryList_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (HistoryList.SelectedItem is HistoryEntry entry) ApplyHistoryEntry(entry);
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
        RemarkBox.Text    = entry.Remark;
        RemarkDstBox.Text = entry.RemarkDst;

        // 恢复决策状态（有历史值则填充，null 时用默认值：美术=替换，子表=保留，自增=是）
        ReplaceArtCheck.IsChecked      = entry.ReplaceArt      ?? true;
        ReplaceSubCheck.IsChecked      = entry.ReplaceSubTable  ?? false;
        IncrementRemarkCheck.IsChecked = entry.IncrementRemark  ?? true;

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

    private void ResetBlockedAlien_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryList.SelectedItem is not HistoryEntry entry) return;
        var groupId = entry.GroupId;
        var targets = string.IsNullOrEmpty(groupId)
            ? new[] { entry }
            : _history.Where(h => h.GroupId == groupId).ToArray();
        var changed = false;
        foreach (var e2 in targets)
            if (e2.BlockedAlienPrefixes.Count > 0)
            { e2.BlockedAlienPrefixes.Clear(); changed = true; }
        if (changed)
        {
            SaveHistory();
            StatusText.Text = $"已重置「{(string.IsNullOrEmpty(groupId) ? entry.Label : groupId)}」的异构ID屏蔽列表。";
        }
        else
            StatusText.Text = "该活动组暂无屏蔽记录。";
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

        // 弹出命名对话框
        var existingGroups = _history
            .Where(h => !string.IsNullOrEmpty(h.GroupId))
            .Select(h => h.GroupId)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        // 推断默认组名：取最近一条有组名的历史
        var lastEntry    = _history.FirstOrDefault(h => !string.IsNullOrEmpty(h.GroupId));
        var suggestGroup = lastEntry?.GroupId ?? RemarkBox.Text.Trim();

        // 推断期次：从当前目标ID尾号推算（如目标ID=73602，尾号2 → 第2期）
        var suggestInstance = "";
        var firstDstId = valid
            .Where(r => string.IsNullOrEmpty(r.BoundTable))
            .Select(r => r.TargetId.Trim())
            .FirstOrDefault(s => long.TryParse(s, out _));
        if (firstDstId != null)
        {
            // 取末尾连续数字作为期号（末尾1~2位，如 73602→2，736012→12）
            var i = firstDstId.Length - 1;
            while (i > 0 && char.IsDigit(firstDstId[i])) i--;
            // i 停在最后一段纯数字前一个字符位置，但整串都是数字时取后1~2位
            var tailStart = firstDstId.Length >= 2 && firstDstId[^2] != '0'
                ? firstDstId.Length - 2
                : firstDstId.Length - 1;
            var tail = firstDstId[tailStart..].TrimStart('0');
            if (!string.IsNullOrEmpty(tail) && int.TryParse(tail, out var phase) && phase > 0)
                suggestInstance = $"第{phase}期";
        }

        var nameWin = new CloneHistoryNameWindow(existingGroups, suggestGroup, suggestInstance)
        {
            Owner = this,
        };
        nameWin.ShowDialog();

        if (nameWin.Result == CloneHistoryNameWindow.DialogResult.Cancel) return;

        var groupId      = nameWin.Result == CloneHistoryNameWindow.DialogResult.Confirm ? nameWin.GroupId      : "";
        var instanceName = nameWin.Result == CloneHistoryNameWindow.DialogResult.Confirm ? nameWin.InstanceName : "";

        // 历史先不带表选择，RunInternal 完成后再追加 table selections
        PushCurrentToHistory(null, groupId, instanceName);
        ResultRows             = valid;
        ResultRemark           = RemarkBox.Text.Trim();
        ResultRemarkDst        = RemarkDstBox.Text.Trim();
        ResultReplaceArt       = ReplaceArtCheck.IsChecked;
        ResultReplaceSubTable  = ReplaceSubCheck.IsChecked;
        ResultIncrementRemark  = IncrementRemarkCheck.IsChecked == true;
        ResultSuspectDecisions = _suspectDecisions.ToList();
        Confirmed = true;
        Close();
    }

    /// <summary>
    /// RunInternal 完成表列表发现后回写，把 tableSelections 补存到最近一条历史记录里。
    /// 历史已由 Run_Click 写入，这里只更新最新一条的 TableSelections 字段。
    /// </summary>
    public void UpdateHistoryWithSuspectDecisions(List<SuspectDecision> decisions)
    {
        if (_history.Count == 0) return;
        var latest = _history[0];
        var newDecisions = decisions.Select(d => new SuspectDecision { TableName = d.TableName, Include = d.Include }).ToList();
        // 内容相同则不写文件
        if (latest.SuspectDecisions.Count == newDecisions.Count &&
            latest.SuspectDecisions.Zip(newDecisions).All(p => p.First.TableName == p.Second.TableName && p.First.Include == p.Second.Include))
            return;
        latest.SuspectDecisions = newDecisions;
        SaveHistory();
    }

    /// <summary>将一张表的一批异构前缀追加到当前活动组所有 entry 的屏蔽列表中。</summary>
    public void UpdateHistoryWithBlockedAlienPrefixes(string tableName, IEnumerable<string> prefixes)
    {
        if (_history.Count == 0) return;
        var groupId = _history[0].GroupId;
        var targets = string.IsNullOrEmpty(groupId)
            ? new[] { _history[0] }
            : _history.Where(h => h.GroupId == groupId).ToArray();
        var prefixList = prefixes.ToList();
        var changed = false;
        foreach (var entry in targets)
        {
            if (!entry.BlockedAlienPrefixes.ContainsKey(tableName))
                entry.BlockedAlienPrefixes[tableName] = new();
            var list = entry.BlockedAlienPrefixes[tableName];
            foreach (var p in prefixList)
                if (!list.Contains(p, StringComparer.OrdinalIgnoreCase))
                { list.Add(p); changed = true; }
        }
        if (changed) SaveHistory();
    }

    /// <summary>清空当前活动组所有 entry 在指定表上的屏蔽前缀（重置）。</summary>
    public void ClearBlockedAlienPrefixes(string tableName)
    {
        if (_history.Count == 0) return;
        var groupId = _history[0].GroupId;
        var targets = string.IsNullOrEmpty(groupId)
            ? new[] { _history[0] }
            : _history.Where(h => h.GroupId == groupId).ToArray();
        var changed = false;
        foreach (var e in targets)
            if (e.BlockedAlienPrefixes.Remove(tableName)) changed = true;
        if (changed) SaveHistory();
    }

    public void UpdateHistoryWithTableSelections(List<TableSelection> tableSelections)
    {
        if (_history.Count == 0) return;
        var latest = _history[0];
        var newSelections = tableSelections.Select(t => new TableSelection { TableName = t.TableName, Selected = t.Selected }).ToList();

        // 更新 tooltip 里的跳过列表
        var skipped = newSelections.Where(t => !t.Selected).Select(t => t.TableName).ToList();
        var tooltip = latest.Tooltip;
        var skipLine = tooltip.Split('\n').FirstOrDefault(l => l.StartsWith("跳过表："));
        if (skipLine != null) tooltip = tooltip.Replace("\n" + skipLine, "");
        if (skipped.Count > 0)
            tooltip += $"\n跳过表：{string.Join("  ", skipped)}";

        // 内容相同则不写文件
        var selectionsUnchanged = latest.TableSelections.Count == newSelections.Count &&
            latest.TableSelections.Zip(newSelections).All(p => p.First.TableName == p.Second.TableName && p.First.Selected == p.Second.Selected);
        if (selectionsUnchanged && latest.Tooltip == tooltip) return;

        latest.TableSelections = newSelections;
        latest.Tooltip = tooltip;
        SaveHistory();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    private void BlockAlien_Click(object sender, RoutedEventArgs e)
    {
        // 把当前窗口中所有行的 SourceId（即异构前缀）打包，供调用方写入屏蔽列表
        ResultBlockAlienTable = true;
        ResultBlockAlienPrefixes = _rows
            .Where(r => !string.IsNullOrWhiteSpace(r.SourceId))
            .Select(r => r.SourceId.Trim())
            .ToList();
        Confirmed = true;
        Close();
    }

    // ── 外部调用辅助 ──────────────────────────────────────────────────────────

    /// <summary>Returns the saved table selections from the last applied history entry (may be empty).</summary>
    public List<TableSelection> SavedTableSelections { get; private set; } = new();

    /// <summary>
    /// 返回当前活动组（GroupId）下所有历史条目合并后的屏蔽前缀字典。
    /// key=表名, value=该表下所有被屏蔽的异构前缀（去重）。
    /// GroupId 为空时退化到只看最新一条。
    /// </summary>
    public Dictionary<string, List<string>> CurrentBlockedAlienPrefixes
    {
        get
        {
            if (_history.Count == 0) return new();
            var groupId = _history[0].GroupId;
            var sources = string.IsNullOrEmpty(groupId)
                ? new[] { _history[0] }
                : _history.Where(h => h.GroupId == groupId);
            var result = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in sources)
                foreach (var (table, prefixes) in entry.BlockedAlienPrefixes)
                {
                    if (!result.ContainsKey(table)) result[table] = new();
                    foreach (var p in prefixes)
                        if (!result[table].Contains(p, StringComparer.OrdinalIgnoreCase))
                            result[table].Add(p);
                }
            return result;
        }
    }

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
    private void Window_EscClose(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == System.Windows.Input.Key.Escape) Close();
    }

}
