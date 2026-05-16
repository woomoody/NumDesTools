using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using ExcelDna.Integration;
using MahApps.Metro.Controls;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace NumDesTools.UI
{
    // ── 单行规则（双向绑定）──────────────────────────────
    public class RuleRow : INotifyPropertyChanged
    {
        private string _from = "";
        private string _to = "";
        public string From
        {
            get => _from;
            set
            {
                _from = value;
                OnPropertyChanged();
            }
        }
        public string To
        {
            get => _to;
            set
            {
                _to = value;
                OnPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? n = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    // ── 历史条目 ─────────────────────────────────────────
    public class HistoryEntry
    {
        public List<(string From, string To)> Rules { get; set; } = [];
        public string Display =>
            Rules.Count > 0
                ? $"{Rules[0].From}→{Rules[0].To}{(Rules.Count > 1 ? $" +{Rules.Count - 1}" : "")}"
                : "";
        public string Detail => string.Join("\n", Rules.Select(r => $"{r.From} → {r.To}"));
    }

    // ── 主窗口 ───────────────────────────────────────────
    public partial class BatchReplaceWindow : MetroWindow
    {
        private static BatchReplaceWindow? _instance;
        private static readonly string HistoryFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "NumDesTools",
            "batch_replace_history.json"
        );

        private readonly ObservableCollection<RuleRow> _rows = [];
        private readonly ObservableCollection<HistoryEntry> _history = [];

        // 外部回调：执行替换时由 NumDesAddIn 注入
        public static Action<List<(string From, string To)>>? OnExecute { get; set; }

        public static BatchReplaceWindow GetOrCreate()
        {
            if (_instance == null || !_instance.IsLoaded)
            {
                _instance = new BatchReplaceWindow();
                _instance.Show();
            }
            return _instance;
        }

        private BatchReplaceWindow()
        {
            MahAppsHelper.EnsureInitialized();
            InitializeComponent();
            RuleRows.ItemsSource = _rows;
            HistoryList.ItemsSource = _history;

            // 挂到 Excel 主窗口，使 WPF 窗口能正常接收键盘输入
            var helper = new WindowInteropHelper(this);
            helper.Owner = (IntPtr)ExcelDnaUtil.WindowHandle;

            LoadHistory();
            if (_rows.Count == 0)
                AddEmptyRow();

            Activated += (_, _) => FocusRowAt(0);
        }

        // ── 公开：从外部触发执行（快捷键重复按） ────────
        public void TriggerExecute() => DoExecute();

        // ── 行管理 ──────────────────────────────────────
        private void AddEmptyRow() => _rows.Add(new RuleRow());

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            AddEmptyRow();
            FocusLastRow();
        }

        private void RemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is RuleRow row)
                _rows.Remove(row);
            if (_rows.Count == 0)
                AddEmptyRow();
        }

        private void FocusLastRow() => FocusRowAt(_rows.Count - 1);

        private void FocusRowAt(int index)
        {
            Dispatcher.InvokeAsync(
                () =>
                {
                    var container = RuleRows.ItemContainerGenerator.ContainerFromIndex(index);
                    (container as FrameworkElement)?.MoveFocus(
                        new TraversalRequest(FocusNavigationDirection.First)
                    );
                },
                System.Windows.Threading.DispatcherPriority.Loaded
            );
        }

        // ── 执行替换 ────────────────────────────────────
        private void Execute_Click(object sender, RoutedEventArgs e) => DoExecute();

        private void DoExecute()
        {
            var rules = _rows
                .Where(r => !string.IsNullOrEmpty(r.From))
                .Select(r => (r.From, r.To))
                .ToList();

            if (rules.Count == 0)
            {
                StatusText.Text = "请至少填写一条规则（查找值不能为空）";
                StatusText.Foreground = System.Windows.Media.Brushes.OrangeRed;
                return;
            }

            OnExecute?.Invoke(rules);
            SaveToHistory(rules);
        }

        public void SetStatus(string msg, bool ok)
        {
            StatusText.Text = msg;
            StatusText.Foreground = ok
                ? System.Windows.Media.Brushes.Green
                : System.Windows.Media.Brushes.OrangeRed;
        }

        public static void SetStatusStatic(string msg, bool ok) =>
            _instance?.Dispatcher.BeginInvoke(() => _instance?.SetStatus(msg, ok));

        // ── 清空 ────────────────────────────────────────
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            AddEmptyRow();
            StatusText.Text = "";
        }

        // ── 历史 ────────────────────────────────────────
        private void HistoryList_SelectionChanged(
            object sender,
            System.Windows.Controls.SelectionChangedEventArgs e
        )
        {
            if (HistoryList.SelectedItem is HistoryEntry entry)
            {
                _rows.Clear();
                foreach (var (from, to) in entry.Rules)
                    _rows.Add(new RuleRow { From = from, To = to });
                HistoryList.SelectedItem = null;
            }
        }

        private void ClearHistory_Click(object sender, RoutedEventArgs e)
        {
            _history.Clear();
            SaveHistory();
        }

        private void SaveToHistory(List<(string From, string To)> rules)
        {
            var entry = new HistoryEntry { Rules = rules };
            // 去重（相同规则集）
            var dup = _history.FirstOrDefault(h =>
                h.Rules.Count == rules.Count
                && h.Rules.Zip(rules)
                    .All(p => p.First.From == p.Second.From && p.First.To == p.Second.To)
            );
            if (dup != null)
                _history.Remove(dup);

            _history.Insert(0, entry);
            while (_history.Count > 30)
                _history.RemoveAt(_history.Count - 1);
            SaveHistory();
        }

        private void LoadHistory()
        {
            try
            {
                if (!File.Exists(HistoryFile))
                    return;
                var json = File.ReadAllText(HistoryFile);
                var list = JsonSerializer.Deserialize<List<HistorySerialized>>(json);
                if (list == null)
                    return;
                foreach (var item in list)
                    _history.Add(
                        new HistoryEntry { Rules = item.Rules.Select(r => (r.From, r.To)).ToList() }
                    );
            }
            catch { }
        }

        private void SaveHistory()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(HistoryFile)!);
                var list = _history
                    .Select(h => new HistorySerialized
                    {
                        Rules = h
                            .Rules.Select(r => new RuleSerialized { From = r.From, To = r.To })
                            .ToList()
                    })
                    .ToList();
                File.WriteAllText(
                    HistoryFile,
                    JsonSerializer.Serialize(
                        list,
                        new JsonSerializerOptions { WriteIndented = true }
                    )
                );
            }
            catch { }
        }

        // ── 置顶 ────────────────────────────────────────
        private void TopmostCheck_Changed(object sender, RoutedEventArgs e) =>
            Topmost = TopmostCheck.IsChecked == true;

        // ── 键盘 ────────────────────────────────────────
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Hide();
                e.Handled = true;
            }
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) != 0)
            {
                DoExecute();
                e.Handled = true;
            }
        }

        // 关闭时隐藏，不销毁
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }

        // ── 序列化 DTOs ─────────────────────────────────
        private class RuleSerialized
        {
            public string From { get; set; } = "";
            public string To { get; set; } = "";
        }

        private class HistorySerialized
        {
            public List<RuleSerialized> Rules { get; set; } = [];
        }
    }
}
