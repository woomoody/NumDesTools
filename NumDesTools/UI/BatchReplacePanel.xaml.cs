using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using WinInput = System.Windows.Input;

namespace NumDesTools.UI
{
    public partial class BatchReplacePanel : System.Windows.Controls.UserControl, INotifyPropertyChanged
    {
        // ── 主题色属性（XAML绑定）────────────────────────
        public SolidColorBrush BgMain    { get; private set; } = new(Colors.White);
        public SolidColorBrush BgPanel   { get; private set; } = new(Colors.White);
        public SolidColorBrush BgInput   { get; private set; } = new(Colors.White);
        public SolidColorBrush FgMain    { get; private set; } = new(Colors.Black);
        public SolidColorBrush FgDim     { get; private set; } = new(Colors.Gray);
        public SolidColorBrush BorderCol { get; private set; } = new(Colors.Gray);
        public SolidColorBrush AccentCol { get; private set; } = new(Colors.DodgerBlue);

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Notify(string n) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));

        // ── 历史存储 ─────────────────────────────────────
        private static readonly string HistoryFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "NumDesTools", "batch_replace_history.json");

        private readonly ObservableCollection<RuleRow>      _rows    = [];
        private readonly ObservableCollection<HistoryEntry> _history = [];

        // 执行替换的外部回调
        public static Action<List<(string From, string To)>>? OnExecute { get; set; }

        // 占位提示属性
        public string PlaceholderText => _rows.Count == 1 && string.IsNullOrEmpty(_rows[0].From)
            ? "输入查找值开始替换..." : "";

        public BatchReplacePanel()
        {
            DataContext = this;
            InitializeComponent();
            RuleRows.ItemsSource    = _rows;
            HistoryList.ItemsSource = _history;
            ApplyTheme();
            LoadHistory();
            // 有历史则加载最近一条，否则空行+提示
            if (_history.Count > 0)
            {
                foreach (var (from, to) in _history[0].Rules)
                    _rows.Add(new RuleRow { From = from, To = to });
            }
            else
            {
                AddEmptyRow();
            }
            _rows.CollectionChanged += (_, _) =>
            {
                foreach (var row in _rows) row.PropertyChanged -= OnRowChanged;
                foreach (var row in _rows) row.PropertyChanged += OnRowChanged;
                Notify(nameof(PlaceholderText));
            };
        }

        private void OnRowChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
            => Notify(nameof(PlaceholderText));

        // ── 主题 ─────────────────────────────────────────
        private static SolidColorBrush B(byte r, byte g, byte b)
            => new(System.Windows.Media.Color.FromRgb(r, g, b));

        private void ApplyTheme()
        {
            bool dark = IsDarkMode();
            if (dark)
            {
                BgMain    = B(0x1E, 0x1E, 0x1E);
                BgPanel   = B(0x16, 0x16, 0x16);
                BgInput   = B(0x2D, 0x2D, 0x2D);
                FgMain    = B(0xD4, 0xD4, 0xD4);
                FgDim     = B(0x88, 0x88, 0x88);
                BorderCol = B(0x55, 0x55, 0x55);
                AccentCol = B(0x0E, 0x63, 0x9C);
            }
            else
            {
                BgMain    = new SolidColorBrush(Colors.White);
                BgPanel   = B(0xF3, 0xF3, 0xF3);
                BgInput   = new SolidColorBrush(Colors.White);
                FgMain    = B(0x1E, 0x1E, 0x1E);
                FgDim     = B(0x66, 0x66, 0x66);
                BorderCol = B(0xCC, 0xCC, 0xCC);
                AccentCol = B(0x00, 0x78, 0xD4);
            }
            Notify(nameof(BgMain));    Notify(nameof(BgPanel));
            Notify(nameof(BgInput));   Notify(nameof(FgMain));
            Notify(nameof(FgDim));     Notify(nameof(BorderCol));
            Notify(nameof(AccentCol));
        }

        private static bool IsDarkMode()
        {
            try
            {
                var v = Registry.GetValue(
                    @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                    "AppsUseLightTheme", 1);
                return v is int i && i == 0;
            }
            catch { return false; }
        }

        // ── 公开：外部聚焦到第一行 ───────────────────────
        public void FocusFirst()
        {
            Dispatcher.InvokeAsync(() =>
            {
                var container = RuleRows.ItemContainerGenerator.ContainerFromIndex(0);
                (container as FrameworkElement)?.MoveFocus(new WinInput.TraversalRequest(WinInput.FocusNavigationDirection.First));
            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        // ── 行管理 ───────────────────────────────────────
        private void AddEmptyRow() => _rows.Add(new RuleRow());

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            AddEmptyRow();
            Dispatcher.InvokeAsync(() =>
            {
                var container = RuleRows.ItemContainerGenerator.ContainerFromIndex(_rows.Count - 1);
                (container as FrameworkElement)?.MoveFocus(new WinInput.TraversalRequest(WinInput.FocusNavigationDirection.First));
            }, System.Windows.Threading.DispatcherPriority.Loaded);
        }

        private void RemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is RuleRow row)
                _rows.Remove(row);
            if (_rows.Count == 0) AddEmptyRow();
        }

        // ── 执行 ─────────────────────────────────────────
        private void Execute_Click(object sender, RoutedEventArgs e) => DoExecute();

        public void DoExecute()
        {
            var rules = _rows
                .Where(r => !string.IsNullOrEmpty(r.From))
                .Select(r => (r.From, r.To))
                .ToList();

            if (rules.Count == 0)
            {
                SetStatus("请至少填写一条规则（查找值不能为空）", false);
                return;
            }

            OnExecute?.Invoke(rules);
            SaveToHistory(rules);
        }

        public void SetStatus(string msg, bool ok)
        {
            StatusText.Text       = msg;
            StatusText.Foreground = ok
                ? B(0x4E, 0xC9, 0xB0)
                : new SolidColorBrush(Colors.OrangeRed);
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            AddEmptyRow();
            StatusText.Text = "";
        }

        // ── 历史 ─────────────────────────────────────────
        private void HistoryList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (HistoryList.SelectedItem is HistoryEntry entry)
            {
                _rows.Clear();
                foreach (var (from, to) in entry.Rules)
                    _rows.Add(new RuleRow { From = from, To = to });
                HistoryList.SelectedItem = null;
            }
        }

        private void DeleteHistory_Click(object sender, RoutedEventArgs e)
        {
            e.Handled = true;
            if (sender is System.Windows.Controls.Button btn && btn.Tag is HistoryEntry entry)
            {
                _history.Remove(entry);
                SaveHistory();
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
            var dup = _history.FirstOrDefault(h =>
                h.Rules.Count == rules.Count &&
                h.Rules.Zip(rules).All(p => p.First.From == p.Second.From && p.First.To == p.Second.To));
            if (dup != null) _history.Remove(dup);
            _history.Insert(0, entry);
            while (_history.Count > 30) _history.RemoveAt(_history.Count - 1);
            SaveHistory();
        }

        private void LoadHistory()
        {
            try
            {
                if (!File.Exists(HistoryFile)) return;
                var list = JsonSerializer.Deserialize<List<HistorySerialized>>(File.ReadAllText(HistoryFile));
                if (list == null) return;
                foreach (var item in list)
                    _history.Add(new HistoryEntry { Rules = item.Rules.Select(r => (r.From, r.To)).ToList() });
            }
            catch { }
        }

        private void SaveHistory()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(HistoryFile)!);
                var list = _history.Select(h => new HistorySerialized
                {
                    Rules = h.Rules.Select(r => new RuleSerialized { From = r.From, To = r.To }).ToList()
                }).ToList();
                File.WriteAllText(HistoryFile, JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { }
        }

        // ── 键盘 ─────────────────────────────────────────
        private void Grid_KeyDown(object sender, WinInput.KeyEventArgs e)
        {
            if (e.Key == WinInput.Key.Enter && (WinInput.Keyboard.Modifiers & WinInput.ModifierKeys.Control) != 0)
            { DoExecute(); e.Handled = true; }
        }

        // ── 序列化 DTOs ──────────────────────────────────
        private class RuleSerialized   { public string From { get; set; } = ""; public string To { get; set; } = ""; }
        private class HistorySerialized { public List<RuleSerialized> Rules { get; set; } = []; }
    }
}
