using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Microsoft.Win32;
using WinInput = System.Windows.Input;

namespace NumDesTools.UI
{
    public partial class BatchReplacePanel
        : System.Windows.Controls.UserControl
    {
        private static readonly string HistoryFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "NumDesTools",
            "batch_replace_history.json"
        );

        private readonly ObservableCollection<RuleRow> _rows = [];
        private readonly ObservableCollection<HistoryEntry> _history = [];

        public static Action<List<(string From, string To)>>? OnExecute { get; set; }

        public BatchReplacePanel()
        {
            InitializeComponent();
            RuleRows.ItemsSource = _rows;
            HistoryList.ItemsSource = _history;
            LoadHistory();
            if (_history.Count > 0)
                foreach (var (from, to) in _history[0].Rules)
                    _rows.Add(new RuleRow { From = from, To = to });
            else
                AddEmptyRow();

            _rows.CollectionChanged += OnRowsCollectionChanged;
        }

        private void OnRowsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
                foreach (RuleRow row in e.OldItems)
                    row.PropertyChanged -= OnRowPropertyChanged;
            if (e.NewItems != null)
                foreach (RuleRow row in e.NewItems)
                    row.PropertyChanged += OnRowPropertyChanged;
        }

        private void OnRowPropertyChanged(object? sender, PropertyChangedEventArgs e) { }

        private void AddEmptyRow() => _rows.Add(new RuleRow());

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            AddEmptyRow();
            Dispatcher.InvokeAsync(
                () =>
                {
                    var container = RuleRows.ItemContainerGenerator.ContainerFromIndex(
                        _rows.Count - 1
                    );
                    (container as FrameworkElement)?.MoveFocus(
                        new WinInput.TraversalRequest(WinInput.FocusNavigationDirection.First)
                    );
                },
                System.Windows.Threading.DispatcherPriority.Loaded
            );
        }

        private void RemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is RuleRow row)
                _rows.Remove(row);
            if (_rows.Count == 0)
                AddEmptyRow();
        }

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
            StatusText.Text = msg;
            StatusText.Foreground = ok
                ? TryFindBrush("SystemFillColorSuccessBrush")
                : TryFindBrush("SystemFillColorCriticalBrush");
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            _rows.Clear();
            AddEmptyRow();
            StatusText.Text = "";
        }

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
                var list = JsonSerializer.Deserialize<List<HistorySerialized>>(
                    File.ReadAllText(HistoryFile)
                );
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
                            .ToList(),
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

        private void Grid_KeyDown(object sender, WinInput.KeyEventArgs e)
        {
            if (
                e.Key == WinInput.Key.Enter
                && (WinInput.Keyboard.Modifiers & WinInput.ModifierKeys.Control) != 0
            )
            {
                DoExecute();
                e.Handled = true;
            }
        }

        private static Brush TryFindBrush(string key) =>
            (Brush)System.Windows.Application.Current.TryFindResource(key) ?? Brushes.Gray;

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