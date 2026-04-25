using System.Collections.ObjectModel;
using System.Windows;
using Window = System.Windows.Window;

namespace NumDesTools.UI;

public partial class CloneTableSelectionWindow : Window
{
    public bool Confirmed { get; private set; }

    private readonly ObservableCollection<TableSelection> _tables = new();

    public CloneTableSelectionWindow(
        IEnumerable<string> tableNames,
        IEnumerable<TableSelection>? saved = null)
    {
        InitializeComponent();
        TableList.ItemsSource = _tables;

        var savedMap = saved?.ToDictionary(t => t.TableName, t => t.Selected,
                            StringComparer.OrdinalIgnoreCase)
                       ?? new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        foreach (var name in tableNames)
        {
            _tables.Add(new TableSelection
            {
                TableName = name,
                Selected  = savedMap.TryGetValue(name, out var v) ? v : true,
            });
        }

        UpdateStatus();
    }

    public List<TableSelection> Result => _tables.ToList();

    private void UpdateStatus()
    {
        var total    = _tables.Count;
        var selected = _tables.Count(t => t.Selected);
        StatusText.Text = $"共 {total} 张表，已选 {selected} 张";
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _tables) t.Selected = true;
        UpdateStatus();
    }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _tables) t.Selected = false;
        UpdateStatus();
    }

    private void Invert_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _tables) t.Selected = !t.Selected;
        UpdateStatus();
    }

    private void Confirm_Click(object sender, RoutedEventArgs e)
    {
        Confirmed = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
}
