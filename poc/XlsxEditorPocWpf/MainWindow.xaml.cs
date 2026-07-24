using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using OfficeOpenXml;

namespace XlsxEditorPocWpf;

/// <summary>
/// WPF + MahApps.Metro 版本，验证"现代 UI 外观 + 速度优先"能否兼得：
/// 数据层沿用 WinForms POC 已验证的后台线程 + range 批量枚举懒加载方案，
/// UI 层换成 WPF DataGrid（原生虚拟化）+ MetroWindow 深色主题。
/// </summary>
public partial class MainWindow
{
    private readonly Dictionary<TabItem, SheetState> _sheets = new();
    private string? _filePath;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers == ModifierKeys.Control)
        {
            switch (e.Key)
            {
                case Key.O: OnOpenClick(sender, e); e.Handled = true; break;
                case Key.S: OnSaveClick(sender, e); e.Handled = true; break;
                case Key.N: OnAddRowClick(sender, e); e.Handled = true; break;
                case Key.D: OnDeleteRowClick(sender, e); e.Handled = true; break;
            }
        }
        else if (e.Key == Key.Escape)
        {
            Close();
        }
    }

    private void OnOpenClick(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
        if (dlg.ShowDialog(this) == true)
        {
            LoadFile(dlg.FileName);
        }
    }

    internal async void LoadFile(string path)
    {
        Tabs.IsEnabled = false;
        Cursor = Cursors.Wait;
        StatusText.Text = $"正在加载：{Path.GetFileName(path)}…";

        var sw = Stopwatch.StartNew();
        var built = await Task.Run(() => BuildAllSheets(path));

        _filePath = path;
        Tabs.Items.Clear();
        _sheets.Clear();
        foreach (var (name, table, comments) in built)
        {
            var grid = BuildGrid(table, comments);
            var tab = new TabItem { Header = name, Content = grid };
            Tabs.Items.Add(tab);
            _sheets[tab] = new SheetState(name, table, comments);
        }

        sw.Stop();
        Tabs.IsEnabled = true;
        Cursor = Cursors.Arrow;
        StatusText.Text = $"已加载：{Path.GetFileName(path)}（{_sheets.Count} 个 Sheet，耗时 {sw.ElapsedMilliseconds} ms）";
    }

    /// <summary>
    /// 后台线程执行：range 批量枚举取值，避免大表在 UI 线程卡死（与 WinForms POC 一致）。
    /// </summary>
    private static List<(string Name, DataTable Table, Dictionary<(int Row, int Col), string> Comments)> BuildAllSheets(string path)
    {
        var result = new List<(string, DataTable, Dictionary<(int Row, int Col), string>)>();
        using var package = new ExcelPackage(new FileInfo(path));
        foreach (var sheet in package.Workbook.Worksheets)
        {
            var dims = sheet.Dimension;
            var colCount = dims?.End.Column ?? 0;
            var rowCount = dims?.End.Row ?? 0;
            var grid = new string?[rowCount, colCount];
            var comments = new Dictionary<(int Row, int Col), string>();

            if (dims is not null)
            {
                foreach (var cell in sheet.Cells[1, 1, rowCount, colCount])
                {
                    grid[cell.Start.Row - 1, cell.Start.Column - 1] = cell.Value?.ToString();
                    if (cell.Comment is not null)
                    {
                        comments[(cell.Start.Row, cell.Start.Column)] = cell.Comment.Text;
                    }
                }
            }

            var table = new DataTable();
            for (var c = 1; c <= colCount; c++)
            {
                table.Columns.Add($"C{c}", typeof(string));
            }

            table.BeginLoadData();
            for (var r = 0; r < rowCount; r++)
            {
                var row = table.NewRow();
                for (var c = 0; c < colCount; c++)
                {
                    row[c] = grid[r, c] ?? string.Empty;
                }

                table.Rows.Add(row);
            }

            table.EndLoadData();
            result.Add((sheet.Name, table, comments));
        }

        return result;
    }

    private static DataGrid BuildGrid(DataTable table, Dictionary<(int Row, int Col), string> commentMap)
    {
        var grid = new DataGrid
        {
            ItemsSource = table.DefaultView,
            AutoGenerateColumns = true,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            EnableRowVirtualization = true,
            EnableColumnVirtualization = true,
        };

        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new EventSetter(MouseEnterEvent, new System.Windows.Input.MouseEventHandler((s, _) =>
        {
            if (s is not DataGridCell { DataContext: DataRowView view } cell) return;
            var rowIndex = table.Rows.IndexOf(view.Row);
            var colIndex = cell.Column.DisplayIndex;
            ToolTipService.SetToolTip(cell,
                commentMap.TryGetValue((rowIndex + 1, colIndex + 1), out var text) ? text : null);
        })));
        grid.CellStyle = cellStyle;

        return grid;
    }

    private void OnAddRowClick(object sender, RoutedEventArgs e)
    {
        if (Tabs.SelectedItem is not TabItem tab || !_sheets.TryGetValue(tab, out var state)) return;
        state.Table.Rows.Add(state.Table.NewRow());
    }

    private void OnDeleteRowClick(object sender, RoutedEventArgs e)
    {
        if (Tabs.SelectedItem is not TabItem { Content: DataGrid grid } || !_sheets.TryGetValue((TabItem)Tabs.SelectedItem, out var state)) return;

        foreach (var item in grid.SelectedItems.Cast<DataRowView>().ToList())
        {
            item.Row.Delete();
        }

        state.Table.AcceptChanges();
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        if (_filePath is null) return;
        var sw = Stopwatch.StartNew();

        using var package = new ExcelPackage(new FileInfo(_filePath));
        foreach (var state in _sheets.Values)
        {
            var sheet = package.Workbook.Worksheets[state.SheetName];
            var table = state.Table;

            var existingRows = sheet.Dimension?.End.Row ?? 0;
            if (table.Rows.Count < existingRows)
            {
                sheet.DeleteRow(table.Rows.Count + 1, existingRows - table.Rows.Count);
            }

            for (var r = 0; r < table.Rows.Count; r++)
            {
                for (var c = 0; c < table.Columns.Count; c++)
                {
                    sheet.Cells[r + 1, c + 1].Value = table.Rows[r][c];
                }
            }
        }

        package.Save();
        sw.Stop();
        StatusText.Text = $"已保存：{Path.GetFileName(_filePath)}（耗时 {sw.ElapsedMilliseconds} ms）";
    }

    private sealed class SheetState(string sheetName, DataTable table, Dictionary<(int Row, int Col), string> comments)
    {
        public string SheetName { get; } = sheetName;
        public DataTable Table { get; } = table;
        public Dictionary<(int Row, int Col), string> Comments { get; } = comments;
    }
}
