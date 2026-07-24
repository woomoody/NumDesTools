using System.Data;
using System.Diagnostics;
using OfficeOpenXml;

namespace XlsxEditorPoc;

/// <summary>
/// 最小实现，验证 xlsx 轻量编辑器技术路线：EPPlus 读写 + WinForms DataGridView，
/// 支持增/删行、改值、备注只读预览，不保留公式/图表/透视表。
/// </summary>
internal sealed class MainForm : Form
{
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };
    private readonly Label _status = new() { Dock = DockStyle.Bottom, Height = 24, TextAlign = ContentAlignment.MiddleLeft };
    private readonly Dictionary<TabPage, SheetState> _sheets = new();
    private string? _filePath;

    public MainForm(string? filePath)
    {
        Text = "xlsx 轻量编辑器 POC";
        WindowState = FormWindowState.Maximized;
        KeyPreview = true;

        var toolbar = BuildToolbar();
        Controls.Add(_status);
        Controls.Add(_tabs);
        Controls.Add(toolbar);

        KeyDown += OnKeyDown;

        if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
        {
            LoadFile(filePath);
        }
    }

    private ToolStrip BuildToolbar()
    {
        var open = new ToolStripButton("打开(Ctrl+O)") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        open.Click += (_, _) => OpenFileDialog();

        var save = new ToolStripButton("保存(Ctrl+S)") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        save.Click += (_, _) => SaveFile();

        var addRow = new ToolStripButton("增行(Ctrl+N)") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        addRow.Click += (_, _) => AddRow();

        var delRow = new ToolStripButton("删行(Ctrl+D)") { DisplayStyle = ToolStripItemDisplayStyle.Text };
        delRow.Click += (_, _) => DeleteSelectedRows();

        return new ToolStrip(open, save, new ToolStripSeparator(), addRow, delRow) { Dock = DockStyle.Top };
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.O) { OpenFileDialog(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.S) { SaveFile(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.N) { AddRow(); e.Handled = true; }
        else if (e.Control && e.KeyCode == Keys.D) { DeleteSelectedRows(); e.Handled = true; }
        else if (e.KeyCode == Keys.Escape) { Close(); }
    }

    private void OpenFileDialog()
    {
        using var dlg = new OpenFileDialog { Filter = "Excel (*.xlsx)|*.xlsx" };
        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            LoadFile(dlg.FileName);
        }
    }

    private async void LoadFile(string path)
    {
        _tabs.Enabled = false;
        Cursor.Current = Cursors.WaitCursor;
        _status.Text = $"正在加载：{Path.GetFileName(path)}…";

        var sw = Stopwatch.StartNew();
        var built = await Task.Run(() => BuildAllSheets(path));

        _filePath = path;
        _tabs.TabPages.Clear();
        _sheets.Clear();
        foreach (var (name, table, comments) in built)
        {
            var page = new TabPage(name);
            var grid = BuildGrid(table, comments);
            page.Controls.Add(grid);
            _tabs.TabPages.Add(page);
            _sheets[page] = new SheetState(name, table, comments);
        }

        sw.Stop();
        _tabs.Enabled = true;
        Cursor.Current = Cursors.Default;
        _status.Text = $"已加载：{Path.GetFileName(path)}（{_sheets.Count} 个 Sheet，耗时 {sw.ElapsedMilliseconds} ms）";
    }

    /// <summary>
    /// 后台线程执行：一次性按 range 枚举取值（比逐格索引器 Cells[r,c] 快一个量级），
    /// 避免大表（如 6.5 万行 x 85 列）在 UI 线程卡死。
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

    private static DataGridView BuildGrid(DataTable table, Dictionary<(int Row, int Col), string> commentMap)
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            DataSource = table,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            ShowCellToolTips = true,
            RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing,
        };
        grid.CellToolTipTextNeeded += (_, e) =>
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            if (commentMap.TryGetValue((e.RowIndex + 1, e.ColumnIndex + 1), out var text))
            {
                e.ToolTipText = text;
            }
        };

        return grid;
    }

    private void AddRow()
    {
        if (_tabs.SelectedTab is not { } page || !_sheets.TryGetValue(page, out var state)) return;
        state.Table.Rows.Add(state.Table.NewRow());
        state.Dirty = true;
    }

    private void DeleteSelectedRows()
    {
        if (_tabs.SelectedTab is not { } page || !_sheets.TryGetValue(page, out var state)) return;
        if (page.Controls[0] is not DataGridView grid) return;

        foreach (DataGridViewRow row in grid.SelectedRows)
        {
            if (row.DataBoundItem is DataRowView view)
            {
                view.Row.Delete();
            }
        }

        state.Dirty = true;
    }

    private void SaveFile()
    {
        if (_filePath is null) return;
        var sw = Stopwatch.StartNew();

        using var package = new ExcelPackage(new FileInfo(_filePath));
        foreach (var state in _sheets.Values)
        {
            var sheet = package.Workbook.Worksheets[state.SheetName];
            var table = state.Table;
            table.AcceptChanges();

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
        _status.Text = $"已保存：{Path.GetFileName(_filePath)}（耗时 {sw.ElapsedMilliseconds} ms）";
    }

    private sealed class SheetState(string sheetName, DataTable table, Dictionary<(int Row, int Col), string> comments)
    {
        public string SheetName { get; } = sheetName;
        public DataTable Table { get; } = table;
        public Dictionary<(int Row, int Col), string> Comments { get; } = comments;
        public bool Dirty { get; set; }
    }
}
