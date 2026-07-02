using ListBox = System.Windows.Forms.ListBox;
using TextBox = System.Windows.Forms.TextBox;

namespace NumDesTools;

// 跨表同步映射的配置窗口：左侧映射列表，右侧编辑详情。
internal sealed class XlsxSyncSettingsForm : Form
{
    private readonly List<XlsxCrossSync.Mapping> _mappings;
    private readonly ListBox _list;
    private readonly TextBox _name;
    private readonly TextBox _sourcePath;
    private readonly TextBox _sourceSheet;
    private readonly TextBox _targetPath;
    private readonly TextBox _targetSheet;
    private readonly TextBox _keyColumn;
    private readonly NumericUpDown _groupPrefixLen;
    private readonly TextBox _forwardColumns;
    private readonly TextBox _reverseColumns;

    internal XlsxSyncSettingsForm(List<XlsxCrossSync.Mapping> mappings)
    {
        _mappings = mappings;

        Text = "跨表同步设置";
        Width = 780;
        Height = 520;
        MinimumSize = new System.Drawing.Size(700, 460);
        StartPosition = FormStartPosition.CenterScreen;
        KeyPreview = true;
        Padding = new Padding(12);
        Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
        KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        };

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _list = new ListBox { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 10, 0) };
        _list.SelectedIndexChanged += (_, _) => LoadSelected();

        var listPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            RowCount = 2,
            ColumnCount = 1,
            Margin = new Padding(0, 0, 10, 0),
        };
        listPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        listPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var listButtons = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0),
        };
        var addButton = new System.Windows.Forms.Button
        {
            Text = "新建",
            Width = 90,
            Height = 30,
        };
        var deleteButton = new System.Windows.Forms.Button
        {
            Text = "删除",
            Width = 90,
            Height = 30,
        };
        addButton.Click += (_, _) => AddNew();
        deleteButton.Click += (_, _) => DeleteSelected();
        listButtons.Controls.Add(addButton);
        listButtons.Controls.Add(deleteButton);

        _list.Margin = new Padding(0);
        listPanel.Controls.Add(_list, 0, 0);
        listPanel.Controls.Add(listButtons, 0, 1);

        var detail = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 10,
            AutoSize = false,
        };
        detail.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < 9; i++)
            detail.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        detail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        _name = NewTextBox();
        _sourcePath = NewTextBox();
        _sourceSheet = NewTextBox();
        _targetPath = NewTextBox();
        _targetSheet = NewTextBox();
        _keyColumn = NewTextBox();
        _groupPrefixLen = new NumericUpDown
        {
            Minimum = 1,
            Maximum = 32,
            Value = 6,
            Width = 80,
            Margin = new Padding(0, 0, 0, 8),
        };
        _forwardColumns = NewTextBox();
        _reverseColumns = NewTextBox();

        AddRow(detail, 0, "名称", _name);
        AddRow(detail, 1, "源文件 (a)", BrowseRow(_sourcePath));
        AddRow(detail, 2, "源 Sheet", _sourceSheet);
        AddRow(detail, 3, "目标文件 (b)", BrowseRow(_targetPath));
        AddRow(detail, 4, "目标 Sheet", _targetSheet);
        AddRow(detail, 5, "Key 列名", _keyColumn);
        AddRow(detail, 6, "分组前缀长度", _groupPrefixLen);
        AddRow(detail, 7, "正向列 a→b（逗号分隔）", _forwardColumns);
        AddRow(detail, 8, "反向列 b→a（逗号分隔，低频）", _reverseColumns);

        var saveButton = new System.Windows.Forms.Button
        {
            Text = "保存这条映射",
            Dock = DockStyle.Bottom,
            Height = 34,
            Margin = new Padding(0, 12, 0, 0),
        };
        saveButton.Click += (_, _) => SaveSelected();
        detail.Controls.Add(saveButton, 0, 9);
        detail.SetColumnSpan(saveButton, 2);

        root.Controls.Add(listPanel, 0, 0);
        root.Controls.Add(detail, 1, 0);
        Controls.Add(root);

        RefreshList();
        if (_mappings.Count > 0)
            _list.SelectedIndex = 0;
    }

    private static TextBox NewTextBox() =>
        new() { Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 8) };

    private static Control BrowseRow(TextBox pathBox)
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 0, 0, 8),
            AutoSize = true,
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70));
        pathBox.Dock = DockStyle.Fill;
        pathBox.Margin = new Padding(0);
        var browseButton = new System.Windows.Forms.Button
        {
            Text = "浏览…",
            Dock = DockStyle.Fill,
            Margin = new Padding(6, 0, 0, 0),
        };
        browseButton.Click += (_, _) =>
        {
            using var dlg = new OpenFileDialog { Filter = "Excel 文件|*.xlsx" };
            if (dlg.ShowDialog() == DialogResult.OK)
                pathBox.Text = dlg.FileName;
        };
        panel.Controls.Add(pathBox, 0, 0);
        panel.Controls.Add(browseButton, 1, 0);
        return panel;
    }

    private static void AddRow(TableLayoutPanel table, int row, string label, Control input)
    {
        var lbl = new System.Windows.Forms.Label
        {
            Text = label,
            Dock = DockStyle.Top,
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
            Margin = new Padding(0, 6, 8, 8),
            AutoSize = true,
        };
        table.Controls.Add(lbl, 0, row);
        table.Controls.Add(input, 1, row);
    }

    private void RefreshList()
    {
        _list.Items.Clear();
        foreach (var m in _mappings)
            _list.Items.Add(m.Name);
    }

    private void LoadSelected()
    {
        if (_list.SelectedIndex < 0 || _list.SelectedIndex >= _mappings.Count)
            return;
        var m = _mappings[_list.SelectedIndex];
        _name.Text = m.Name;
        _sourcePath.Text = m.SourcePath;
        _sourceSheet.Text = m.SourceSheet;
        _targetPath.Text = m.TargetPath;
        _targetSheet.Text = m.TargetSheet;
        _keyColumn.Text = m.KeyColumn;
        _groupPrefixLen.Value = Math.Clamp(m.GroupPrefixLen, 1, 32);
        _forwardColumns.Text = string.Join(",", m.ForwardColumns);
        _reverseColumns.Text = string.Join(",", m.ReverseColumns);
    }

    private void AddNew()
    {
        _mappings.Add(
            new XlsxCrossSync.Mapping(
                $"新映射{_mappings.Count + 1}",
                "",
                "",
                "",
                "",
                "id",
                6,
                [],
                []
            )
        );
        RefreshList();
        _list.SelectedIndex = _mappings.Count - 1;
    }

    private void DeleteSelected()
    {
        if (_list.SelectedIndex < 0)
            return;
        _mappings.RemoveAt(_list.SelectedIndex);
        XlsxCrossSync.SaveMappings(_mappings);
        RefreshList();
    }

    private void SaveSelected()
    {
        if (_list.SelectedIndex < 0)
            return;
        if (string.IsNullOrWhiteSpace(_name.Text))
        {
            MessageBox.Show("名称不能为空。", "同步设置");
            return;
        }

        _mappings[_list.SelectedIndex] = new XlsxCrossSync.Mapping(
            _name.Text.Trim(),
            _sourcePath.Text.Trim(),
            _sourceSheet.Text.Trim(),
            _targetPath.Text.Trim(),
            _targetSheet.Text.Trim(),
            _keyColumn.Text.Trim(),
            (int)_groupPrefixLen.Value,
            SplitColumns(_forwardColumns.Text),
            SplitColumns(_reverseColumns.Text)
        );
        XlsxCrossSync.SaveMappings(_mappings);
        RefreshList();
        _list.SelectedIndex = _list.Items.Count - 1;
        MessageBox.Show("已保存。", "同步设置");
    }

    private static List<string> SplitColumns(string text) =>
        text.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .ToList();
}
