using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OfficeOpenXml;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// WPF + MahApps.Metro 版本。两级 tab：工作簿（带 ✕ 关闭）→ sheet（DataGrid）。
/// 数据层后台线程 + range 批量枚举懒加载；UI 层 WPF DataGrid 虚拟化 + MetroWindow 深色主题。
/// 冻结窗格：列 = 原生 FrozenColumnCount；行 = 双 DataGrid（顶部冻结行 grid + 主 grid，
/// 共享同 DataTable 的两个 ListCollectionView 按行号谓词切分，横向滚 + 列宽 + 冻列数三同步）。
/// </summary>
public partial class MainWindow
{
    private static readonly Brush DirtyCellBrush = new SolidColorBrush(Color.FromRgb(43, 145, 76));

    // 深色主题单元格边框颜色（比背景略亮）
    private static readonly Brush GridLineBrush = new SolidColorBrush(Color.FromRgb(60, 60, 60));

    // 聚光灯：选中单元格所在行列高亮
    private static readonly Brush SpotlightBrush = new SolidColorBrush(
        Color.FromArgb(60, 0, 120, 215)
    );

    // sheet tab -> state（扁平，所有打开工作簿的 sheet 都在这，便于机械替换）
    private readonly Dictionary<TabItem, SheetState> _sheets = new();

    // 工作簿 tab -> 该工作簿的所有 sheet tab（关工作簿时批量清理）
    private readonly Dictionary<TabItem, List<TabItem>> _workbookSheets = new();

    // filePath -> 工作簿 tab（LoadFile 复用 + 关工作簿反查路径）
    private readonly Dictionary<string, TabItem> _workbookByPath = new();

    private readonly HashSet<string> _dirtyFiles = new();
    private SheetState? _activeSheetState;

    public MainWindow()
    {
        InitializeComponent();
    }

    // ── 当前选中定位（集中逻辑，供 ~22 处机械替换用） ──────────────────

    private TabItem? CurrentWorkbookTab => Tabs.SelectedItem as TabItem;

    private static TabControl? GetSheetTabs(TabItem workbookTab) =>
        workbookTab.Content switch
        {
            Border { Child: TabControl sheetTabs } => sheetTabs,
            TabControl sheetTabs => sheetTabs,
            _ => null,
        };

    /// <summary>
    /// 当前工作簿内嵌 TabControl 的选中 sheet tab。
    /// SelectionChanged 是冒泡路由事件，内层 sheet 切换会冒泡到外层 OnTabsSelectionChanged，
    /// 所以这里直接读内层 SelectedItem 即可覆盖两层切换。
    /// </summary>
    private TabItem? CurrentSheetTab =>
        Tabs.SelectedItem is TabItem workbookTab && GetSheetTabs(workbookTab) is { } sheetTabs
            ? sheetTabs.SelectedItem as TabItem
            : null;

    private SheetState? CurrentSheetState =>
        CurrentSheetTab is TabItem t && _sheets.TryGetValue(t, out var s) ? s : null;

    /// <summary>
    /// 当前选中 sheet 所属文件路径（null 表示无选中）。
    /// </summary>
    private string? CurrentFilePath => CurrentSheetState?.FilePath;

    /// <summary>
    /// 当前 sheet 的主 DataGrid（冻结行模式下显示 row N..end；非冻结显示全表）。
    /// 替代旧版 `Content is DataGrid` 的机械解包。
    /// </summary>
    private DataGrid? CurrentMainGrid => CurrentSheetState?.MainGrid;

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers == ModifierKeys.Control)
        {
            switch (e.Key)
            {
                case Key.O:
                    OnOpenClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.S:
                    OnSaveClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.N:
                    OnAddRowClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.D:
                    OnDeleteRowClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.Z:
                    OnUndoClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.Y:
                    OnRedoClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.V:
                    OnPaste(sender, e);
                    e.Handled = true;
                    break;
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
        var built = await Task.Run(() => BuildAllSheetsLazy(path));

        var fileName = Path.GetFileName(path);

        // 复用已有工作簿 tab（同一文件再次打开则追加 sheet），否则建工作簿 tab + 内嵌 sheet TabControl
        TabItem wbTab;
        TabControl sheetTabs;
        if (_workbookByPath.TryGetValue(path, out var existing))
        {
            wbTab = existing;
            sheetTabs = GetSheetTabs(wbTab)!;
        }
        else
        {
            sheetTabs = new TabControl();
            wbTab = BuildWorkbookTab(fileName, sheetTabs);
            _workbookByPath[path] = wbTab;
            _workbookSheets[wbTab] = new List<TabItem>();
            Tabs.Items.Add(wbTab);
        }

        TabItem? firstNewTab = null;
        foreach (var (name, table, comments, totalRows, loadedRows) in built)
        {
            // 双 DataGrid 布局：外层 Grid(panel) 含两行——row0=冻结行 grid（按需出现），row1=主 grid。
            // 主 grid 永远在 row1，冻结行 grid 由 ApplyFreezeLayout 在冻结行时插入 row0。
            var grid = BuildGrid(table, comments);
            var panel = new Grid();
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 冻结行槽（空时 Auto=0 高度）
            panel.RowDefinitions.Add(
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }
            );
            panel.Children.Add(grid);
            Grid.SetRow(grid, 1);

            var sheetTab = new TabItem
            {
                Header = name,
                FontSize = 12,
                Content = panel,
            };
            var state = new SheetState(name, table, comments, path, totalRows, loadedRows)
            {
                MainGrid = grid,
                Panel = panel,
            };
            // 行号头：冻结行模式下主 grid 偏移 FrozenRows；非冻结（FrozenRows=0）原样 1..N。
            // 闭包读 state.FrozenRows，冻结/取消后自动正确。
            grid.LoadingRow += (_, args) =>
                args.Row.Header = state.FrozenRows + args.Row.GetIndex() + 1;
            sheetTabs.Items.Add(sheetTab);
            _sheets[sheetTab] = state;
            _workbookSheets[wbTab].Add(sheetTab);
            firstNewTab ??= sheetTab;

            // 应用冻结配置（列 + 行）
            var (fc, fr) = FreezeConfig.GetFreeze(fileName, name);
            if (fc > 0 && fc <= grid.Columns.Count)
            {
                grid.FrozenColumnCount = fc;
                state.FrozenColumns = fc;
            }
            if (fr > 0 && fr < table.Rows.Count)
            {
                state.FrozenRows = fr;
                ApplyFreezeLayout(state);
            }
        }

        // 选中新工作簿 + 其第一个 sheet
        Tabs.SelectedItem = wbTab;
        if (firstNewTab is not null)
            sheetTabs.SelectedItem = firstNewTab;

        sw.Stop();
        Tabs.IsEnabled = true;
        Cursor = Cursors.Arrow;
        UpdateTitle();
        if (CurrentSheetState is { } currentState)
        {
            StatusText.Text =
                $"已加载首屏 {currentState.LoadedRows}/{currentState.TotalRows} 行（{currentState.SheetName}）";
        }

        foreach (var sheetTab in _workbookSheets[wbTab])
        {
            if (_sheets.TryGetValue(sheetTab, out var state) && state.LoadedRows < state.TotalRows)
            {
                StartBackgroundLoad(state);
            }
        }
    }

    /// <summary>
    /// 建工作簿 tab：Header = 文件名 + ✕ 关闭按钮，Content = 内嵌 sheet TabControl。
    /// </summary>
    private TabItem BuildWorkbookTab(string fileName, TabControl sheetTabs)
    {
        sheetTabs.Padding = new Thickness(4, 0, 4, 2);
        var sheetTabStyle = new Style(
            typeof(TabItem),
            Application.Current.TryFindResource(typeof(TabItem)) as Style
        );
        sheetTabStyle.Setters.Add(new Setter(Control.FontSizeProperty, 12.0));
        sheetTabs.ItemContainerStyle = sheetTabStyle;
        var sheetTabsBorder = new Border
        {
            Margin = new Thickness(0, 6, 0, 0),
            Padding = new Thickness(2),
            BorderBrush = GridLineBrush,
            BorderThickness = new Thickness(0, 2, 0, 0),
            Background = new SolidColorBrush(Color.FromRgb(30, 30, 30)),
            Child = sheetTabs,
        };
        var header = new StackPanel { Orientation = Orientation.Horizontal };
        header.Children.Add(
            new TextBlock { Text = fileName, VerticalAlignment = VerticalAlignment.Center }
        );
        var closeBtn = new Button
        {
            Content = "✕",
            Margin = new Thickness(6, 0, 0, 0),
            Padding = new Thickness(2),
            BorderThickness = new Thickness(0),
            Background = Brushes.Transparent,
            Foreground = Brushes.White,
            Cursor = Cursors.Hand,
        };
        var wbTab = new TabItem { Header = header, Content = sheetTabsBorder };
        closeBtn.Tag = wbTab; // click 时反查所属工作簿 tab
        closeBtn.Click += OnWorkbookCloseClick;
        header.Children.Add(closeBtn);
        return wbTab;
    }

    /// <summary>
    /// ✕ 关闭工作簿：有脏数据先提示保存，然后移除 tab + 清理所有相关 state。
    /// </summary>
    private async void OnWorkbookCloseClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not TabItem wbTab)
            return;
        var filePath = _workbookByPath.FirstOrDefault(kv => kv.Value == wbTab).Key;
        if (filePath is null)
            return;

        if (_dirtyFiles.Contains(filePath))
        {
            Tabs.SelectedItem = wbTab; // 切到该工作簿，让 SaveCurrentFileAsync 取到当前 sheet
            var result = MessageBox.Show(
                this,
                $"{Path.GetFileName(filePath)} 有未保存的更改，是否保存？",
                "关闭工作簿",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question
            );
            switch (result)
            {
                case MessageBoxResult.Yes:
                {
                    var ok = await SaveCurrentFileAsync();
                    if (!ok)
                        return; // 保存失败不关
                    break;
                }
                case MessageBoxResult.Cancel:
                    return;
                case MessageBoxResult.No:
                    break; // 不保存直接关
            }
        }

        // 批量清理该工作簿的所有 sheet state + 索引
        foreach (var sheetTab in _workbookSheets[wbTab])
        {
            _sheets[sheetTab].LoadCts?.Cancel();
            _sheets.Remove(sheetTab);
        }
        _workbookSheets.Remove(wbTab);
        _workbookByPath.Remove(filePath);
        _dirtyFiles.Remove(filePath);
        Tabs.Items.Remove(wbTab);
        UpdateTitle();
    }

    /// <summary>
    /// 后台线程执行：只读取首屏，避免打开大表时等待完整工作表解析。
    /// 列名用 Excel 原生列名（A, B, ... Z, AA, AB, ...）。
    /// </summary>
    private static List<(
        string Name,
        DataTable Table,
        Dictionary<(int Row, int Col), string> Comments,
        int TotalRows,
        int LoadedRows
    )> BuildAllSheetsLazy(string path)
    {
        const int firstScreenRows = 200;
        var result =
            new List<(string, DataTable, Dictionary<(int Row, int Col), string>, int, int)>();
        foreach (var sheetName in OoxmlLazyReader.ReadSheetNames(path))
        {
            if (sheetName.StartsWith('#'))
            {
                continue;
            }

            var (rowCount, colCount) = OoxmlLazyReader.ReadDimension(path, sheetName);
            if (rowCount is 0 || colCount is 0)
            {
                continue;
            }

            var rawRows = OoxmlLazyReader
                .ReadRows(path, sheetName, maxRows: firstScreenRows, skipRows: 0)
                .ToList();
            var comments = new Dictionary<(int Row, int Col), string>();
            var table = new DataTable(sheetName);
            for (var c = 1; c <= colCount; c++)
            {
                // Excel 列名：1→A, 26→Z, 27→AA
                table.Columns.Add(GetExcelColumnName(c), typeof(string));
            }

            table.BeginLoadData();
            foreach (var rawRow in rawRows)
            {
                AddRawRow(table, comments, rawRow);
            }

            table.EndLoadData();
            result.Add((sheetName, table, comments, rowCount, rawRows.Count));
        }

        return result;
    }

    private static void AddRawRow(
        DataTable table,
        Dictionary<(int Row, int Col), string> comments,
        RawRow rawRow
    )
    {
        var row = table.NewRow();
        for (var column = 0; column < table.Columns.Count; column++)
        {
            var columnName = table.Columns[column].ColumnName;
            row[column] = rawRow.Cells.GetValueOrDefault(columnName, string.Empty);
        }

        table.Rows.Add(row);
        foreach (var (cell, comment) in rawRow.Comments)
        {
            comments[cell] = comment;
        }
    }

    private void StartBackgroundLoad(SheetState state)
    {
        var remainingRows = state.TotalRows - state.LoadedRows;
        if (
            remainingRows <= 0
            || state.FilePath is null
            || state.LoadCts is { IsCancellationRequested: false }
        )
        {
            return;
        }

        var cts = new CancellationTokenSource();
        state.LoadCts = cts;
        var skipRows = state.LoadedRows;
        _ = Task.Run(() =>
        {
            try
            {
                foreach (
                    var rawRow in OoxmlLazyReader.ReadRows(
                        state.FilePath,
                        state.SheetName,
                        maxRows: remainingRows,
                        skipRows: skipRows
                    )
                )
                {
                    if (cts.IsCancellationRequested)
                    {
                        return;
                    }

                    Dispatcher.Invoke(() =>
                    {
                        if (cts.IsCancellationRequested)
                        {
                            return;
                        }

                        AddRawRow(state.Table, state.Comments, rawRow);
                        state.LoadedRows++;
                        if (CurrentSheetState == state && state.LoadedRows % 50 is 0)
                        {
                            StatusText.Text = $"已加载 {state.LoadedRows}/{state.TotalRows} 行";
                        }
                    });
                }

                Dispatcher.Invoke(() =>
                {
                    if (cts.IsCancellationRequested)
                    {
                        return;
                    }

                    state.LoadedRows = state.TotalRows;
                    if (CurrentSheetState == state)
                    {
                        StatusText.Text = $"已加载全部 {state.TotalRows} 行";
                    }
                });
            }
            catch (Exception exception)
            {
                if (cts.IsCancellationRequested)
                {
                    return;
                }

                Dispatcher.Invoke(() =>
                {
                    if (CurrentSheetState == state)
                    {
                        StatusText.Text = $"后台加载失败：{exception.Message}";
                    }
                });
            }
        });
    }

    /// <summary>
    /// 把 1-based 列序号转成 Excel 列名（1=A, 26=Z, 27=AA, 703=AAA）。
    /// </summary>
    private static string GetExcelColumnName(int col)
    {
        var name = string.Empty;
        while (col > 0)
        {
            var rem = (col - 1) % 26;
            name = (char)('A' + rem) + name;
            col = (col - 1) / 26;
        }

        return name;
    }

    private DataGrid BuildGrid(DataTable table, Dictionary<(int Row, int Col), string> commentMap)
    {
        var grid = new DataGrid
        {
            ItemsSource = table.DefaultView,
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            CanUserSortColumns = true,
            EnableRowVirtualization = true,
            EnableColumnVirtualization = true,
            SelectionUnit = DataGridSelectionUnit.Cell,
            HeadersVisibility = DataGridHeadersVisibility.All,
            GridLinesVisibility = DataGridGridLinesVisibility.All,
            HorizontalGridLinesBrush = GridLineBrush,
            VerticalGridLinesBrush = GridLineBrush,
            BorderBrush = GridLineBrush,
            BorderThickness = new Thickness(1),
        };

        // 行号列：用 RowHeaderTemplate 强制渲染 TextBlock，不依赖主题默认 RowHeader 可见性
        var rowHeaderTemplate = new DataTemplate();
        var factory = new FrameworkElementFactory(typeof(TextBlock));
        factory.SetBinding(
            TextBlock.TextProperty,
            new Binding("Header")
            {
                RelativeSource = new RelativeSource(
                    RelativeSourceMode.FindAncestor,
                    typeof(DataGridRow),
                    1
                ),
            }
        );
        factory.SetValue(TextBlock.ForegroundProperty, Brushes.White);
        factory.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
        factory.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        factory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
        factory.SetValue(TextBlock.MarginProperty, new Thickness(6, 0, 6, 0));
        rowHeaderTemplate.VisualTree = factory;
        grid.RowHeaderTemplate = rowHeaderTemplate;

        var rowStyle = new Style(typeof(DataGridRow));
        rowStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        grid.RowStyle = rowStyle;
        grid.RowHeaderWidth = 50;

        // 行头样式：显式设背景+前景，确保深色主题下可见
        var rowHeaderStyle = new Style(typeof(DataGridRowHeader));
        rowHeaderStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        rowHeaderStyle.Setters.Add(new Setter(Control.BackgroundProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowHeaderStyle.Setters.Add(
            new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center)
        );
        grid.RowHeaderStyle = rowHeaderStyle;

        // 数据列 + 列头筛选 TextBox
        BuildDataColumns(grid, table, withFilterBox: true);

        // 单元格样式：深色主题边框 + 脏数据高亮 + 备注提示
        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        cellStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        cellStyle.Setters.Add(
            new EventSetter(
                MouseEnterEvent,
                new MouseEventHandler(
                    (s, _) =>
                    {
                        if (s is not DataGridCell { DataContext: DataRowView view } cell)
                            return;
                        var rowIndex = table.Rows.IndexOf(view.Row);
                        var colIndex = cell.Column.DisplayIndex;
                        ToolTipService.SetToolTip(
                            cell,
                            commentMap.TryGetValue((rowIndex + 1, colIndex + 1), out var text)
                                ? text
                                : null
                        );
                    }
                )
            )
        );
        grid.CellStyle = cellStyle;

        // 编辑提交 → 撤销栈
        grid.CellEditEnding += (_, args) =>
        {
            if (
                args.EditAction != DataGridEditAction.Commit
                || args.Column is not DataGridBoundColumn bound
                || args.Row.Item is not DataRowView view
            )
                return;
            var rowIndex = table.Rows.IndexOf(view.Row);
            var colIndex = args.Column.DisplayIndex;
            var oldValue = view[colIndex];
            var newValue = (args.EditingElement as TextBox)?.Text ?? string.Empty;
            if (oldValue?.ToString() == newValue)
                return;
            var state = CurrentSheetState;
            if (state is null)
                return;
            state.UndoStack.Push(new CellEditRecord(rowIndex, colIndex, oldValue, newValue));
            state.RedoStack.Clear();
            MarkDirty(grid, view, colIndex);
            MarkCurrentFileDirty();
        };

        // 右键菜单：增删行列
        var ctxMenu = new ContextMenu();
        var miInsertRowBelow = new MenuItem { Header = "在下方插入行" };
        miInsertRowBelow.Click += (_, _) => InsertRowBelow(grid);
        var miDeleteRow = new MenuItem { Header = "删除当前行" };
        miDeleteRow.Click += (_, _) => DeleteCurrentRow(grid);
        var sep1 = new Separator();
        var miInsertColRight = new MenuItem { Header = "在右侧插入列" };
        miInsertColRight.Click += (_, _) => InsertColumnRight(grid);
        var miDeleteCol = new MenuItem { Header = "删除当前列" };
        miDeleteCol.Click += (_, _) => DeleteCurrentColumn(grid);
        ctxMenu.Items.Add(miInsertRowBelow);
        ctxMenu.Items.Add(miDeleteRow);
        ctxMenu.Items.Add(sep1);
        ctxMenu.Items.Add(miInsertColRight);
        ctxMenu.Items.Add(miDeleteCol);
        grid.ContextMenu = ctxMenu;

        // Ctrl+V 粘贴
        grid.PreviewKeyDown += (_, args) =>
        {
            if (args.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            {
                PasteFromClipboard(grid);
                args.Handled = true;
            }
        };

        return grid;
    }

    /// <summary>
    /// 构造数据列（DataGridTextColumn + 列头）。withFilterBox=true 时列头带筛选 TextBox（主 grid）；
    /// false 时只放纯列名 TextBlock（冻结行 grid，只读、不参与筛选）。
    /// </summary>
    private void BuildDataColumns(DataGrid grid, DataTable table, bool withFilterBox)
    {
        grid.Columns.Clear();
        for (var c = 0; c < table.Columns.Count; c++)
        {
            var colIndex = c;
            var columnName = table.Columns[c].ColumnName;

            FrameworkElement headerElement;
            if (withFilterBox)
            {
                var headerPanel = new StackPanel { Orientation = Orientation.Vertical };
                var headerText = new TextBlock
                {
                    Text = columnName,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                };
                var filterBox = new TextBox
                {
                    Tag = colIndex,
                    Width = double.NaN, // 撑满列头
                    MinWidth = 60,
                    Margin = new Thickness(1),
                    ToolTip = $"筛选 {columnName}",
                };
                filterBox.PreviewKeyDown += (_, args) =>
                {
                    if (args.Key is not Key.Enter)
                        return;

                    ApplyFilter(table, filterBox);
                    args.Handled = true;
                };
                headerPanel.Children.Add(headerText);
                headerPanel.Children.Add(filterBox);
                headerElement = headerPanel;
            }
            else
            {
                headerElement = new TextBlock
                {
                    Text = columnName,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(1),
                };
            }

            var column = new DataGridTextColumn
            {
                Header = headerElement,
                Binding = new Binding($"[{colIndex}]"),
                Width = new DataGridLength(160),
                IsReadOnly = !withFilterBox,
            };
            grid.Columns.Add(column);
        }
    }

    private static void ApplyFilter(DataTable table, TextBox filterBox)
    {
        if (filterBox.Tag is not int colIndex)
            return;
        var keyword = filterBox.Text.Replace("'", "''");
        var columnName = table.Columns[colIndex].ColumnName;
        try
        {
            table.DefaultView.RowFilter = string.IsNullOrWhiteSpace(keyword)
                ? string.Empty
                : $"[{columnName}] LIKE '%{keyword}%'";
        }
        catch
        {
            // RowFilter 语法异常时静默忽略，不阻塞输入
        }
    }

    private void OnClearFilterClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        foreach (var column in state.MainGrid.Columns)
        {
            if (column.Header is StackPanel panel)
            {
                foreach (var child in panel.Children.OfType<TextBox>())
                    child.Text = string.Empty;
            }
        }
        state.Table.DefaultView.RowFilter = string.Empty;
    }

    private void OnAddRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        state.Table.DefaultView.RowFilter = string.Empty;
        state.Table.Rows.Add(state.Table.NewRow());
        ClearUndoRedo(state);
        MarkCurrentFileDirty();
    }

    private void OnDeleteRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        DeleteCurrentRow(state.MainGrid);
    }

    /// <summary>
    /// 在当前选中行下方插入空行。
    /// </summary>
    private void InsertRowBelow(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var table = state.Table;
        var savedFilter = table.DefaultView.RowFilter;
        table.DefaultView.RowFilter = string.Empty;

        var insertAt = GetCurrentRowIndex(grid, table);
        if (insertAt < 0)
            insertAt = table.Rows.Count; // 没选中则追加到末尾
        else
            insertAt += 1; // 在下方

        var newRow = table.NewRow();
        table.Rows.InsertAt(newRow, insertAt);

        table.DefaultView.RowFilter = savedFilter;
        ClearUndoRedo(state);
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 删除当前选中行（支持多选）。
    /// </summary>
    private void DeleteCurrentRow(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var table = state.Table;
        var savedFilter = table.DefaultView.RowFilter;
        table.DefaultView.RowFilter = string.Empty;

        foreach (var item in grid.SelectedItems.OfType<DataRowView>().ToList())
        {
            item.Row.Delete();
        }
        table.AcceptChanges();

        table.DefaultView.RowFilter = savedFilter;
        ClearUndoRedo(state);
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 在当前选中列右侧插入空列。
    /// </summary>
    private void InsertColumnRight(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var table = state.Table;
        var insertAt = GetCurrentColumnIndex(grid);
        if (insertAt < 0)
            insertAt = table.Columns.Count;
        else
            insertAt += 1; // 在右侧

        // 新列名：取原最大列序号+1 对应的 Excel 列名
        var newColName = GetExcelColumnName(table.Columns.Count + 1);
        // 确保列名不重复
        while (table.Columns.Contains(newColName))
            newColName += "_";
        table.Columns.Add(newColName, typeof(string));
        // 调整列顺序
        table.Columns[newColName]!.SetOrdinal(insertAt);

        // DataGrid 需要重建列（因为 AutoGenerateColumns=false）
        RebuildGridColumns(grid, table);
        if (state.FrozenGrid is not null)
            RebuildFrozenColumns(state);

        ClearUndoRedo(state);
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 删除当前选中列。
    /// </summary>
    private void DeleteCurrentColumn(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var table = state.Table;
        var colIndex = GetCurrentColumnIndex(grid);
        if (colIndex < 0 || colIndex >= table.Columns.Count)
            return;

        table.Columns.RemoveAt(colIndex);
        RebuildGridColumns(grid, table);
        if (state.FrozenGrid is not null)
            RebuildFrozenColumns(state);

        ClearUndoRedo(state);
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 增删行列后清空撤销/重做栈，避免索引错位还原到错误位置。
    /// </summary>
    private static void ClearUndoRedo(SheetState state)
    {
        state.UndoStack.Clear();
        state.RedoStack.Clear();
    }

    /// <summary>
    /// 重建主 DataGrid 列（增删列后调用，因为 AutoGenerateColumns=false）。
    /// </summary>
    private void RebuildGridColumns(DataGrid grid, DataTable table)
    {
        BuildDataColumns(grid, table, withFilterBox: true);
        var hasFrozenRows = _sheets.Values.Any(state =>
            state.MainGrid == grid && state.FrozenRows > 0
        );
        SetFilterBoxesReadOnly(grid, hasFrozenRows);
    }

    /// <summary>
    /// 重建冻结行 grid 的列（镜像主 grid 的列结构，简单列头、只读）。
    /// </summary>
    private void RebuildFrozenColumns(SheetState state)
    {
        var fg = state.FrozenGrid;
        if (fg is null)
            return;
        BuildDataColumns(fg, state.Table, withFilterBox: false);
        // 同步宽 + 冻结列数
        SyncFrozenToMain(state);
    }

    /// <summary>
    /// 获取当前选中单元格的行索引（底层 DataTable 行号）。
    /// </summary>
    private static int GetCurrentRowIndex(DataGrid grid, DataTable table)
    {
        if (grid.CurrentCell.Item is DataRowView view)
            return table.Rows.IndexOf(view.Row);
        return -1;
    }

    /// <summary>
    /// 获取当前选中单元格的列索引。
    /// </summary>
    private static int GetCurrentColumnIndex(DataGrid grid)
    {
        return grid.CurrentCell.Column?.DisplayIndex ?? -1;
    }

    private async void OnSaveClick(object sender, RoutedEventArgs e) =>
        await SaveCurrentFileAsync();

    /// <summary>
    /// 保存当前 sheet 所属文件的所有 sheet。返回 true=成功（_dirtyFiles 已移除），
    /// false=失败/无可保存项（已弹错误框+写日志）。供保存按钮(OnSaveClick)和
    /// OnClosing(Yes 分支 await 后再 Close)、OnWorkbookCloseClick 共用。
    /// </summary>
    private async Task<bool> SaveCurrentFileAsync()
    {
        var curState = CurrentSheetState;
        if (curState is null)
            return false;
        var filePath = curState.FilePath;
        if (filePath is null)
            return false;
        var sw = Stopwatch.StartNew();

        // ponytail: 必须在 UI 线程把 DataTable 拷贝成纯数据快照（string[,]），
        // 不能把 DataTable/DataView 本身传进 Task.Run——它们正被 DataGrid 绑定消费，
        // 后台线程 touch RowFilter/Rows 会触发跨线程 ListChanged → WPF 崩溃。
        var snapshots = _sheets
            .Values.Where(s => s.FilePath == filePath)
            .Select(s =>
            {
                var table = s.Table;
                var savedFilter = table.DefaultView.RowFilter;
                table.DefaultView.RowFilter = string.Empty;
                var rows = table.Rows.Count;
                var cols = table.Columns.Count;
                var data = new string[rows, cols];
                for (var r = 0; r < rows; r++)
                for (var c = 0; c < cols; c++)
                    data[r, c] = table.Rows[r][c]?.ToString() ?? string.Empty;
                table.DefaultView.RowFilter = savedFilter;
                return (s.SheetName, Data: data, Rows: rows, Cols: cols);
            })
            .ToList();

        Tabs.IsEnabled = false;
        Cursor = Cursors.Wait;
        StatusText.Text = $"正在保存：{Path.GetFileName(filePath)}…";

        try
        {
            var (elapsedMs, error) = await Task.Run(() =>
            {
                try
                {
                    using var package = new ExcelPackage(new FileInfo(filePath));
                    foreach (var (sheetName, data, rows, cols) in snapshots)
                    {
                        var sheet = package.Workbook.Worksheets[sheetName];

                        var existingRows = sheet.Dimension?.End.Row ?? 0;
                        var existingCols = sheet.Dimension?.End.Column ?? 0;

                        if (rows < existingRows)
                            sheet.DeleteRow(rows + 1, existingRows - rows);

                        if (cols < existingCols)
                            sheet.DeleteColumn(cols + 1, existingCols - cols);

                        // 批量写入：用 range 一次 SetValue，比逐格 sheet.Cells[r,c].Value 快几十倍
                        if (rows > 0 && cols > 0)
                            sheet.Cells[1, 1, rows, cols].Value = data;
                    }

                    package.Save();
                    return (sw.ElapsedMilliseconds, (Exception?)null);
                }
                catch (Exception ex)
                {
                    return (0L, ex);
                }
            });

            if (error is not null)
                throw error;

            sw.Stop();
            _dirtyFiles.Remove(filePath);
            UpdateTitle();
            StatusText.Text = $"已保存：{Path.GetFileName(filePath)}（耗时 {elapsedMs} ms）";
            return true;
        }
        catch (Exception ex)
        {
            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "workspace",
                "xlsx-editor-save-error.log"
            );
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
                File.AppendAllText(
                    logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}\n\n"
                );
            }
            catch { }

            MessageBox.Show(
                this,
                $"保存失败：{ex.Message}\n\n详细信息已写入：{logPath}",
                "保存错误",
                MessageBoxButton.OK,
                MessageBoxImage.Error
            );
            return false;
        }
        finally
        {
            Tabs.IsEnabled = true;
            Cursor = Cursors.Arrow;
        }
    }

    private void OnUndoClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        if (state.UndoStack.Count == 0)
            return;
        var record = state.UndoStack.Pop();
        // DataRowState 索引是底层行号；filtered view 下需要先清筛选
        var savedFilter = state.Table.DefaultView.RowFilter;
        state.Table.DefaultView.RowFilter = string.Empty;
        var current = state.Table.Rows[record.Row][record.Col];
        state.Table.Rows[record.Row][record.Col] = record.OldValue;
        state.RedoStack.Push(
            new CellEditRecord(record.Row, record.Col, current, current?.ToString() ?? string.Empty)
        );
        state.Table.DefaultView.RowFilter = savedFilter;
        MarkCurrentFileDirty();
    }

    private void OnRedoClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        if (state.RedoStack.Count == 0)
            return;
        var record = state.RedoStack.Pop();
        var savedFilter = state.Table.DefaultView.RowFilter;
        state.Table.DefaultView.RowFilter = string.Empty;
        var current = state.Table.Rows[record.Row][record.Col];
        state.Table.Rows[record.Row][record.Col] = record.NewValue;
        state.UndoStack.Push(
            new CellEditRecord(record.Row, record.Col, current, current?.ToString() ?? string.Empty)
        );
        state.Table.DefaultView.RowFilter = savedFilter;
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 从剪贴板粘贴 Excel 多格数据（Tab 分分隔列，CRLF 分隔行）。
    /// </summary>
    private void OnPaste(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        PasteFromClipboard(state.MainGrid);
    }

    private void PasteFromClipboard(DataGrid grid)
    {
        if (grid.SelectedItem is null && grid.CurrentCell.Item is null)
            return;
        var startCell = grid.CurrentCell;
        if (startCell.Column is null || startCell.Item is not DataRowView rowView)
            return;
        var table = rowView.DataView.Table;
        if (table is null)
            return;
        // 用视图索引而非底层表索引：冻结行/筛选模式下连续粘贴进"可见的连续行"，正确且不越界。
        var startViewIndex = grid.Items.IndexOf(rowView);
        var startCol = startCell.Column.DisplayIndex;

        var text = Clipboard.GetText();
        if (string.IsNullOrEmpty(text))
            return;

        var lines = text.Split(["\r\n"], StringSplitOptions.RemoveEmptyEntries);
        for (var i = 0; i < lines.Length; i++)
        {
            var targetView = startViewIndex + i;
            if (targetView < 0 || targetView >= grid.Items.Count)
                break;
            if (grid.Items[targetView] is not DataRowView tv)
                continue;
            var cells = lines[i].Split('\t');
            for (var j = 0; j < cells.Length; j++)
            {
                var targetCol = startCol + j;
                if (targetCol >= table.Columns.Count)
                    break;
                tv[targetCol] = cells[j];
                // 粘贴的格子也要绿色高亮（和手动编辑一致）
                MarkDirty(grid, tv, targetCol);
            }
        }
        // 粘贴绕过 CellEditEnding，得手动标脏，否则关窗不提示保存=数据丢失
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 标脏单元格绿色。按 DataRowView 定位容器，冻结行/筛选/虚拟化下都正确（越界/不可见静默跳过）。
    /// </summary>
    private static void MarkDirty(DataGrid grid, DataRowView view, int col)
    {
        grid.ScrollIntoView(view);
        grid.Dispatcher.BeginInvoke(
            () =>
            {
                if (
                    grid.ItemContainerGenerator.ContainerFromItem(view)
                        is not DataGridRow rowContainer
                    || col < 0
                    || col >= grid.Columns.Count
                    || grid.Columns[col].GetCellContent(rowContainer)?.Parent
                        is not DataGridCell cell
                )
                {
                    return;
                }

                cell.Background = DirtyCellBrush;
                cell.InvalidateVisual();
            },
            System.Windows.Threading.DispatcherPriority.Loaded
        );
    }

    // ── 冻结窗格 ──────────────────────────────────────────────────────────

    private void OnFreezeColumnClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        var grid = state.MainGrid;
        var colIndex = GetCurrentColumnIndex(grid);
        if (colIndex < 0)
        {
            StatusText.Text = "请先选中一列再冻结";
            return;
        }
        // 冻结到当前列左侧所有列（含当前列）
        var n = colIndex + 1;
        grid.FrozenColumnCount = n;
        state.FrozenColumns = n;
        if (state.FrozenGrid is not null)
            state.FrozenGrid.FrozenColumnCount = n;
        SaveFreeze(state);
        StatusText.Text = $"已冻结前 {n} 列";
    }

    private void OnFreezeRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        var grid = state.MainGrid;
        var tableRow = GetCurrentRowIndex(grid, state.Table);
        if (tableRow < 0)
        {
            StatusText.Text = "请先选中一行再冻结";
            return;
        }
        // 冻结 row 0..当前行
        state.FrozenRows = tableRow + 1;
        ApplyFreezeLayout(state);
        SaveFreeze(state);
        StatusText.Text = $"已冻结前 {state.FrozenRows} 行；冻结行时筛选不可用";
    }

    private void OnUnfreezeClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        state.MainGrid.FrozenColumnCount = 0;
        state.FrozenColumns = 0;
        state.FrozenRows = 0;
        ApplyFreezeLayout(state); // 拆冻结行 grid，主 grid 回到全表 DefaultView
        if (state.FilePath is not null)
            FreezeConfig.ClearFreeze(Path.GetFileName(state.FilePath), state.SheetName);
        StatusText.Text = "已取消冻结";
    }

    /// <summary>
    /// 根据 state.FrozenRows 搭建/拆除双 DataGrid 布局。
    /// FrozenRows>0：建冻结行 grid（只读，关纵向滚，横滚由主 grid 驱动同步），主 grid 切到 LCV（row>=N）。
    /// FrozenRows=0：拆冻结行 grid，主 grid 回到全表 DefaultView（零回归，筛选走原生 RowFilter）。
    /// 列冻结（FrozenColumnCount）两 grid 始终一致，由 SyncFrozenToMain 兜底。
    /// </summary>
    private void ApplyFreezeLayout(SheetState state)
    {
        var panel = state.Panel;
        if (panel is null)
            return;
        var table = state.Table;
        var n = state.FrozenRows;
        if (n <= 0 || n >= table.Rows.Count)
        {
            // 拆冻结行 grid
            if (state.FrozenGrid is { } fg)
            {
                panel.Children.Remove(fg);
                state.FrozenGrid = null;
                state.FrozenScroll = null;
            }
            // 主 grid 回到全表（LCV→DefaultView）
            state.MainGrid.ItemsSource = table.DefaultView;
            SetFilterBoxesReadOnly(state.MainGrid, isReadOnly: false);
            return;
        }

        // 建/复用冻结行 grid
        if (state.FrozenGrid is null)
        {
            var fg = BuildFrozenGrid(state);
            panel.Children.Add(fg);
            Grid.SetRow(fg, 0);
            state.FrozenGrid = fg;
            WireFreezeSync(state);
        }

        // 两个 grid 各自的 LCV：按行号谓词切分（IndexOf 是 O(log n)，60k 行约 1M 次操作，可接受）。
        // 主 grid 继续包 table.DefaultView，故列头筛选 RowFilter 仍生效；冻结行 grid 包一个独立
        // 的干净 DataView，不受列筛选影响（冻结行始终可见，符合"冻结"语义）。
        state.MainGrid.ItemsSource = MakeMainView(table, n);
        state.FrozenGrid.ItemsSource = MakeFrozenView(table, n);
        SetFilterBoxesReadOnly(state.MainGrid, isReadOnly: true);
        SyncFrozenToMain(state);
    }

    private static void SetFilterBoxesReadOnly(DataGrid grid, bool isReadOnly)
    {
        foreach (
            var filterBox in grid.Columns.SelectMany(column =>
                column.Header is StackPanel panel
                    ? panel.Children.OfType<TextBox>()
                    : Enumerable.Empty<TextBox>()
            )
        )
        {
            filterBox.IsReadOnly = isReadOnly;
            filterBox.ToolTip = isReadOnly ? "冻结行时筛选不可用" : "输入筛选内容后按 Enter";
        }
    }

    private static ListCollectionView MakeMainView(DataTable table, int n)
    {
        var view = new ListCollectionView(table.DefaultView);
        view.Filter = obj => table.Rows.IndexOf(((DataRowView)obj).Row) >= n;
        return view;
    }

    private static ListCollectionView MakeFrozenView(DataTable table, int n)
    {
        var fresh = new DataView(table);
        var view = new ListCollectionView(fresh);
        view.Filter = obj => table.Rows.IndexOf(((DataRowView)obj).Row) < n;
        return view;
    }

    /// <summary>
    /// 构造冻结行 DataGrid：与主 grid 同列结构、只读、关纵向滚、隐藏横向滚（程序控制）。
    /// 行号头同主 grid（50 宽），保证横向 X 对齐。列头用纯 TextBlock（无筛选框，只读不筛）。
    /// </summary>
    private DataGrid BuildFrozenGrid(SheetState state)
    {
        var grid = new DataGrid
        {
            IsReadOnly = true,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            CanUserSortColumns = false,
            CanUserResizeColumns = false,
            CanUserReorderColumns = false,
            EnableRowVirtualization = false, // 仅 N 行，关虚拟化避免行号头错位
            SelectionUnit = DataGridSelectionUnit.Cell,
            HeadersVisibility = DataGridHeadersVisibility.All,
            GridLinesVisibility = DataGridGridLinesVisibility.All,
            HorizontalGridLinesBrush = GridLineBrush,
            VerticalGridLinesBrush = GridLineBrush,
            BorderBrush = GridLineBrush,
            BorderThickness = new Thickness(1, 0, 1, 1),
            HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
        };

        // 行号列模板/样式（与主 grid 一致，确保行头宽度对齐）
        var rowHeaderTemplate = new DataTemplate();
        var factory = new FrameworkElementFactory(typeof(TextBlock));
        factory.SetBinding(
            TextBlock.TextProperty,
            new Binding("Header")
            {
                RelativeSource = new RelativeSource(
                    RelativeSourceMode.FindAncestor,
                    typeof(DataGridRow),
                    1
                ),
            }
        );
        factory.SetValue(TextBlock.ForegroundProperty, Brushes.White);
        factory.SetValue(TextBlock.FontWeightProperty, FontWeights.Bold);
        factory.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
        factory.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
        factory.SetValue(TextBlock.MarginProperty, new Thickness(6, 0, 6, 0));
        rowHeaderTemplate.VisualTree = factory;
        grid.RowHeaderTemplate = rowHeaderTemplate;

        var rowStyle = new Style(typeof(DataGridRow));
        rowStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        grid.RowStyle = rowStyle;
        grid.RowHeaderWidth = 50;

        var rowHeaderStyle = new Style(typeof(DataGridRowHeader));
        rowHeaderStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        rowHeaderStyle.Setters.Add(new Setter(Control.BackgroundProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowHeaderStyle.Setters.Add(
            new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center)
        );
        grid.RowHeaderStyle = rowHeaderStyle;

        // 列：简单列头、只读，宽度/冻结列数由 SyncFrozenToMain 镜像主 grid
        BuildDataColumns(grid, state.Table, withFilterBox: false);
        // 冻结行 grid 的行号 = 视图行号+1（它显示的是 row 0..N-1，即真实行号 1..N，无偏移）
        grid.LoadingRow += (_, args) => args.Row.Header = args.Row.GetIndex() + 1;
        return grid;
    }

    /// <summary>
    /// 挂三同步：横向滚（主→冻）、列宽（主→冻）、冻结列数（主→冻）。全部走主 grid 的 LayoutUpdated，
    /// 一个 handler 兜底，免逐列 DependencyPropertyDescriptor 的悬挂引用。
    /// </summary>
    private void WireFreezeSync(SheetState state)
    {
        var main = state.MainGrid;
        // 主→冻 横向滚：ScrollViewer.ScrollChanged 是路由事件，从主 grid 内部 ScrollViewer 冒泡。
        main.AddHandler(
            ScrollViewer.ScrollChangedEvent,
            new ScrollChangedEventHandler(
                (_, e) =>
                {
                    if (state.FrozenGrid is not { } fg)
                        return;
                    state.FrozenScroll ??= FindScrollViewer(fg);
                    state.FrozenScroll?.ScrollToHorizontalOffset(e.HorizontalOffset);
                }
            )
        );

        // 列宽 + 冻结列数 + 列数（增删列）镜像：LayoutUpdated 高频但工作极轻（逐列比 width）。
        // 自稳定——镜像完即相等，下一轮不再写。
        main.LayoutUpdated += (_, _) =>
        {
            if (state.FrozenGrid is not { } fg)
                return;
            if (fg.Columns.Count != main.Columns.Count)
                RebuildFrozenColumns(state);
            SyncFrozenToMain(state);
        };
    }

    /// <summary>
    /// 把主 grid 的列宽 + FrozenColumnCount 镜像到冻结行 grid（调用方保证列数一致）。
    /// </summary>
    private static void SyncFrozenToMain(SheetState state)
    {
        var main = state.MainGrid;
        if (state.FrozenGrid is not { } fg)
            return;
        for (var i = 0; i < main.Columns.Count && i < fg.Columns.Count; i++)
        {
            if (fg.Columns[i].Width != main.Columns[i].Width)
                fg.Columns[i].Width = main.Columns[i].Width;
        }
        if (fg.FrozenColumnCount != main.FrozenColumnCount)
            fg.FrozenColumnCount = main.FrozenColumnCount;
    }

    /// <summary>
    /// 在 DataGrid 的可视树里找 ScrollViewer（冻结行 grid 的横滚驱动用）。深递归但只调一次（结果缓存到 state.FrozenScroll）。
    /// </summary>
    private static ScrollViewer? FindScrollViewer(DependencyObject d)
    {
        if (d is ScrollViewer sv)
            return sv;
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(d); i++)
        {
            var found = FindScrollViewer(VisualTreeHelper.GetChild(d, i));
            if (found is not null)
                return found;
        }
        return null;
    }

    private static void SaveFreeze(SheetState state)
    {
        if (state.FilePath is null)
            return;
        FreezeConfig.SetFreeze(
            Path.GetFileName(state.FilePath),
            state.SheetName,
            state.FrozenColumns,
            state.FrozenRows
        );
    }

    // ── 聚光灯 ────────────────────────────────────────────────────────────

    private DataGrid? _spotlightGrid;

    private void OnSpotlightToggleChecked(object sender, RoutedEventArgs e)
    {
        ApplySpotlightToCurrentGrid();
    }

    private void OnSpotlightToggleUnchecked(object sender, RoutedEventArgs e)
    {
        ClearSpotlight();
        _spotlightGrid = null;
    }

    /// <summary>
    /// 切到当前选中 sheet 的 DataGrid，挂事件，立即高亮。
    /// </summary>
    private void ApplySpotlightToCurrentGrid()
    {
        // 先从旧 grid 卸载事件
        if (_spotlightGrid is not null)
            _spotlightGrid.SelectedCellsChanged -= OnSpotlightSelectionChanged;
        ClearSpotlight();

        _spotlightGrid = CurrentMainGrid;
        if (_spotlightGrid is not null)
        {
            _spotlightGrid.SelectedCellsChanged += OnSpotlightSelectionChanged;
            ApplySpotlight();
        }
    }

    private void OnSpotlightSelectionChanged(object sender, SelectedCellsChangedEventArgs e)
    {
        ApplySpotlight();
    }

    private void ApplySpotlight()
    {
        if (_spotlightGrid is null || SpotlightToggle.IsChecked is not true)
            return;
        ClearSpotlight();

        var selectedRows = new HashSet<int>();
        var selectedColumns = new HashSet<int>();
        foreach (var selectedCell in _spotlightGrid.SelectedCells)
        {
            var rowIndex = _spotlightGrid.Items.IndexOf(selectedCell.Item);
            if (rowIndex >= 0)
                selectedRows.Add(rowIndex);
            if (selectedCell.Column is not null)
                selectedColumns.Add(selectedCell.Column.DisplayIndex);
        }

        if (selectedRows.Count is 0 || selectedColumns.Count is 0)
            return;

        HighlightSpotlight(selectedRows, selectedColumns);
    }

    /// <summary>
    /// 高亮选中行 + 选中列（只处理可见行，虚拟化安全）。
    /// </summary>
    private void HighlightSpotlight(
        IReadOnlySet<int> selectedRows,
        IReadOnlySet<int> selectedColumns
    )
    {
        if (_spotlightGrid is null)
            return;

        for (var i = 0; i < _spotlightGrid.Items.Count; i++)
        {
            if (
                _spotlightGrid.ItemContainerGenerator.ContainerFromIndex(i)
                is not DataGridRow rowContainer
            )
                continue; // 虚拟化下不可见行返回 null，跳

            foreach (var col in _spotlightGrid.Columns)
            {
                if (col.GetCellContent(rowContainer)?.Parent is not DataGridCell cell)
                    continue;
                if (
                    (selectedRows.Contains(i) || selectedColumns.Contains(col.DisplayIndex))
                    && cell.Background != DirtyCellBrush
                )
                    cell.Background = SpotlightBrush;
            }
        }
    }

    private void ClearSpotlight()
    {
        if (_spotlightGrid is null)
            return;
        for (var i = 0; i < _spotlightGrid.Items.Count; i++)
        {
            if (
                _spotlightGrid.ItemContainerGenerator.ContainerFromIndex(i)
                is not DataGridRow rowContainer
            )
                continue;
            foreach (var col in _spotlightGrid.Columns)
            {
                if (
                    col.GetCellContent(rowContainer)?.Parent is DataGridCell cell
                    && cell.Background == SpotlightBrush
                )
                    cell.ClearValue(DataGridCell.BackgroundProperty);
            }
        }
    }

    // ── 查看完整值 ──────────────────────────────────────────────────────

    private void OnViewFullValueClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        var grid = state.MainGrid;
        if (grid.CurrentCell.Column is null || grid.CurrentCell.Item is not DataRowView view)
            return;
        var colIndex = grid.CurrentCell.Column.DisplayIndex;
        var value = view[colIndex]?.ToString() ?? string.Empty;
        var header = grid.CurrentCell.Column.Header is StackPanel panel
            ? panel.Children.OfType<TextBlock>().FirstOrDefault()?.Text ?? "?"
            : "?";
        var rowIdx = grid.Items.IndexOf(view);
        var win = new Window
        {
            Title = $"完整值（可编辑，关窗写回）— {header} 行 {rowIdx + 1}",
            Width = 600,
            Height = 400,
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
        };
        var tb = new TextBox
        {
            Text = value,
            IsReadOnly = false, // 可编辑；关窗时若改了写回单元格 + 标脏
            AcceptsReturn = true,
            TextWrapping = TextWrapping.Wrap,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            FontFamily = new FontFamily("Consolas"),
            FontSize = 14,
        };
        win.Content = tb;
        win.Closed += (_, _) =>
        {
            if (tb.Text != value)
            {
                view[colIndex] = tb.Text;
                MarkCurrentFileDirty();
            }
        };
        win.Show();
    }

    // ── 搜索 ─────────────────────────────────────────────────────────────

    private void OnSearchBoxKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            OnSearchClick(sender, e);
            e.Handled = true;
        }
    }

    private async void OnSearchClick(object sender, RoutedEventArgs e)
    {
        var keyword = SearchBox.Text.Trim();
        if (string.IsNullOrEmpty(keyword))
        {
            SearchResults.ItemsSource = null;
            SearchResults.Visibility = Visibility.Collapsed;
            SearchResultText.Text = "0 条结果";
            return;
        }

        var scope = ((ComboBoxItem)SearchScope.SelectedItem).Content?.ToString();

        // ponytail: 搜索必须在后台线程，不能在 UI 线程遍历 6 万行 DataTable。
        // 且不能把 DataTable 传进 Task.Run——它是 UI 绑定对象。先在 UI 线程拷成快照。
        var targets = new List<(string FileName, string SheetName, string[,] Data)>();
        switch (scope)
        {
            case "当前 Sheet":
                if (CurrentSheetState is { } cur)
                    targets.Add(BuildSearchSnapshot(cur));
                break;
            case "当前工作簿":
                // 当前工作簿的所有 sheet（不是全部打开文件）
                if (
                    CurrentWorkbookTab is TabItem wb
                    && _workbookSheets.TryGetValue(wb, out var sheetList)
                )
                    foreach (var t in sheetList)
                        if (_sheets.TryGetValue(t, out var s))
                            targets.Add(BuildSearchSnapshot(s));
                break;
            case "所有工作簿":
                foreach (var state in _sheets.Values)
                    targets.Add(BuildSearchSnapshot(state));
                break;
        }

        if (targets.Count == 0)
        {
            SearchResultText.Text = "0 条结果";
            return;
        }

        SearchResultText.Text = "搜索中…";
        var kw = keyword;
        var results = await Task.Run(() => SearchSnapshots(targets, kw));

        SearchResults.ItemsSource = results;
        SearchResults.DisplayMemberPath = nameof(SearchResultItem.Display);
        SearchResults.Visibility = results.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        SearchResultText.Text = $"{results.Count} 条结果";
        if (results.Count > 0)
            SearchResults.SelectedIndex = 0;
    }

    /// <summary>
    /// 在 UI 线程把 DataTable 拷成 string[,] 快照，供后台搜索。
    /// </summary>
    private (string, string, string[,]) BuildSearchSnapshot(SheetState state)
    {
        var table = state.Table;
        var savedFilter = table.DefaultView.RowFilter;
        table.DefaultView.RowFilter = string.Empty;
        var rows = table.Rows.Count;
        var cols = table.Columns.Count;
        var data = new string[rows, cols];
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < cols; c++)
            data[r, c] = table.Rows[r][c]?.ToString() ?? string.Empty;
        table.DefaultView.RowFilter = savedFilter;
        var fileName = state.FilePath is not null ? Path.GetFileName(state.FilePath) : "?";
        return (fileName, state.SheetName, data);
    }

    private static List<SearchResultItem> SearchSnapshots(
        List<(string FileName, string SheetName, string[,] Data)> targets,
        string keyword
    )
    {
        var results = new List<SearchResultItem>();
        foreach (var (fileName, sheetName, data) in targets)
        {
            var rows = data.GetLength(0);
            var cols = data.GetLength(1);
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < cols; c++)
                {
                    var val = data[r, c];
                    if (val.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        var colName = GetExcelColumnName(c + 1);
                        results.Add(
                            new SearchResultItem(fileName, sheetName, r + 1, c + 1, val, colName)
                        );
                        if (results.Count >= 500)
                            return results;
                    }
                }
            }
        }

        return results;
    }

    private void OnSearchResultSelected(object sender, SelectionChangedEventArgs e)
    {
        if (SearchResults.SelectedItem is not SearchResultItem item)
            return;
        // 按文件名找对应工作簿 tab（同名足够，路径通常唯一）
        var wbPair = _workbookByPath.FirstOrDefault(kv =>
            Path.GetFileName(kv.Key) == item.FileName
        );
        if (wbPair.Value is not TabItem wbTab)
            return;
        Tabs.SelectedItem = wbTab;
        if (GetSheetTabs(wbTab) is not { } sheetTabs)
            return;
        // 切到对应 sheet
        var sheetTab = sheetTabs
            .Items.OfType<TabItem>()
            .FirstOrDefault(t => t.Header?.ToString() == item.SheetName);
        if (sheetTab is null)
            return;
        sheetTabs.SelectedItem = sheetTab;
        if (!_sheets.TryGetValue(sheetTab, out var state))
            return;
        var grid = state.MainGrid;
        // 滚动到行并选中（item.Row 是 1-based 真实行号；冻结行模式下主 grid 视图索引 = 行号 - FrozenRows - 1）
        var viewIndex = item.Row - 1 - state.FrozenRows;
        if (viewIndex < 0)
            return; // 该行在冻结行区，主 grid 里选不到，跳过
        if (viewIndex < grid.Items.Count)
        {
            grid.Focus();
            grid.ScrollIntoView(grid.Items[viewIndex]);
            grid.SelectedIndex = viewIndex;
        }
    }

    private record SearchResultItem(
        string FileName,
        string SheetName,
        int Row,
        int Col,
        string Value,
        string ColumnName
    )
    {
        public string Display =>
            $"[{SheetName}] {ColumnName}{Row}: {Value[..Math.Min(60, Value.Length)]}…";
    }

    private void UpdateTitle()
    {
        var filePath = CurrentFilePath;
        var isDirty = filePath is not null && _dirtyFiles.Contains(filePath);
        var dirtyMark = isDirty ? " *" : string.Empty;
        Title = filePath is null
            ? "xlsx 轻量编辑器"
            : $"{Path.GetFileName(filePath)}{dirtyMark} - xlsx 轻量编辑器";
    }

    /// <summary>
    /// 标记当前 sheet 所属文件为脏（有未保存改动）。
    /// </summary>
    private void MarkCurrentFileDirty()
    {
        var fp = CurrentFilePath;
        if (fp is not null)
            _dirtyFiles.Add(fp);
        UpdateTitle();
    }

    /// <summary>
    /// tab 切换（工作簿级 + sheet 级，SelectionChanged 冒泡）时更新标题 + 重新应用聚光灯。
    /// </summary>
    private void OnTabsSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var selectedState = CurrentSheetState;
        if (_activeSheetState is not null && _activeSheetState != selectedState)
        {
            _activeSheetState.LoadCts?.Cancel();
        }

        _activeSheetState = selectedState;
        UpdateTitle();
        if (SpotlightToggle.IsChecked == true)
            ApplySpotlightToCurrentGrid();

        if (selectedState is not null)
        {
            StatusText.Text =
                selectedState.LoadedRows >= selectedState.TotalRows
                    ? $"已加载全部 {selectedState.TotalRows} 行"
                    : $"已加载 {selectedState.LoadedRows}/{selectedState.TotalRows} 行";
            if (
                selectedState.LoadedRows < selectedState.TotalRows
                && selectedState.LoadCts is not { IsCancellationRequested: false }
            )
            {
                StartBackgroundLoad(selectedState);
            }
        }
    }

    private async void OnClosing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        foreach (var state in _sheets.Values)
        {
            state.LoadCts?.Cancel();
        }

        if (_dirtyFiles.Count == 0)
            return;
        e.Cancel = true; // 取消本次关闭，等用户确认后决定
        var files = string.Join(", ", _dirtyFiles.Select(Path.GetFileName));
        var result = MessageBox.Show(
            this,
            $"以下文件有未保存的更改：\n{files}\n\n是否保存当前文件？",
            "未保存的更改",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question
        );
        switch (result)
        {
            // Yes: await 真正保存（旧版 async-void 不 await 致竞态/2-close），
            // 成功后再 Close()——await 之后已离开 closing 阶段，安全（触发新 OnClosing，
            // _dirtyFiles 空 → 放行）。不用 e.Cancel=false 是因为原始关闭早已被取消，
            // 只能靠重新 Close() 起一次。
            case MessageBoxResult.Yes:
            {
                var ok = await SaveCurrentFileAsync();
                if (!ok || _dirtyFiles.Count > 0)
                    return; // 保存失败，不关
                Close();
                break;
            }
            // No: 不调 Close()（那是崩溃根因——closing 期间 Close → VerifyNotClosing 抛），
            // 设 e.Cancel=false 让原始关闭走
            case MessageBoxResult.No:
                _dirtyFiles.Clear();
                e.Cancel = false;
                break;
            case MessageBoxResult.Cancel:
            default:
                break; // e.Cancel 保持 true，不关
        }
    }

    private sealed class SheetState(
        string sheetName,
        DataTable table,
        Dictionary<(int Row, int Col), string> comments,
        string? filePath,
        int totalRows,
        int loadedRows
    )
    {
        public string SheetName { get; } = sheetName;
        public DataTable Table { get; } = table;
        public Dictionary<(int Row, int Col), string> Comments { get; } = comments;
        public string? FilePath { get; set; } = filePath;
        public int TotalRows { get; } = totalRows;
        public int LoadedRows { get; set; } = loadedRows;
        public CancellationTokenSource? LoadCts { get; set; }
        public Stack<CellEditRecord> UndoStack { get; } = new();
        public Stack<CellEditRecord> RedoStack { get; } = new();

        // ── 冻结窗格 ──
        public DataGrid MainGrid { get; set; } = null!;
        public Grid? Panel { get; set; }
        public DataGrid? FrozenGrid { get; set; }
        public ScrollViewer? FrozenScroll { get; set; }
        public int FrozenRows { get; set; }
        public int FrozenColumns { get; set; }
    }
}
