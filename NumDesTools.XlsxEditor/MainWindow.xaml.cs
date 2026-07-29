using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using DataGridExtensions;
using Microsoft.Win32;
using OfficeOpenXml;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// WPF + MahApps.Metro 版本。两级 tab：工作簿（带 ✕ 关闭）→ sheet（DataGrid）。
/// 数据层后台线程 + range 批量枚举懒加载；UI 层 WPF DataGrid 虚拟化 + MetroWindow 深色主题。
/// 冻结窗格：列 = 原生 FrozenColumnCount；行 = 双 DataGrid（顶部冻结行 grid + 主 grid，
/// 共享同 DataTable 的两个 ListCollectionView 按行号谓词切分，横向滚 + 列宽 + 冻列数三同步）。
/// </summary>
public partial class MainWindow : Wpf.Ui.Controls.FluentWindow
{
    private static readonly Brush DirtyCellBrush = new SolidColorBrush(Color.FromRgb(43, 145, 76));

    // 深色主题单元格边框颜色（比背景略亮）
    private static readonly Brush GridLineBrush = new SolidColorBrush(Color.FromRgb(60, 60, 60));

    // #P8-1：列头（Excel 字母坐标 A/B/C 所在行）背景色——比数据区略深，与深色主题一致。
    private static readonly Brush HeaderBackgroundBrush = new SolidColorBrush(
        Color.FromRgb(45, 45, 45)
    );

    // 聚光灯：选中区域外框（亮黄色），行列用半透明背景色指示
    private static readonly Brush SpotlightBorderBrush = new SolidColorBrush(
        Color.FromRgb(255, 220, 50)
    );
    private static readonly Thickness SpotlightBorderThickness = new(2);
    private static readonly Brush SpotlightRowColBrush = new SolidColorBrush(
        Color.FromArgb(40, 0, 120, 215)
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
        ApplyWorkbookTabStyle();
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

    /// <summary>
    /// 焦点是否在文本输入控件里（筛选框 / 单元格编辑器 / 搜索框）。此时窗口级快捷键不应劫持
    /// 文本编辑键（Delete/Backspace 等），否则用户在筛选框里删不了字（#7）。
    /// 纯静态、可单测：只看焦点元素类型。
    /// </summary>
    internal static bool IsTextInputFocused(IInputElement? focused) =>
        focused is TextBox or System.Windows.Controls.Primitives.TextBoxBase;

    /// <summary>
    /// 行号头文本（1-based 绝对 Excel 行号）。#4：优先用 <see cref="RowView.RowIndex"/>（真实 store 行号），
    /// 冻结/主区/排序/筛选下都稳定对齐；拿不到 RowView 时回退到 grid 内的相对索引 +1。纯静态、可单测。
    /// </summary>
    internal static int RowHeaderNumber(object? rowItem, int fallbackIndex) =>
        rowItem is RowView view ? view.RowIndex + 1 : fallbackIndex + 1;

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        var editingText = IsTextInputFocused(Keyboard.FocusedElement);

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
                case Key.Z when !editingText:
                    OnUndoClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.Y when !editingText:
                    OnRedoClick(sender, e);
                    e.Handled = true;
                    break;
                case Key.V when !editingText:
                    OnPaste(sender, e);
                    e.Handled = true;
                    break;
            }
        }
        else if (e.Key == Key.Escape)
        {
            Close();
        }
        else if (e.Key is Key.Back or Key.Delete && !editingText)
        {
            // #7：焦点在筛选框/单元格编辑器/搜索框内时，不劫持 Delete/Backspace，
            // 让 TextBox 正常删除字符。仅在焦点不在文本输入时才走"清空选中单元格"。
            DeleteSelectedCells();
            e.Handled = true;
        }
    }

    /// <summary>
    /// 删除所有选中单元格的内容（Backspace/Delete）。
    /// </summary>
    private void DeleteSelectedCells()
    {
        if (CurrentSheetState is not { } state)
            return;
        var grid = state.MainGrid;
        if (grid.SelectedCells.Count == 0)
            return;
        // #6：多格删除作为一个复合撤销单元（一次 Ctrl+Z 整体撤销这次删除）。
        var batch = new List<CellEditRecord>();
        foreach (var cell in grid.SelectedCells)
        {
            if (cell.Item is not RowView view || cell.Column is null)
                continue;
            var colIndex = cell.Column.DisplayIndex;
            var rowIndex = view.RowIndex;
            var oldValue = view[colIndex];
            batch.Add(new CellEditRecord(rowIndex, colIndex, oldValue, string.Empty));
            view[colIndex] = string.Empty; // RowView 索引器 → SetCell 标脏
            MarkDirty(grid, view, colIndex);
        }
        if (batch.Count > 0)
        {
            state.UndoStack.Push(new CellBatchAction(batch));
            state.RedoStack.Clear();
        }
        MarkCurrentFileDirty();
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
        // P14：文件名带 #（如 #【A-LTE】配置模版【cent】.xlsx）的文件拒绝打开——这类文件
        // 通常含隐藏 sheet / 外链表，Sylvan 枚举不到会崩溃。给友好提示而不是抛未捕获异常。
        var fileNameForCheck = Path.GetFileName(path);
        if (fileNameForCheck.StartsWith('#'))
        {
            StatusText.Text = $"跳过：{fileNameForCheck}（文件名以 # 开头，暂不支持打开）";
            return;
        }

        Tabs.IsEnabled = false;
        Cursor = Cursors.Wait;
        StatusText.Text = $"正在加载：{Path.GetFileName(path)}…";

        var sw = Stopwatch.StartNew();
        // P3.2: 一次性流式构建 ColumnStore（弃用旧的"首屏 200 行 + 后台逐行 Dispatcher.Invoke 灌 DataTable"路径）。
        var built = await Task.Run(() => BuildStoresFromExcel(path));

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
        foreach (var (name, store, comments, totalRows) in built)
        {
            // 双 DataGrid 布局：外层 Grid(panel) 含两行——row0=冻结行 grid（按需出现），row1=主 grid。
            // 主 grid 永远在 row1，冻结行 grid 由 ApplyFreezeLayout 在冻结行时插入 row0。
            var grid = BuildGrid(store, comments);
            var panel = new Grid();
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 冻结行槽（空时 Auto=0 高度）
            panel.RowDefinitions.Add(
                new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }
            );
            panel.Children.Add(grid);
            Grid.SetRow(grid, 1);

            var sheetTab = new TabItem
            {
                Header = new TextBlock { Text = name, FontSize = 11 },
                Content = panel,
            };
            var state = new SheetState(name, store, comments, path, totalRows)
            {
                MainGrid = grid,
                Panel = panel,
            };
            grid.ItemsSource = state.View;
            // P5：把 DataGridExtensions 的筛选路由到 ColumnStore（ICustomFilter 适配器设为 grid.DataContext）。
            // 列类型一次性采样缓存，避免每次筛选变化都重新 DetectColumnType。
            var typeCache = new Dictionary<int, ColumnType>();
            ColumnType TypeResolver(int col)
            {
                if (!typeCache.TryGetValue(col, out var t))
                {
                    t = DetectColumnType(store, col);
                    typeCache[col] = t;
                }

                return t;
            }

            var stateForStatus = state;
            var filterAdapter = new ColumnStoreFilterAdapter(
                store,
                state.View,
                TypeResolver,
                filteredCount =>
                {
                    if (CurrentSheetState == stateForStatus)
                    {
                        StatusText.Text =
                            filteredCount == store.RowCount
                                ? $"已加载全部 {store.RowCount} 行"
                                : $"筛选后 {filteredCount}/{store.RowCount} 行";
                    }
                }
            );
            state.Filter = filterAdapter;
            grid.DataContext = filterAdapter;
            // 行号头（#4）：直接用 RowView 的绝对 store 行号 +1，冻结/非冻结/排序/筛选下都对齐、连续。
            // 冻结区 RowRangeView(0,n) → 行号 1..n；主区 RowRangeView(n,..) → 行号 n+1..，无缝衔接。
            grid.LoadingRow += (_, args) =>
                args.Row.Header = RowHeaderNumber(args.Row.Item, args.Row.GetIndex());
            sheetTabs.Items.Add(sheetTab);
            _sheets[sheetTab] = state;
            _workbookSheets[wbTab].Add(sheetTab);
            firstNewTab ??= sheetTab;

            // 应用冻结配置：列冻结（原生 FrozenColumnCount）+ 行冻结（P4 重做，RowRangeView 双 grid）
            var (fc, fr) = FreezeConfig.GetFreeze(fileName, name);
            if (fc > 0 && fc <= grid.Columns.Count)
            {
                grid.FrozenColumnCount = fc;
                state.FrozenColumns = fc;
                ApplyFreezeColumnDivider(state, fc);
            }
            if (fr > 0 && fr < store.RowCount)
            {
                state.FrozenRows = fr;
                ApplyFreezeRows(state);
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
                $"已加载全部 {currentState.TotalRows} 行（{currentState.SheetName}），耗时 {sw.ElapsedMilliseconds} ms";
        }
    }

    /// <summary>
    /// 建工作簿 tab：Header = 文件名 + ✕ 关闭按钮，Content = 内嵌 sheet TabControl。
    /// </summary>
    private TabItem BuildWorkbookTab(string fileName, TabControl sheetTabs)
    {
        sheetTabs.Padding = new Thickness(4, 0, 4, 2);
        sheetTabs.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
        var sheetTabStyle = new Style(
            typeof(TabItem),
            Application.Current.TryFindResource(typeof(TabItem)) as Style
        );
        sheetTabStyle.Setters.Add(new Setter(Control.FontSizeProperty, 11.0));
        sheetTabStyle.Setters.Add(new Setter(TextBlock.FontSizeProperty, 11.0));
        // P14：选中 tab 高亮——WPF-UI 模板触发器用 DynamicResource 盖 Background 导致颜色不变，
        // 改用 Foreground（走 TextElement 继承到 TextBlock，模板盖不住）+ FontWeight 做视觉区分。
        sheetTabStyle.Setters.Add(
            new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(45, 45, 45)))
        );
        sheetTabStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.Gray));
        sheetTabStyle.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Normal));
        sheetTabStyle.Triggers.Add(
            new Trigger
            {
                Property = Selector.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(80, 80, 80))),
                    new Setter(Control.ForegroundProperty, Brushes.White),
                    new Setter(Control.FontWeightProperty, FontWeights.Bold),
                    new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(86, 156, 214))),
                    new Setter(Control.BorderThicknessProperty, new Thickness(3, 0, 0, 0)),
                },
            }
        );
        sheetTabs.ItemContainerStyle = sheetTabStyle;
        // 文件 tab 和 sheet tab 之间用粗亮线分隔（Excel 风格的凹凸感）
        var sheetTabsBorder = new Border
        {
            Margin = new Thickness(0, 10, 0, 0),
            Padding = new Thickness(0),
            BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120)),
            BorderThickness = new Thickness(0, 6, 0, 2),
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
    /// 给外层工作簿 TabControl 设选中高亮样式（与 sheet tab 同样的视觉风格）。
    /// </summary>
    private void ApplyWorkbookTabStyle()
    {
        var wbTabStyle = new Style(
            typeof(TabItem),
            Application.Current.TryFindResource(typeof(TabItem)) as Style
        );
        // P14：同 sheet tab，用 Foreground+FontWeight 做选中态视觉区分。
        wbTabStyle.Setters.Add(
            new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(50, 50, 50)))
        );
        wbTabStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.Gray));
        wbTabStyle.Triggers.Add(
            new Trigger
            {
                Property = Selector.IsSelectedProperty,
                Value = true,
                Setters =
                {
                    new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(90, 90, 90))),
                    new Setter(Control.ForegroundProperty, Brushes.White),
                    new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(86, 156, 214))),
                    new Setter(Control.BorderThicknessProperty, new Thickness(3, 0, 0, 0)),
                },
            }
        );
        Tabs.ItemContainerStyle = wbTabStyle;
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

    /// <summary>
    /// P3.2 新加载路径：一次性流式把每个非 <c>#</c> 前缀、非空 sheet 读进 <see cref="ColumnStore"/>。
    /// 用 <see cref="ColumnStoreExcelLoader.Load"/>（Sylvan 单趟前向扫描），不再首屏 200 行 + 后台逐行
    /// Dispatcher.Invoke 灌 DataTable。批注（comment）当前不随流式加载读取（Sylvan 只读值），
    /// 故 comments 为空字典——批注提示是已知降级，见 status.md。
    /// </summary>
    private static List<(
        string Name,
        ColumnStore Store,
        Dictionary<(int Row, int Col), string> Comments,
        int TotalRows
    )> BuildStoresFromExcel(string path)
    {
        var result = new List<(string, ColumnStore, Dictionary<(int Row, int Col), string>, int)>();
        foreach (var sheetName in OoxmlLazyReader.ReadSheetNames(path))
        {
            if (sheetName.StartsWith('#'))
            {
                continue;
            }

            var store = ColumnStoreExcelLoader.Load(path, sheetName);
            if (store.RowCount is 0 || store.ColumnCount is 0)
            {
                continue;
            }

            result.Add(
                (sheetName, store, new Dictionary<(int Row, int Col), string>(), store.RowCount)
            );
        }

        return result;
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

    /// <summary>
    /// #P8-1：列头样式——强制 Excel 字母坐标（A/B/C，来自 <see cref="ColumnStore.ColumnNames"/>）在深色主题下
    /// 以白色粗体可见。根因：DataGridExtensions 的列头模板（ColumnHeaderTemplateKey）用一个
    /// <c>ContentPresenter Content="{Binding}"</c> 呈现列头对象（即字母字符串），其前景色继承自
    /// <see cref="DataGridColumnHeader"/>；而 WPF-UI 深色皮肤未给该 ContentPresenter 明确前景，
    /// 导致字母以近乎不可见的颜色绘制（UIA 仍能读到 HeaderItem.Name="A"，但视觉上看不见——这正是
    /// "UIA 读到字母 vs 截图看不到字母"两组观察的调和点：内容模型有字母，只是没被painted出来）。
    /// 显式设 Foreground=White + Background=深色 + 足够行高，字母才真正可见。两个 grid（主 + 冻结）
    /// 共用同一套样式，保证冻结态/非冻结态下字母都可见。
    /// </summary>
    private static Style BuildColumnHeaderStyle()
    {
        var style = new Style(typeof(DataGridColumnHeader));
        style.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        // #P9-1：DGX 列头模板里的 PART_Content 是 ContentPresenter，字母（"A"/"B"...）以 TextBlock 呈现；
        // 其颜色靠 TextElement.Foreground 继承。WPF-UI 深色皮肤的 DataGridColumnHeader 模板可能自带内层
        // ContentPresenter 前景，盖过 Control.Foreground（截图证实冻结后主区字母不可见）。同时显式设
        // TextElement.Foreground + TextBlock.Foreground=White，让字母无论经哪层呈现都强制白色。
        style.Setters.Add(
            new Setter(System.Windows.Documents.TextElement.ForegroundProperty, Brushes.White)
        );
        style.Setters.Add(new Setter(TextBlock.ForegroundProperty, Brushes.White));
        style.Setters.Add(new Setter(Control.FontWeightProperty, FontWeights.Bold));
        style.Setters.Add(new Setter(Control.BackgroundProperty, HeaderBackgroundBrush));
        style.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0, 0, 1, 1)));
        style.Setters.Add(
            new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Left)
        );
        style.Setters.Add(
            new Setter(Control.VerticalContentAlignmentProperty, VerticalAlignment.Center)
        );
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(6, 2, 2, 2)));
        style.Setters.Add(new Setter(FrameworkElement.MinHeightProperty, 24d));
        return style;
    }

    private DataGrid BuildGrid(ColumnStore store, Dictionary<(int Row, int Col), string> commentMap)
    {
        var grid = new DataGrid
        {
            // ItemsSource = state.View 在 LoadFile 里构建 state 后设置（虚拟化视图包装 ColumnStore）。
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            CanUserSortColumns = false,
            EnableRowVirtualization = true,
            
            EnableColumnVirtualization = true,
            SelectionUnit = DataGridSelectionUnit.Cell,
            HeadersVisibility = DataGridHeadersVisibility.All,
            GridLinesVisibility = DataGridGridLinesVisibility.All,
            HorizontalGridLinesBrush = GridLineBrush,
            VerticalGridLinesBrush = GridLineBrush,
            BorderBrush = GridLineBrush,
            BorderThickness = new Thickness(1),
            RowHeight = double.NaN,
        };

        // #P8-1：列头样式（Excel 字母坐标白色粗体可见）
        grid.ColumnHeaderStyle = BuildColumnHeaderStyle();
        VirtualizingPanel.SetScrollUnit(grid, ScrollUnit.Pixel);


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

        // P5：必须在添加列之前开启 DataGridExtensions 浮动筛选行——DGX 通过 Columns.CollectionChanged
        // 给"新加入且 HeaderTemplate==null 且 IsFilterVisible"的列套上筛选头模板；若在加列之后才开启，
        // 既有列不会被回溯套模板，筛选行就不渲染（实测踩坑）。筛选执行由 grid.DataContext 上的
        // ColumnStoreFilterAdapter(ICustomFilter) 路由到 VirtualizingSortableView.ApplyFilter（不整表物化）。
        // #P9-2：先注入 DGX 筛选框深色样式资源（覆盖 ColumnHeaderSearchTextBoxStyleKey），保证筛选框深底白字。
        ApplyDarkFilterResources(grid);
        DataGridExtensions.DataGridFilter.SetIsAutoFilterEnabled(grid, true);

        // 数据列 + 列头筛选 TextBox
        BuildDataColumns(grid, store, withFilterBox: true);

        // 单元格样式：深色主题边框 + 脏数据高亮 + 备注提示
        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        cellStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        // P13-2：不再对单元格设 MaxHeight 限高——RowHeight=NaN 撑高整行后，被限高的单元格会悬空
        // 居中在高行中间，自身四边框如实画出，视觉上呈现为行内部多余的横线。改为默认 Stretch，
        // 单元格随行高拉伸铺满，消除悬空边框。
        cellStyle.Setters.Add(
            new EventSetter(
                MouseEnterEvent,
                new MouseEventHandler(
                    (s, _) =>
                    {
                        if (s is not DataGridCell { DataContext: RowView view } cell)
                            return;
                        var rowIndex = view.RowIndex;
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
                || args.Column is not DataGridBoundColumn
                || args.Row.Item is not RowView view
            )
                return;

            var rowIndex = view.RowIndex;
            var colIndex = args.Column.DisplayIndex;
            var oldValue = view[colIndex];
            var newValue = (args.EditingElement as TextBox)?.Text ?? string.Empty;
            if (oldValue == newValue)
                return;
            var state = CurrentSheetState;
            if (state is null)
                return;
            // 写回 ColumnStore（RowView 索引器 set → SetCell 标脏）。DataGrid 提交时会自己写回，
            // 这里显式写一次确保脏跟踪一致（等值早已 return，不会重复标脏）。
            view[colIndex] = newValue;
            state.UndoStack.Push(
                new CellBatchAction([new CellEditRecord(rowIndex, colIndex, oldValue, newValue)])
            );
            state.RedoStack.Clear();
            MarkDirty(grid, view, colIndex);
            MarkCurrentFileDirty();

            // P13：修复编辑态残留的 MaxHeight，强制重新测量行高
            grid.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            {
                args.Row.InvalidateMeasure();
            }));
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
    /// P12：单元格编辑态样式——超长文本自动换行显示（不再横向裁切/触发横向自动滚动），Enter 键
    /// 提交编辑而不插入换行符（<see cref="TextBox.AcceptsReturn"/>=false）。行高由 <c>RowHeight=NaN</c>
    /// 自适应换行后的实际高度（见 BuildGrid/BuildFrozenGrid）。
    /// </summary>
    private static readonly Style DataGridEditingTextBoxStyle = BuildDataGridEditingTextBoxStyle();

    private static Style BuildDataGridEditingTextBoxStyle()
    {
        var style = new Style(typeof(TextBox));
        style.Setters.Add(new Setter(TextBox.TextWrappingProperty, TextWrapping.Wrap));
        style.Setters.Add(new Setter(TextBox.AcceptsReturnProperty, false));
        style.Setters.Add(new Setter(TextBox.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Disabled));
        style.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Stretch));
        return style;
    }

    /// <summary>
    /// 构造数据列（DataGridTextColumn）。P5：列头筛选改由 DataGridExtensions 的浮动筛选行提供
    /// （<c>DataGridFilter.IsAutoFilterEnabled</c> 在 BuildGrid 开启），本方法只建纯列名头 + 标记
    /// 每列的 store 列号（SortMemberPath）+ 是否显示筛选框（withFilterBox → SetIsFilterVisible）。
    /// 类型感知的匹配在 <see cref="ColumnStoreFilterAdapter"/> 里按列类型执行，下沉 ColumnStore。
    /// withFilterBox=false（冻结行 grid）：仍可编辑，但不显示筛选框（冻结模式筛选禁用）。
    /// </summary>
    private void BuildDataColumns(DataGrid grid, ColumnStore store, bool withFilterBox)
    {
        grid.Columns.Clear();
        for (var c = 0; c < store.ColumnCount; c++)
        {
            var colIndex = c;
            var columnName = store.ColumnNames[c];

            var column = new DataGridTextColumn
            {
                Header = columnName,
                // RowView 的 this[int col] 索引器 —— 与旧 DataRowView[int] 相同的索引器绑定语法。
                Binding = new Binding($"[{colIndex}]"),
                Width = new DataGridLength(160),
                IsReadOnly = false, // 两个 grid 都可编辑（冻结区也可编辑）
                CanUserResize = true,
                // 携带 store 列号，供 ColumnStoreFilterAdapter 从 DataGridColumn 反查列（列不可重排）。
                SortMemberPath = colIndex.ToString(),
                EditingElementStyle = DataGridEditingTextBoxStyle,
            };
            // DataGridExtensions 浮动筛选行：主 grid 显示筛选框，冻结 grid 不显示。
            column.SetIsFilterVisible(withFilterBox);
            if (withFilterBox)
            {
                // #2：给筛选框套暗色主题模板（默认 DGX 筛选框是白底，与 Fluent 暗色皮肤不搭）。
                column.SetTemplate(DarkFilterTemplate);
            }
            grid.Columns.Add(column);
        }
    }

    // #2/#P8-2：DataGridExtensions 暗色筛选框模板——TextBox 深底白字，双向绑定到列的 Filter 属性
    // （DGX 约定：模板里控件绑 DataGridFilterColumnControl.Filter）。全列共享一个只读模板实例。
    private static readonly ControlTemplate DarkFilterTemplate = BuildDarkFilterTemplate();

    // #P8-2：筛选框 TextBox 自己的显式 ControlTemplate——用我们自己的 Border 画深色背景，
    // 彻底绕开 WPF-UI 隐式 TextBox 样式和任何主题默认模板。全列共享一个只读实例。
    private static readonly ControlTemplate DarkFilterTextBoxTemplate =
        BuildDarkFilterTextBoxTemplate();

    // #P8-2：筛选框 TextBox 的显式 Style（非 null）——WPF 只有在 Style 为"未设置"时才回落到隐式样式；
    // 上一轮把 Style 设成 null 并不能可靠阻断 ControlsDictionary 里 {x:Type TextBox} 隐式样式的应用，
    // 于是筛选框仍走 WPF-UI Fluent 模板（白底）。这里给一个显式非 null Style（自带深色 ControlTemplate），
    // 显式 Style 会硬性阻断隐式样式查找，深色背景由我们的模板 Border 亲自绘制，白底不再有机会出现。
    private static readonly Style DarkFilterTextBoxStyle = BuildDarkFilterTextBoxStyle();

    /// <summary>
    /// #P9-2：把 DGX 筛选框 TextBox 的暗色样式注入到 grid 自身的 Resources，键 = DGX 的
    /// <c>DataGridFilter.ColumnHeaderSearchTextBoxStyleKey</c>。DGX 的默认列头模板（ColumnHeaderTemplateKey）
    /// 内部的筛选 TextBox 用 <c>DynamicResource {ColumnHeaderSearchTextBoxStyleKey}</c> 取样式（见 DGX Generic.xaml），
    /// 默认那份是浅色/透明（白底）。在 grid.Resources 里用同一个 key 覆盖成深色，命中的是 DGX 真正使用的那个
    /// TextBox——不依赖 <c>column.SetTemplate</c> 是否成功替换整个筛选控件模板（P7/P8 靠 SetTemplate 在
    /// 主 grid 生效但冻结 grid 仍白底，说明 SetTemplate 路径在两个 grid 上不稳定；grid 级 Resources 覆盖是
    /// DGX 原生样式钩子，对该 grid 内所有筛选框稳定生效）。两个 grid（主 + 冻结）都注入同一份，保证一致。
    /// </summary>
    private static void ApplyDarkFilterResources(DataGrid grid)
    {
        grid.Resources[DataGridFilter.ColumnHeaderSearchTextBoxStyleKey] =
            BuildDgxSearchTextBoxDarkStyle();
    }

    /// <summary>
    /// 构造 DGX 筛选框 TextBox 的深色 Style：深底(45,45,45)白字 + 自带深色 Border 模板（含 PART_ContentHost），
    /// 保留 DGX 默认的"无值时 Opacity=0、悬停/聚焦时 Opacity=1"触发器（否则空筛选框会一直占位可见）。
    /// </summary>
    private static Style BuildDgxSearchTextBoxDarkStyle()
    {
        var style = new Style(typeof(TextBox));
        style.Setters.Add(new Setter(FrameworkElement.MinWidthProperty, 20d));
        style.Setters.Add(new Setter(FrameworkElement.MarginProperty, new Thickness(4, 0, 2, 0)));
        style.Setters.Add(
            new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(45, 45, 45)))
        );
        style.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        style.Setters.Add(new Setter(TextBox.CaretBrushProperty, Brushes.White));
        style.Setters.Add(
            new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(90, 90, 90)))
        );
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
        style.Setters.Add(new Setter(Control.TemplateProperty, DarkFilterTextBoxTemplate));

        // DGX 默认：空值时隐藏筛选框（Opacity=0），悬停/聚焦时显示——保留该交互，否则每列都常驻一个深色框。
        var emptyTrigger = new Trigger { Property = TextBox.TextProperty, Value = "" };
        emptyTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 0d));
        var hoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
        hoverTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 1d));
        var focusTrigger = new Trigger { Property = UIElement.IsFocusedProperty, Value = true };
        focusTrigger.Setters.Add(new Setter(UIElement.OpacityProperty, 1d));
        style.Triggers.Add(emptyTrigger);
        style.Triggers.Add(hoverTrigger);
        style.Triggers.Add(focusTrigger);
        return style;
    }

    private static ControlTemplate BuildDarkFilterTextBoxTemplate()
    {
        // Border { Background=45,45,45; BorderBrush=90,90,90 } > ScrollViewer x:Name=PART_ContentHost
        var border = new FrameworkElementFactory(typeof(Border));
        border.SetValue(Border.BackgroundProperty, new SolidColorBrush(Color.FromRgb(45, 45, 45)));
        border.SetValue(Border.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(90, 90, 90)));
        border.SetValue(Border.BorderThicknessProperty, new Thickness(1));
        border.SetValue(Border.SnapsToDevicePixelsProperty, true);

        var host = new FrameworkElementFactory(typeof(ScrollViewer), "PART_ContentHost");
        host.SetValue(ScrollViewer.FocusableProperty, false);
        host.SetValue(
            ScrollViewer.HorizontalScrollBarVisibilityProperty,
            ScrollBarVisibility.Hidden
        );
        host.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Hidden);
        host.SetValue(FrameworkElement.MarginProperty, new Thickness(2, 0, 2, 0));
        border.AppendChild(host);

        return new ControlTemplate(typeof(TextBox)) { VisualTree = border };
    }

    private static Style BuildDarkFilterTextBoxStyle()
    {
        var style = new Style(typeof(TextBox));
        style.Setters.Add(
            new Setter(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(45, 45, 45)))
        );
        style.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        style.Setters.Add(new Setter(TextBox.CaretBrushProperty, Brushes.White));
        style.Setters.Add(
            new Setter(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(90, 90, 90)))
        );
        style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
        style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(2, 0, 2, 0)));
        // 关键：显式 ControlTemplate（我们的深色 Border），彻底绕开任何主题默认/隐式模板。
        style.Setters.Add(new Setter(Control.TemplateProperty, DarkFilterTextBoxTemplate));
        return style;
    }

    private static ControlTemplate BuildDarkFilterTemplate()
    {
        var tb = new FrameworkElementFactory(typeof(TextBox));
        // #P8-2：给 TextBox 一个显式非 null Style（自带深色 ControlTemplate）。上一轮（P7）把 Style
        // 设为 null 试图退出 WPF-UI 隐式样式，实机证明无效（截图仍白底）——WPF 对 Style=null 与 Style
        // 未设置在隐式样式回落上行为不可靠。显式 Style 会硬性阻断 {x:Type TextBox} 隐式样式，深色背景
        // 由该 Style 里的 ControlTemplate Border 亲自绘制，45,45,45 真正落地。
        tb.SetValue(FrameworkElement.StyleProperty, DarkFilterTextBoxStyle);
        tb.SetValue(Control.BackgroundProperty, new SolidColorBrush(Color.FromRgb(45, 45, 45)));
        tb.SetValue(Control.ForegroundProperty, Brushes.White);
        tb.SetValue(Control.BorderBrushProperty, new SolidColorBrush(Color.FromRgb(90, 90, 90)));
        tb.SetValue(Control.BorderThicknessProperty, new Thickness(1));
        tb.SetValue(FrameworkElement.MarginProperty, new Thickness(1));
        tb.SetValue(Control.PaddingProperty, new Thickness(2, 0, 2, 0));
        tb.SetValue(FrameworkElement.MinHeightProperty, 20d);
        tb.SetValue(TextBox.CaretBrushProperty, Brushes.White);
        tb.SetValue(FrameworkElement.ToolTipProperty, "输入筛选值（回车/停顿后生效）");
        // 绑定到 DataGridFilterColumnControl.Filter（DGX 把该控件作为模板承载者）
        tb.SetBinding(
            TextBox.TextProperty,
            new Binding("Filter")
            {
                RelativeSource = new RelativeSource(
                    RelativeSourceMode.FindAncestor,
                    typeof(DataGridExtensions.DataGridFilterColumnControl),
                    1
                ),
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                Mode = BindingMode.TwoWay,
            }
        );
        return new ControlTemplate(typeof(DataGridExtensions.DataGridFilterColumnControl))
        {
            VisualTree = tb,
        };
    }

    /// <summary>
    /// 从 ColumnStore 采样前 100 行非空值推断列类型（等价旧 ColumnTypeDetector.Detect(DataTable,...) 的规则）。
    /// </summary>
    private static ColumnType DetectColumnType(ColumnStore store, int col, int sampleSize = 100)
    {
        if (store.RowCount is 0)
            return ColumnType.Text;
        var table = new DataTable();
        var name = store.ColumnNames[col];
        table.Columns.Add(name, typeof(string));
        var maxRows = Math.Min(sampleSize, store.RowCount);
        for (var r = 0; r < maxRows; r++)
        {
            var row = table.NewRow();
            row[0] = store.GetCell(r, col) ?? string.Empty;
            table.Rows.Add(row);
        }

        return ColumnTypeDetector.Detect(table, name, sampleSize);
    }

    /// <summary>
    /// P5：清除当前 sheet 的所有列头筛选。清 DataGridExtensions 的浮动筛选行（会触发 ICustomFilter
    /// 的 OnFilterChanged → View.ClearFilter），并兜底直接 ClearFilter 恢复全部行。
    /// </summary>
    private void OnClearFilterClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;

        // 冻结模式下筛选行本就禁用，只需确保视图无残留筛选
        if (state.FrozenRows > 0)
        {
            StatusText.Text = "冻结模式下列筛选不可用";
            return;
        }

        // 清 DataGridExtensions 各列筛选框（触发 adapter.OnFilterChanged → View.ClearFilter）
        DataGridExtensions.DataGridFilter.GetFilter(state.MainGrid)?.Clear();
        // 兜底：直接清视图筛选（若无激活筛选，Clear 不回调 adapter）
        state.View.ClearFilter();
        StatusText.Text = $"筛选已清除（{state.Store.RowCount} 行）";
    }

    private void OnAddRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        var at = state.Store.RowCount;
        state.Store.AppendRow(); // ColumnStore 末尾追加空行（AppendRow 不标脏、不置 StructureChanged）
        state.TotalRows = state.Store.RowCount;
        state.LoadedRows = state.Store.RowCount;
        RefreshMainViewAfterStructuralChange(state);
        // #6：增行进撤销栈（撤销=删该行，重做=在同位再插空行），不再清空撤销栈。
        state.UndoStack.Push(new InsertRowAction(at));
        state.RedoStack.Clear();
        MarkCurrentFileDirty();
    }

    private void OnDeleteRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        DeleteCurrentRow(state.MainGrid);
    }

    /// <summary>
    /// 在当前选中行下方插入空行。<see cref="ColumnStore.InsertRow"/> 现 remap 脏跟踪并置 StructureChanged（P4 WF1）。
    /// </summary>
    private void InsertRowBelow(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var insertAt = GetCurrentRowIndex(grid);
        if (insertAt < 0)
            insertAt = state.Store.RowCount; // 没选中则追加到末尾
        else
            insertAt += 1; // 在下方

        state.Store.InsertRow(insertAt);
        state.TotalRows = state.Store.RowCount;
        state.LoadedRows = state.Store.RowCount;
        RefreshMainViewAfterStructuralChange(state);
        // #6：插行进撤销栈（撤销=删该行，重做=在同位再插空行）。
        state.UndoStack.Push(new InsertRowAction(insertAt));
        state.RedoStack.Clear();
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 删除当前选中行（支持多选）。<see cref="ColumnStore.DeleteRow"/> 现 remap 脏跟踪并置 StructureChanged（P4 WF1）。
    /// 多选时按行号降序删除，避免删除后行号移位。#6：删除前抓取每行完整内容快照，压 <see cref="DeleteRowsAction"/>
    /// 供 Ctrl+Z 精确还原内容 + 位置。
    /// </summary>
    private void DeleteCurrentRow(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var rowIndices = grid
            .SelectedItems.OfType<RowView>()
            .Select(v => v.RowIndex)
            .Distinct()
            .OrderByDescending(i => i)
            .ToList();
        if (rowIndices.Count is 0)
        {
            var cur = GetCurrentRowIndex(grid);
            if (cur >= 0)
                rowIndices.Add(cur);
        }

        var store = state.Store;
        var cols = store.ColumnCount;
        // #6：删除前抓取被删行的完整内容快照（行号 + 整行值），供撤销时按位置回填。
        var snapshots = new List<(int Row, string?[] Values)>(rowIndices.Count);
        foreach (var rowIndex in rowIndices)
        {
            if (rowIndex >= 0 && rowIndex < store.RowCount)
            {
                var values = new string?[cols];
                for (var c = 0; c < cols; c++)
                    values[c] = store.GetCell(rowIndex, c);
                snapshots.Add((rowIndex, values));
                store.DeleteRow(rowIndex);
            }
        }

        state.TotalRows = store.RowCount;
        state.LoadedRows = store.RowCount;
        RefreshMainViewAfterStructuralChange(state);
        if (snapshots.Count > 0)
        {
            state.UndoStack.Push(new DeleteRowsAction(snapshots));
            state.RedoStack.Clear();
        }
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 结构变更（增删行）后刷新视图：冻结模式下重建两个 RowRangeView（行数变了），否则重建 SortableView 行序。
    /// </summary>
    private void RefreshMainViewAfterStructuralChange(SheetState state)
    {
        if (state.FrozenRows > 0 && state.FrozenRows < state.Store.RowCount)
        {
            ApplyFreezeRows(state);
        }
        else
        {
            state.View.ClearFilter(); // 重建 _rowOrder 到 [0..RowCount)（结构变更后同步视图 + Reset）
        }
    }

    /// <summary>
    /// 在当前选中列右侧插入空列。ColumnStore 仅支持在末尾扩列（<see cref="ColumnStore.EnsureColumnCount"/>），
    /// 故新列固定加到最右，不支持任意位置插入（已知限制，见 status.md）。
    /// </summary>
    private void InsertColumnRight(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        var store = state.Store;
        var newCount = store.ColumnCount + 1;
        // 列名工厂：与撤销后重做保持一致（同一 store 状态下生成相同列名）。
        string NameFactory(int col) => NextColumnName(store, col);
        store.EnsureColumnCount(newCount, NameFactory);

        RebuildGridColumns(grid, store);
        // #6：插列进撤销栈（撤销=删最右列，重做=再加一列）。
        state.UndoStack.Push(new InsertColumnAction(NameFactory));
        state.RedoStack.Clear();
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 删除当前选中列：ColumnStore 不支持删列（无 RemoveColumn API），当前作为已知限制提示用户。
    /// </summary>
    private void DeleteCurrentColumn(DataGrid grid)
    {
        if (CurrentSheetState is not { } state)
            return;
        _ = state;
        StatusText.Text = "删除列在 P3.2 列式存储下暂不支持（已知限制）";
    }

    /// <summary>为末尾新增列生成不重复的 Excel 列名。</summary>
    private static string NextColumnName(ColumnStore store, int col)
    {
        var name = GetExcelColumnName(col + 1);
        var existing = store.ColumnNames.ToList();
        while (existing.Contains(name))
            name += "_";
        return name;
    }

    /// <summary>
    /// 重建主 DataGrid 列（增删列后调用，因为 AutoGenerateColumns=false）。
    /// </summary>
    private void RebuildGridColumns(DataGrid grid, ColumnStore store)
    {
        BuildDataColumns(grid, store, withFilterBox: true);
    }

    /// <summary>
    /// 获取当前选中单元格的行索引（ColumnStore 行号）。
    /// </summary>
    private static int GetCurrentRowIndex(DataGrid grid)
    {
        if (grid.CurrentCell.Item is RowView view)
            return view.RowIndex;
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

        // P4: 在 UI 线程从 ColumnStore 组装写回计划（纯数据，不把 Store 传进后台线程）。
        // 无结构变更(StructureChanged=false) → 只写 DirtyCells（增量）；有结构变更 → 整表全量写回。
        var savedStores = _sheets.Values.Where(s => s.FilePath == filePath).ToList();
        var plans = savedStores
            .Select(s =>
            {
                var store = s.Store;
                var rows = store.RowCount;
                var cols = store.ColumnCount;
                if (store.StructureChanged)
                {
                    var data = new string[rows, cols];
                    for (var r = 0; r < rows; r++)
                    for (var c = 0; c < cols; c++)
                        data[r, c] = store.GetCell(r, c) ?? string.Empty;
                    return new SheetWritePlan(s.SheetName, Full: true, rows, cols, data, []);
                }

                var dirty = store
                    .DirtyCells.Select(cell =>
                        (cell.Row, cell.Col, (string?)store.GetCell(cell.Row, cell.Col))
                    )
                    .ToList();
                return new SheetWritePlan(s.SheetName, Full: false, rows, cols, null, dirty);
            })
            .ToList();

        var totalDirty = plans.Sum(p => p.Full ? p.RowCount * p.ColCount : p.DirtyCells.Count);

        Tabs.IsEnabled = false;
        Cursor = Cursors.Wait;
        StatusText.Text = $"正在保存：{Path.GetFileName(filePath)}…";

        try
        {
            var (elapsedMs, error) = await Task.Run(() =>
            {
                try
                {
                    // 原子写：ExcelWriteBack 以原文件为模板写到 tmp（保留格式+剥离图表公式+只写脏/全量），
                    // 成功后 AtomicFileWriter 用 File.Replace(tmp, 原文件, .bak) 原子替换。P0 机制不变。
                    var result = AtomicFileWriter.Write(
                        filePath,
                        tempPath => ExcelWriteBack.Write(filePath, tempPath, plans)
                    );

                    return (sw.ElapsedMilliseconds, result.Error);
                }
                catch (Exception ex)
                {
                    return (0L, ex);
                }
            });

            if (error is not null)
                throw error;

            sw.Stop();
            // 保存成功：清脏跟踪（下次无编辑即可秒过），移除文件脏标记。ColumnStore 单线程，UI 线程调。
            foreach (var s in savedStores)
                s.Store.ClearDirty();
            _dirtyFiles.Remove(filePath);
            UpdateTitle();
            StatusText.Text =
                $"已保存：{Path.GetFileName(filePath)}（耗时 {elapsedMs} ms，写入 {totalDirty} 格）";
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
        // #6：撤销单元是 IUndoableAction（单格/粘贴/增删行列统一）。结构性动作撤销后需刷新行序/列。
        var isStructural = state.UndoStack.Peek().IsStructural;
        UndoableStack.Undo(state.Store, state.UndoStack, state.RedoStack);
        AfterUndoRedo(state, isStructural);
    }

    private void OnRedoClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        if (state.RedoStack.Count == 0)
            return;
        var isStructural = state.RedoStack.Peek().IsStructural;
        UndoableStack.Redo(state.Store, state.UndoStack, state.RedoStack);
        AfterUndoRedo(state, isStructural);
    }

    /// <summary>
    /// 撤销/重做后刷新 UI：结构性动作（增删行列）需重建行序/列并同步行数；非结构性只刷新单元格显示。
    /// </summary>
    private void AfterUndoRedo(SheetState state, bool isStructural)
    {
        if (isStructural)
        {
            // 列数可能变了（增删列）→ 重建主 grid 列；行数可能变了 → 重建行序 + 同步计数。
            RebuildGridColumns(state.MainGrid, state.Store);
            state.TotalRows = state.Store.RowCount;
            state.LoadedRows = state.Store.RowCount;
            RefreshMainViewAfterStructuralChange(state);
        }

        state.MainGrid.Items.Refresh();
        state.FrozenGrid?.Items.Refresh();
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
        if (startCell.Column is null || startCell.Item is not RowView rowView)
            return;
        if (CurrentSheetState is not { } state)
            return;
        var colCount = state.Store.ColumnCount;
        // 用视图索引：筛选模式下连续粘贴进"可见的连续行"，正确且不越界。
        var startViewIndex = grid.Items.IndexOf(rowView);
        var startCol = startCell.Column.DisplayIndex;

        var text = Clipboard.GetText();
        if (string.IsNullOrEmpty(text))
            return;

        var lines = text.Split(["\r\n"], StringSplitOptions.RemoveEmptyEntries);
        // #6：把这次粘贴产生的所有格改动记为一个复合撤销单元（一次 Ctrl+Z 整体撤销这次粘贴）。
        var batch = new List<CellEditRecord>();
        for (var i = 0; i < lines.Length; i++)
        {
            var targetView = startViewIndex + i;
            if (targetView < 0 || targetView >= grid.Items.Count)
                break;
            if (grid.Items[targetView] is not RowView tv)
                continue;
            var cells = lines[i].Split('\t');
            for (var j = 0; j < cells.Length; j++)
            {
                var targetCol = startCol + j;
                if (targetCol >= colCount)
                    break;
                var oldValue = tv[targetCol];
                if (oldValue == cells[j])
                    continue; // 值未变不记录
                batch.Add(new CellEditRecord(tv.RowIndex, targetCol, oldValue, cells[j]));
                tv[targetCol] = cells[j]; // RowView 索引器 → SetCell 标脏
                // 粘贴的格子也要绿色高亮（和手动编辑一致）
                MarkDirty(grid, tv, targetCol);
            }
        }
        // 粘贴绕过 CellEditEnding，得手动记撤销 + 标脏，否则 Ctrl+Z 撤不掉、关窗不提示保存=数据丢失
        if (batch.Count > 0)
        {
            state.UndoStack.Push(new CellBatchAction(batch));
            state.RedoStack.Clear();
        }
        MarkCurrentFileDirty();
    }

    /// <summary>
    /// 标脏单元格绿色。按 RowView 定位容器，筛选/虚拟化下都正确（越界/不可见静默跳过）。
    /// </summary>
    private static void MarkDirty(DataGrid grid, RowView view, int col)
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
        // #P9-3：一条独立贯穿 Border 画分界线（跨冻结行 grid + 主 grid），不再在两个 grid 各自列上画边框。
        ApplyFreezeColumnDivider(state, n);
        SaveFreeze(state);
        StatusText.Text = $"已冻结前 {n} 列";
    }

    // #P9-3：分界竖线颜色（亮灰），供独立贯穿 Border 使用。
    private static readonly Brush FreezeDividerBrush = new SolidColorBrush(
        Color.FromRgb(120, 120, 120)
    );

    /// <summary>
    /// #P9-3：冻结列分界竖线——彻底改用【一条独立 Border，跨 panel 两行（冻结行 grid + 主 grid）贯穿绘制】，
    /// 取代 P7/P8 的"在两个 DataGrid 各自的列 CellStyle/HeaderStyle 里画右边框、再指望它们自然对齐"方案。
    /// 旧方案根因：列头分界线画在【冻结行 grid】的列上，数据区分界线画在【主 grid】的列上——两个 grid 是
    /// 不同控件，各自的冻结列实际宽度/行号头宽在布局时刻可能不一致（列宽同步是异步的 ActualWidth 事件驱动），
    /// 于是上下两段线 X 坐标错位、交界断层（P7/P8/本轮截图三次证实）。
    /// 现在只画一条 Border，X = 行号头宽 + 前 N 个冻结列实际宽之和，RowSpan 跨两行，物理上就是一条线，
    /// 不存在跨 grid 对齐问题。X 随列宽变化在 LayoutUpdated 里重算（冻结列被 FrozenColumnCount 钉住不随横滚移动）。
    /// n&lt;=0 或 n&gt;=列数时移除分界线。
    /// </summary>
    private void ApplyFreezeColumnDivider(SheetState state, int n)
    {
        var panel = state.Panel;
        if (panel is null)
            return;

        // 先移除旧分界线（含事件挂钩）
        if (state.FrozenDivider is { } old)
        {
            panel.Children.Remove(old);
            state.FrozenDivider = null;
        }

        if (n <= 0 || n >= state.MainGrid.Columns.Count)
            return;

        var divider = new Border
        {
            Width = 2,
            Background = FreezeDividerBrush,
            HorizontalAlignment = HorizontalAlignment.Left,
            VerticalAlignment = VerticalAlignment.Stretch,
            IsHitTestVisible = false, // 纯装饰，不拦截鼠标（拖选/编辑不受影响）
            SnapsToDevicePixels = true,
        };
        panel.Children.Add(divider);
        Grid.SetRow(divider, 0);
        Grid.SetRowSpan(divider, panel.RowDefinitions.Count); // 跨冻结行 grid + 主 grid 两行
        Panel.SetZIndex(divider, 100); // 压在两个 grid 之上，保证可见
        state.FrozenDivider = divider;

        // X 位置随布局/列宽变化重算：行号头宽 + 前 N 个冻结列实际宽之和。
        void Reposition()
        {
            if (state.FrozenDivider is not { } d)
                return;
            // 冻结列在冻结行 grid（若有）和主 grid 里宽度经 SyncFrozenColumnWidths 保持一致，取主 grid 为准。
            var grid = state.MainGrid;
            var x = grid.RowHeaderActualWidth;
            var take = Math.Min(n, grid.Columns.Count);
            for (var i = 0; i < take; i++)
            {
                var w = grid.Columns[i].ActualWidth;
                x += w > 0 ? w : grid.Columns[i].Width.DisplayValue;
            }
            // Border 用左对齐 + 左 Margin 定位到 X（减去自身一半宽让线压在列边界上）。
            d.Margin = new Thickness(x - 1, 0, 0, 0);
        }

        // 首次 + 后续布局变化时重算 X（列宽调整、字体等都会触发 LayoutUpdated）。
        panel.LayoutUpdated -= state.DividerLayoutHandler;
        state.DividerLayoutHandler = (_, _) => Reposition();
        panel.LayoutUpdated += state.DividerLayoutHandler;
        Reposition();
    }

    private void OnFreezeRowClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        var rowIndex = GetCurrentRowIndex(state.MainGrid);
        if (rowIndex < 0)
        {
            StatusText.Text = "请先选中一行再冻结到此行";
            return;
        }

        // 冻结 row 0..当前行（含），即前 N 行固定
        var n = rowIndex + 1;
        if (n >= state.Store.RowCount)
        {
            StatusText.Text = "不能冻结全部行";
            return;
        }

        state.FrozenRows = n;
        ApplyFreezeRows(state);
        SaveFreeze(state);
        StatusText.Text = $"已冻结前 {n} 行（冻结区固定表头；筛选作用于下方数据区）";
    }

    /// <summary>
    /// P4 冻结行重做 + #1 冻结与筛选共存：上下两个 DataGrid 共享同一 ColumnStore。
    /// 冻结 grid（顶部 row0）绑 <see cref="RowRangeView"/>(0, N) 且承载列头+浮动筛选行；
    /// 主 grid 绑可筛选的 <see cref="VirtualizingSortableView"/>，加基础谓词 row&gt;=N（只显示数据行区），
    /// 冻结 grid 的筛选路由到主区 View（BasePredicate AND 列筛选）。横向滚 + 列宽双向同步。两区都经 RowView 写回。
    /// </summary>
    private void ApplyFreezeRows(SheetState state)
    {
        var panel = state.Panel;
        if (panel is null)
            return;
        var n = state.FrozenRows;

        if (n <= 0 || n >= state.Store.RowCount)
        {
            RemoveFrozenRows(state);
            return;
        }

        // 建/复用冻结行 grid
        if (state.FrozenGrid is null)
        {
            var fg = BuildFrozenGrid(state);
            panel.Children.Add(fg);
            Grid.SetRow(fg, 0);
            state.FrozenGrid = fg;
            WireFreezeRowSync(state);
        }

        // 冻结 grid = 前 N 行（固定表头，RowRangeView 不筛不排，常驻可见）。
        state.FrozenGrid!.ItemsSource = new RowRangeView(state.Store, 0, n);

        // #1：主区仍绑可筛选的 VirtualizingSortableView，加"只显示第 N 行之后数据行"的基础谓词
        // （row>=n），与列头筛选 AND 组合——冻结与筛选共存。筛选框显示在冻结 grid 的浮动筛选行里
        // （主 grid 冻结时隐藏列头/筛选行，避免与冻结 grid 顶部列头重复）；冻结 grid 的筛选路由到主区 View。
        state.MainGrid.ItemsSource = state.View;
        DataGridExtensions.DataGridFilter.SetIsAutoFilterEnabled(state.MainGrid, false);
        if (state.Filter is { } filter)
        {
            filter.BasePredicate = row => row >= n;
            // 冻结 grid 的筛选行驱动主区筛选：其 DataContext 指向同一 adapter（filter 主区 View）。
            state.FrozenGrid.DataContext = filter;
            DataGridExtensions.DataGridFilter.SetIsAutoFilterEnabled(state.FrozenGrid, true);
            filter.Reapply(GetFilteredColumns(state.FrozenGrid));
        }
        else
        {
            state.View.ApplyFilter(row => row >= n);
        }

        // 主 grid 隐藏列头（冻结 grid 顶部已显示列头+筛选行，避免重复）
        state.MainGrid.HeadersVisibility = DataGridHeadersVisibility.Row;
        SyncFrozenColumnWidths(state, mainToFrozen: true);

        // #P9-3：冻结行 grid 建好后，冻结列分界线需要重建，让贯穿 Border 跨新增的冻结行 grid + 主 grid
        // 两行（此前若只冻结了列，分界线只跨主 grid 一行；现在两行都要覆盖）。
        if (state.FrozenColumns > 0)
            ApplyFreezeColumnDivider(state, state.FrozenColumns);
    }

    /// <summary>取一个 DataGrid 当前带激活筛选值的列（供 ColumnStoreFilterAdapter.Reapply）。</summary>
    private static IReadOnlyCollection<DataGridColumn> GetFilteredColumns(DataGrid grid) =>
        grid.Columns.Where(c => !string.IsNullOrEmpty(c.GetFilter()?.ToString())).ToList();

    /// <summary>拆冻结行：移除冻结 grid，主 grid 恢复全表 VirtualizingSortableView（清基础谓词，筛选/排序恢复）。</summary>
    private void RemoveFrozenRows(SheetState state)
    {
        if (state.Panel is { } panel && state.FrozenGrid is { } fg)
        {
            panel.Children.Remove(fg);
        }
        state.FrozenGrid = null;
        state.FrozenScroll = null;
        state.MainScroll = null;
        state.FrozenRows = 0;
        state.MainGrid.ItemsSource = state.View;
        state.MainGrid.HeadersVisibility = DataGridHeadersVisibility.All;
        // #1：清掉"row>=n"基础谓词，恢复全表筛选。
        if (state.Filter is { } filter)
        {
            filter.BasePredicate = null;
            filter.Reapply(GetFilteredColumns(state.MainGrid));
        }
        else
        {
            state.View.ClearFilter();
        }
        DataGridExtensions.DataGridFilter.SetIsAutoFilterEnabled(state.MainGrid, true);

        // #P9-3：冻结行 grid 已移除，若仍冻结着列，分界线需重建为只跨主 grid 一行（否则贯穿 Border 会
        // 悬在已折叠的 row0 空槽里）。若列也没冻结，ApplyFreezeColumnDivider(state,0) 会移除分界线。
        ApplyFreezeColumnDivider(state, state.FrozenColumns);
    }

    /// <summary>
    /// 构造冻结行 DataGrid：#8 视觉与主 grid 完全一致（同背景/字体/行高，无特殊装饰），
    /// 只靠底部一条加粗分割线与主区分界（在 ApplyFreezeRows 后由 grid 的 BorderThickness 体现）。
    /// 关纵向滚、隐藏横向滚（由主 grid 驱动），可编辑；显示列头（主 grid 冻结时隐藏列头）。
    /// </summary>
    private DataGrid BuildFrozenGrid(SheetState state)
    {
        var grid = new DataGrid
        {
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            CanUserSortColumns = false,
            CanUserReorderColumns = false,
            RowHeight = double.NaN,
            EnableRowVirtualization = false, // 仅 N 行（典型 4），关虚拟化保证行号头稳定
            SelectionUnit = DataGridSelectionUnit.Cell,
            HeadersVisibility = DataGridHeadersVisibility.All,
            GridLinesVisibility = DataGridGridLinesVisibility.All,
            HorizontalGridLinesBrush = GridLineBrush,
            VerticalGridLinesBrush = GridLineBrush,
            // #8：与主区分界处一条加粗亮线（底 3px 亮灰），其余边框与主 grid 一致（1px 暗灰）。
            BorderThickness = new Thickness(1, 1, 1, 3),
            HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
            VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            RowHeaderWidth = 50,
        };
        // #P8-1：冻结 grid 也用同一列头样式，保证冻结态下 Excel 字母坐标白色粗体可见。
        grid.ColumnHeaderStyle = BuildColumnHeaderStyle();
        VirtualizingPanel.SetScrollUnit(grid, ScrollUnit.Pixel);

        // #8：分界线——冻结 grid 底边框用加粗亮灰（与冻结列分界线同色 120,120,120），其余边框暗灰。
        // 用 BorderBrush 单一色 + 底部加粗厚度即可呈现"一条分割线"效果，且四周细边与主 grid 融为一体。
        grid.BorderBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120));

        // 行号头：与主 grid 一致（白字、居中、加粗），显示绝对 Excel 行号
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

        // #8：行头样式与主 grid 完全一致（同背景 GridLineBrush、白字、居中）——去掉之前的差异化。
        var rowHeaderStyle = new Style(typeof(DataGridRowHeader));
        rowHeaderStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        rowHeaderStyle.Setters.Add(new Setter(Control.BackgroundProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowHeaderStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowHeaderStyle.Setters.Add(
            new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center)
        );
        grid.RowHeaderStyle = rowHeaderStyle;

        // #8：行样式与主 grid 一致（同边框、白字，无特殊深色背景）。
        var rowStyle = new Style(typeof(DataGridRow));
        rowStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        rowStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        rowStyle.Setters.Add(new Setter(Control.ForegroundProperty, Brushes.White));
        grid.RowStyle = rowStyle;

        // 单元格样式与主 grid 一致（同边框），保证行高/字体/网格线完全统一。
        var cellStyle = new Style(typeof(DataGridCell));
        cellStyle.Setters.Add(new Setter(Control.BorderBrushProperty, GridLineBrush));
        cellStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0.5)));
        // P13-2：同主 grid，不再限高，避免悬空单元格残留多余横线。
        grid.CellStyle = cellStyle;

        // #1：冻结 grid 顶部承载列头 + DGX 浮动筛选行（主 grid 冻结时隐藏列头）。必须在加列前开启
        // 自动筛选（DGX 靠 Columns.CollectionChanged 给新列套筛选头模板，加列后再开启不回溯——P5 踩坑）。
        // #P9-2：先注入 DGX 筛选框深色样式资源（与主 grid 同一份），保证冻结 grid 的筛选框也深底白字
        // （此前冻结 grid 筛选框白底 = SetTemplate 路径在冻结 grid 上未生效；grid 级 Resources 覆盖稳定）。
        ApplyDarkFilterResources(grid);
        DataGridExtensions.DataGridFilter.SetIsAutoFilterEnabled(grid, true);
        // 列：与主 grid 同结构，带筛选框（withFilterBox=true）。编辑经 RowView 写回；筛选路由到主区 View
        // （在 ApplyFreezeRows 里把本 grid 的 DataContext 指向 state.Filter，带 BasePredicate=row>=n）。
        BuildDataColumns(grid, state.Store, withFilterBox: true);
        // #4：行号用绝对 store 行号 +1（冻结区 RowRangeView(0,n) → 1..n），与主区连续对齐。
        grid.LoadingRow += (_, args) =>
            args.Row.Header = RowHeaderNumber(args.Row.Item, args.Row.GetIndex());

        // #5：用户在冻结 grid 里调列宽 → 同步到主 grid（冻结时列头显示在冻结 grid，用户主要在这里拖列宽）。
        foreach (var col in grid.Columns)
        {
            System
                .ComponentModel.DependencyPropertyDescriptor.FromProperty(
                    DataGridColumn.ActualWidthProperty,
                    typeof(DataGridColumn)
                )
                .AddValueChanged(
                    col,
                    (_, _) =>
                    {
                        if (!_syncingColumnWidths)
                            SyncFrozenColumnWidths(state, mainToFrozen: false);
                    }
                );
        }

        // 编辑提交 → 写回 ColumnStore + 撤销栈（与主 grid 同一套逻辑）
        grid.CellEditEnding += (_, args) =>
        {
            if (
                args.EditAction != DataGridEditAction.Commit
                || args.Column is not DataGridBoundColumn
                || args.Row.Item is not RowView view
            )
                return;

            var colIndex = args.Column.DisplayIndex;
            var oldValue = view[colIndex];
            var newValue = (args.EditingElement as TextBox)?.Text ?? string.Empty;
            if (oldValue == newValue)
                return;
            view[colIndex] = newValue; // RowView → ColumnStore.SetCell（真实行号，标脏）
            state.UndoStack.Push(
                new CellBatchAction([
                    new CellEditRecord(view.RowIndex, colIndex, oldValue, newValue),
                ])
            );
            state.RedoStack.Clear();
            MarkDirty(grid, view, colIndex);
            MarkCurrentFileDirty();

            // P13：修复编辑态残留的 MaxHeight，强制重新测量行高
            grid.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            {
                args.Row.InvalidateMeasure();
            }));
        };

        return grid;
    }

    /// <summary>横向滚动同步（主 grid ↔ 冻结 grid 双向）：只转发横向偏移，不在滚动里同步列宽（#3：
    /// 编辑超长文本拖选会触发滚动，若此时同步 ActualWidth 会把编辑态的瞬时宽度写进列宽导致分界线偏移）。
    /// #P10：必须双向——冻结 grid 承载列头/筛选行，用户在冻结 grid 里点表头/筛选框或 DataGrid 自动
    /// BringIntoView 会让冻结 grid 独立横滚；旧的单向（主→冻结）不回传，导致冻结 grid 横滚偏移后
    /// 主区所有未冻结列相对冻结区整体水平错开一个恒定距离（=两 grid 横滚偏移差）。</summary>
    private void WireFreezeRowSync(SheetState state)
    {
        state.MainGrid.AddHandler(
            ScrollViewer.ScrollChangedEvent,
            new ScrollChangedEventHandler(
                (_, e) =>
                {
                    if (state.FrozenGrid is not { } fg)
                        return;
                    if (_syncingScroll)
                        return;
                    // 只在横向偏移真正变化时转发；不做列宽同步（列宽同步走 ActualWidth 变更事件，见 #3/#5）。
                    if (Math.Abs(e.HorizontalChange) > 0.01)
                    {
                        state.FrozenScroll ??= FindScrollViewer(fg);
                        if (state.FrozenScroll is { } fs)
                        {
                            _syncingScroll = true;
                            try
                            {
                                fs.ScrollToHorizontalOffset(e.HorizontalOffset);
                            }
                            finally
                            {
                                _syncingScroll = false;
                            }
                        }
                    }
                }
            )
        );

        // #P10：反向同步——冻结 grid 独立横滚（点表头/筛选框/自动 BringIntoView）时回传到主 grid，
        // 保证两 grid 未冻结列横向偏移始终一致（否则出现恒定水平错位，见方法注释）。
        state.FrozenGrid?.AddHandler(
            ScrollViewer.ScrollChangedEvent,
            new ScrollChangedEventHandler(
                (_, e) =>
                {
                    if (_syncingScroll)
                        return;
                    if (Math.Abs(e.HorizontalChange) > 0.01)
                    {
                        state.MainScroll ??= FindScrollViewer(state.MainGrid);
                        if (state.MainScroll is { } ms)
                        {
                            _syncingScroll = true;
                            try
                            {
                                ms.ScrollToHorizontalOffset(e.HorizontalOffset);
                            }
                            finally
                            {
                                _syncingScroll = false;
                            }
                        }
                    }
                }
            )
        );

        // #5：用户在主 grid 里调列宽 → 同步到冻结 grid。
        foreach (var col in state.MainGrid.Columns)
        {
            System
                .ComponentModel.DependencyPropertyDescriptor.FromProperty(
                    DataGridColumn.ActualWidthProperty,
                    typeof(DataGridColumn)
                )
                .AddValueChanged(
                    col,
                    (_, _) =>
                    {
                        if (!_syncingColumnWidths)
                            SyncFrozenColumnWidths(state, mainToFrozen: true);
                    }
                );
        }
    }

    // #3/#5：列宽同步的重入保护——同步时设 true，避免"设 A 宽→触发 A 的 ActualWidth 变更→又同步"死循环。
    private bool _syncingColumnWidths;

    // #P10：横滚双向同步的重入保护——一侧滚动写另一侧偏移会再触发对方 ScrollChanged，设 true 阻断回环。
    private bool _syncingScroll;

    /// <summary>
    /// #5：列宽双向同步。<paramref name="mainToFrozen"/>=true 时主→冻结，false 时冻结→主。
    /// 用 <see cref="_syncingColumnWidths"/> 防重入；只在宽度确有差异时写，避免抖动。
    /// </summary>
    private void SyncFrozenColumnWidths(SheetState state, bool mainToFrozen)
    {
        var main = state.MainGrid;
        if (state.FrozenGrid is not { } fg)
            return;
        if (_syncingColumnWidths)
            return;
        _syncingColumnWidths = true;
        try
        {
            var (src, dst) = mainToFrozen ? (main, fg) : (fg, main);
            for (var i = 0; i < src.Columns.Count && i < dst.Columns.Count; i++)
            {
                var w = src.Columns[i].ActualWidth;
                if (w > 0 && Math.Abs(dst.Columns[i].ActualWidth - w) > 0.5)
                {
                    dst.Columns[i].Width = new DataGridLength(w);
                }
            }
            if (fg.FrozenColumnCount != main.FrozenColumnCount)
                fg.FrozenColumnCount = main.FrozenColumnCount;
        }
        finally
        {
            _syncingColumnWidths = false;
        }
    }

    /// <summary>在 DataGrid 可视树里找内部 ScrollViewer（横滚同步用，结果缓存到 state.FrozenScroll）。</summary>
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

    private void OnUnfreezeClick(object sender, RoutedEventArgs e)
    {
        if (CurrentSheetState is not { } state)
            return;
        state.MainGrid.FrozenColumnCount = 0;
        state.FrozenColumns = 0;
        ApplyFreezeColumnDivider(state, 0);
        RemoveFrozenRows(state); // 拆冻结行 grid + 主 grid 恢复 VirtualizingSortableView
        if (state.FilePath is not null)
            FreezeConfig.ClearFreeze(Path.GetFileName(state.FilePath), state.SheetName);
        StatusText.Text = "已取消冻结";
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
        {
            _spotlightGrid.SelectedCellsChanged -= OnSpotlightSelectionChanged;
            _spotlightGrid.RemoveHandler(
                ScrollViewer.ScrollChangedEvent,
                (ScrollChangedEventHandler)OnSpotlightScrollChanged
            );
        }
        ClearSpotlight();

        _spotlightGrid = CurrentMainGrid;
        if (_spotlightGrid is not null)
        {
            _spotlightGrid.SelectedCellsChanged += OnSpotlightSelectionChanged;
            _spotlightGrid.AddHandler(
                ScrollViewer.ScrollChangedEvent,
                (ScrollChangedEventHandler)OnSpotlightScrollChanged
            );
            ApplySpotlight();
        }
    }

    private void OnSpotlightScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        // 滚动后重新触发聚光灯（虚拟化下可见行变了，要重新高亮）
        ApplySpotlight();
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
    /// 冻结行模式下同时处理冻结行 grid 的列高亮。
    /// </summary>
    private void HighlightSpotlight(
        IReadOnlySet<int> selectedRows,
        IReadOnlySet<int> selectedColumns
    )
    {
        if (_spotlightGrid is null)
            return;

        // 主 grid 高亮
        HighlightSpotlightInGrid(_spotlightGrid, selectedRows, selectedColumns);

        // 冻结行 grid 同步列高亮（行高亮不需要，冻结行不参与选中）
        if (CurrentSheetState?.FrozenGrid is { } frozenGrid)
            HighlightSpotlightInGrid(frozenGrid, new HashSet<int>(), selectedColumns);
    }

    private static void HighlightSpotlightInGrid(
        DataGrid grid,
        IReadOnlySet<int> selectedRows,
        IReadOnlySet<int> selectedColumns
    )
    {
        // 找选中区域的行列范围
        if (selectedRows.Count is 0 || selectedColumns.Count is 0)
            return;

        var minRow = selectedRows.Min();
        var maxRow = selectedRows.Max();
        var minCol = selectedColumns.Min();
        var maxCol = selectedColumns.Max();

        for (var i = 0; i < grid.Items.Count; i++)
        {
            if (grid.ItemContainerGenerator.ContainerFromIndex(i) is not DataGridRow rowContainer)
                continue;

            foreach (var col in grid.Columns)
            {
                if (col.GetCellContent(rowContainer)?.Parent is not DataGridCell cell)
                    continue;

                var isSelRow = selectedRows.Contains(i);
                var isSelCol = selectedColumns.Contains(col.DisplayIndex);

                // 选中行列的其他单元格：半透明背景色
                if (!isSelRow || !isSelCol)
                {
                    if ((isSelRow || isSelCol) && cell.Background != DirtyCellBrush)
                        cell.Background = SpotlightRowColBrush;
                    continue;
                }

                // 选中区域内部单元格：只在边缘画亮黄色边框（外框效果）
                var isTopEdge = i == minRow;
                var isBottomEdge = i == maxRow;
                var isLeftEdge = col.DisplayIndex == minCol;
                var isRightEdge = col.DisplayIndex == maxCol;

                var thickness = new Thickness(
                    isLeftEdge ? 2 : 0,
                    isTopEdge ? 2 : 0,
                    isRightEdge ? 2 : 0,
                    isBottomEdge ? 2 : 0
                );

                // 只在边缘画边框，避免每个单元格都框
                if (isTopEdge || isBottomEdge || isLeftEdge || isRightEdge)
                {
                    cell.BorderBrush = SpotlightBorderBrush;
                    cell.BorderThickness = thickness;
                }
            }
        }
    }

    private void ClearSpotlight()
    {
        if (_spotlightGrid is null)
            return;
        ClearSpotlightInGrid(_spotlightGrid);

        // 冻结行 grid 也清除
        if (CurrentSheetState?.FrozenGrid is { } frozenGrid)
            ClearSpotlightInGrid(frozenGrid);
    }

    private static void ClearSpotlightInGrid(DataGrid grid)
    {
        for (var i = 0; i < grid.Items.Count; i++)
        {
            if (grid.ItemContainerGenerator.ContainerFromIndex(i) is not DataGridRow rowContainer)
                continue;
            foreach (var col in grid.Columns)
            {
                if (col.GetCellContent(rowContainer)?.Parent is not DataGridCell cell)
                    continue;
                if (cell.BorderBrush == SpotlightBorderBrush)
                {
                    cell.ClearValue(DataGridCell.BorderBrushProperty);
                    cell.ClearValue(DataGridCell.BorderThicknessProperty);
                }
                if (cell.Background == SpotlightRowColBrush)
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
        // P3.2 后 grid 行项是 RowView（不再是 DataRowView）——旧的 DataRowView 判断恒失败使本功能失效，此处修正。
        if (grid.CurrentCell.Column is null || grid.CurrentCell.Item is not RowView view)
            return;
        var colIndex = grid.CurrentCell.Column.DisplayIndex;
        var value = view[colIndex] ?? string.Empty;
        // 列头现为纯列名字符串（BuildDataColumns 设 Header=columnName）。
        var header = grid.CurrentCell.Column.Header?.ToString() ?? "?";
        var rowIdx = view.RowIndex;
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
                view[colIndex] = tb.Text; // RowView 索引器 → SetCell 标脏
                // #6：关窗写回也进撤销栈（一次 Ctrl+Z 撤回这次编辑）。
                state.UndoStack.Push(
                    new CellBatchAction([
                        new CellEditRecord(view.RowIndex, colIndex, value, tb.Text),
                    ])
                );
                state.RedoStack.Clear();
                MarkDirty(grid, view, colIndex);
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
    /// 在 UI 线程从 ColumnStore 读出 string[,] 快照，供后台搜索（SearchSnapshots 保持纯 string[,] 逻辑不变）。
    /// </summary>
    private (string, string, string[,]) BuildSearchSnapshot(SheetState state)
    {
        var store = state.Store;
        var rows = store.RowCount;
        var cols = store.ColumnCount;
        var data = new string[rows, cols];
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < cols; c++)
            data[r, c] = store.GetCell(r, c) ?? string.Empty;
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
            StatusText.Text = $"已加载全部 {selectedState.TotalRows} 行";
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
        ColumnStore store,
        Dictionary<(int Row, int Col), string> comments,
        string? filePath,
        int totalRows
    )
    {
        public string SheetName { get; } = sheetName;

        // P3.2: 列式存储 + 虚拟化视图取代 DataTable。整表一次性加载进 Store（无后台逐行灌数据），
        // View 是 DataGrid 的 ItemsSource（按需物化 RowView，不预建整表对象树）。
        // 用 VirtualizingSortableView：支持 ApplyFilter/ClearFilter（列头筛选），不整表复制。
        public ColumnStore Store { get; } = store;
        public VirtualizingSortableView View { get; } = new(store);
        public Dictionary<(int Row, int Col), string> Comments { get; } = comments;
        public string? FilePath { get; set; } = filePath;

        // #1：主 grid 的 DataGridExtensions 筛选适配器（冻结时设 BasePredicate=row>=n 让主区数据行仍可筛）。
        public ColumnStoreFilterAdapter? Filter { get; set; }

        // Store 一次性全量加载，TotalRows == LoadedRows == Store.RowCount（保留字段供状态栏/兼容显示）。
        public int TotalRows { get; set; } = totalRows;
        public int LoadedRows { get; set; } = totalRows;
        public CancellationTokenSource? LoadCts { get; set; }

        // #6：撤销/重做单元 = IUndoableAction，覆盖单格编辑、多格粘贴、增删行、增删列。
        // 单格编辑/粘贴用 CellBatchAction（一批 CellEditRecord），结构操作用 InsertRow/DeleteRows/InsertColumn Action，
        // 一次 Ctrl+Z 整体撤销一次动作。重放逻辑在 MainWindowUndo（IUndoableAction 重载）。
        public Stack<IUndoableAction> UndoStack { get; } = new();
        public Stack<IUndoableAction> RedoStack { get; } = new();

        // ── 冻结窗格（P3.2 后 FrozenRows 双 grid 方案暂时退化，见 status.md；FrozenColumns 仍走原生 FrozenColumnCount）──
        public DataGrid MainGrid { get; set; } = null!;
        public Grid? Panel { get; set; }
        public DataGrid? FrozenGrid { get; set; }
        public ScrollViewer? FrozenScroll { get; set; }
        public ScrollViewer? MainScroll { get; set; }
        public int FrozenRows { get; set; }
        public int FrozenColumns { get; set; }

        // #P9-3：冻结列分界竖线——一条独立的 Border，跨 panel 两行（冻结行 grid + 主 grid）绘制，
        // X 位置 = 行号头宽 + 前 N 个冻结列实际宽之和。用单条贯穿线取代"在两个 grid 各自列上画边框
        // 再指望自然对齐"的旧方案（P7/P8 两轮都因跨 grid 坐标不一致而断层）。
        public Border? FrozenDivider { get; set; }

        // #P9-3：分界线 X 重定位的 LayoutUpdated 处理器（存字段以便重挂时先解绑，避免重复订阅）。
        public EventHandler? DividerLayoutHandler { get; set; }
    }
}
