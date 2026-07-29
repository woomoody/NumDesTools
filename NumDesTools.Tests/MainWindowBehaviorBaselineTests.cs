using System.Collections;
using System.Data;
using System.Reflection;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 基线特征化测试（characterization tests）：锁定 <c>MainWindow.xaml.cs</c> 当前**基于
/// DataTable** 的可观察行为，作为 P3.2（把加载/绑定/编辑/撤销/搜索切到 ColumnStore +
/// 虚拟化视图）大改动的安全网。改完后重跑这批测试，行为若变立刻可见。
///
/// <para>
/// 只覆盖能抽出纯逻辑的部分：<c>MainWindow</c> 是 WPF Window，构造需 XAML 上下文，
/// 大量方法耦合 DataGrid/Dispatcher。以下方法虽是 <c>private static</c> 但**不碰任何 WPF 控件**，
/// 通过反射调用（沿用 <see cref="OoxmlLazyReaderTests"/> 直接加载 XlsxEditor.dll 反射私有成员的既有约定）：
/// </para>
/// <list type="bullet">
///   <item><description><c>BuildAllSheetsLazy(string)</c> — DataTable 首屏加载路径（行/列/前 200 行）。</description></item>
///   <item><description><c>AddRawRow(DataTable, comments, RawRow)</c> — 单行落表 + 批注归集。</description></item>
///   <item><description><c>GetExcelColumnName(int)</c> — 1→A/26→Z/27→AA 列名生成。</description></item>
///   <item><description><c>SearchSnapshots(targets, keyword)</c> — 不区分大小写子串搜索 + 500 上限。</description></item>
///   <item><description><c>CellEditRecord</c>（internal record）+ Undo/Redo 的 DataTable 状态迁移契约。</description></item>
/// </list>
///
/// <para>
/// **覆盖不到**（需真实进程 + WPF 渲染验证，见 <c>docs/xlsx-editor-manual-qa-baseline.md</c>）：
/// DataGrid 实际渲染、双 DataGrid 冻结行视觉同步、聚光灯高亮、tab 切换、脏标记刷新等。
/// </para>
/// </summary>
public sealed class MainWindowBehaviorBaselineTests(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    /// <summary>Item.xlsx 首个工作表的真实行数：动态取自 EPPlus dimension（上游刷表会增行，硬编码基线脆断）。</summary>
    private static readonly int ItemFirstSheetTotalRows = ReadDimension().Rows;

    /// <summary>Item.xlsx 首个工作表的真实列数（含末尾有数据列）：动态取自 EPPlus dimension。</summary>
    private static readonly int ItemFirstSheetTotalCols = ReadDimension().Cols;

    /// <summary>BuildAllSheetsLazy 硬编码的首屏行数（firstScreenRows = 200）。</summary>
    private const int FirstScreenRows = 200;

    private static readonly Type MainWindowType = LoadXlsxEditorType(
        "NumDesTools.XlsxEditor.MainWindow"
    );

    static MainWindowBehaviorBaselineTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    private static (int Rows, int Cols) ReadDimension()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var dim = package.Workbook.Worksheets[0].Dimension;
        return (dim.End.Row, dim.End.Column);
    }

    // ─────────────────────────────────────────────────────────────────
    //  GetExcelColumnName — 1-based 列序号 → Excel 列名
    // ─────────────────────────────────────────────────────────────────

    [Theory]
    [InlineData(1, "A")]
    [InlineData(2, "B")]
    [InlineData(25, "Y")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(28, "AB")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(85, "CG")] // Item.xlsx 第 85 列
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")] // Excel 最大列
    public void GetExcelColumnName_MapsOneBasedIndexToExcelName(int col, string expected)
    {
        var actual = InvokeGetExcelColumnName(col);

        Assert.Equal(expected, actual);
    }

    // ─────────────────────────────────────────────────────────────────
    //  BuildAllSheetsLazy — DataTable 首屏加载路径（被 P3.2 替换的核心）
    // ─────────────────────────────────────────────────────────────────

    [Fact]
    public void BuildAllSheetsLazy_Item_ProducesAtLeastOneSheet()
    {
        var sheets = InvokeBuildAllSheetsLazy(ItemPath);

        Assert.NotEmpty(sheets);
        // 交叉验证：跳过 # 前缀后的 sheet 数应与 OoxmlLazyReader 报告一致
        // OoxmlLazyReader 是 internal，沿用 OoxmlLazyReaderTests 反射调用其静态方法
        var expectedNames = ReadSheetNamesViaReflection(ItemPath)
            .Where(name => !name.StartsWith('#'))
            .ToList();
        var builtNames = sheets.Select(s => s.Name).ToList();

        output.WriteLine(
            $"[BuildAllSheetsLazy] sheets={sheets.Count}: {string.Join(", ", builtNames)}"
        );

        // BuildAllSheetsLazy 额外跳过 dimension 为空的 sheet，故 built ⊆ expected
        Assert.All(builtNames, name => Assert.Contains(name, expectedNames));
    }

    [Fact]
    public void BuildAllSheetsLazy_Item_FirstSheet_HasExpectedShape()
    {
        var sheets = InvokeBuildAllSheetsLazy(ItemPath);
        var first = sheets[0];

        // 列数 = dimension 报告的真实列数（所有列都建成 string 列）
        Assert.Equal(ItemFirstSheetTotalCols, first.Table.Columns.Count);
        // TotalRows = dimension 报告的真实行数（不是首屏行数）
        Assert.Equal(ItemFirstSheetTotalRows, first.TotalRows);
        // 首屏只加载前 200 行：Table.Rows.Count == LoadedRows == 200
        Assert.Equal(FirstScreenRows, first.LoadedRows);
        Assert.Equal(FirstScreenRows, first.Table.Rows.Count);

        // 列名走 GetExcelColumnName：第 1 列 A、第 26 列 Z、第 27 列 AA、末列 CG
        Assert.Equal("A", first.Table.Columns[0].ColumnName);
        Assert.Equal("Z", first.Table.Columns[25].ColumnName);
        Assert.Equal("AA", first.Table.Columns[26].ColumnName);
        Assert.Equal("CG", first.Table.Columns[ItemFirstSheetTotalCols - 1].ColumnName);

        // 所有列都是 string 类型
        Assert.All(
            first.Table.Columns.Cast<DataColumn>(),
            column => Assert.Equal(typeof(string), column.DataType)
        );

        output.WriteLine(
            $"[BuildAllSheetsLazy] first sheet '{first.Name}': "
                + $"TotalRows={first.TotalRows}, LoadedRows={first.LoadedRows}, "
                + $"Cols={first.Table.Columns.Count}, firstCol='{first.Table.Columns[0].ColumnName}', "
                + $"lastCol='{first.Table.Columns[^1].ColumnName}'"
        );
    }

    [Fact]
    public void BuildAllSheetsLazy_Item_FirstSheet_FirstRowsMatchEpPlus()
    {
        var sheets = InvokeBuildAllSheetsLazy(ItemPath);
        var table = sheets[0].Table;

        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[0];

        // BuildAllSheetsLazy 的 DataTable 是 0-based 行/列；EPPlus 是 1-based。
        // 行标题行/列名行也在其中（未跳前 N 行，与 ColumnStore 加载路径同样保留原始行）。
        (int Row, int Col)[] samples =
        [
            (0, 0), // Excel(1,1) = "#"（标题行第一列）
            (1, 1), // Excel(2,2) = "id"（列名行）
            (4, 1), // Excel(5,2) = 首个数据行 id
            (99, 1), // Excel(100,2)
            (199, 1), // Excel(200,2) — 首屏最后一行仍在范围内
        ];

        foreach (var (row, col) in samples)
        {
            var expected = sheet.Cells[row + 1, col + 1].Value?.ToString() ?? string.Empty;
            var actual = table.Rows[row][col]?.ToString();
            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public void BuildAllSheetsLazy_Item_FirstSheet_KnownLiteralValues()
    {
        var sheets = InvokeBuildAllSheetsLazy(ItemPath);
        var table = sheets[0].Table;

        // 与编码无关的纯 ASCII/数字锚点，独立于 EPPlus 的第二重基线
        Assert.Equal("#", table.Rows[0][0]?.ToString());
        Assert.Equal("id", table.Rows[1][1]?.ToString());
        Assert.Equal("11010001", table.Rows[4][1]?.ToString());
        Assert.Equal("13010504", table.Rows[99][1]?.ToString());
    }

    // ─────────────────────────────────────────────────────────────────
    //  AddRawRow — 单行落表 + 批注归集
    // ─────────────────────────────────────────────────────────────────

    [Fact]
    public void AddRawRow_MissingColumnsBecomeEmptyString()
    {
        var table = MakeStringTable("A", "B", "C");
        var comments = new Dictionary<(int Row, int Col), string>();
        // 只给 A、C 列，B 缺失
        var raw = new RawRow(
            5,
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["A"] = "alpha",
                ["C"] = "gamma",
            },
            []
        );

        InvokeAddRawRow(table, comments, raw);

        Assert.Single(table.Rows);
        Assert.Equal("alpha", table.Rows[0]["A"]);
        Assert.Equal(string.Empty, table.Rows[0]["B"]); // 缺失列 → 空字符串
        Assert.Equal("gamma", table.Rows[0]["C"]);
    }

    [Fact]
    public void AddRawRow_ColumnsFollowTableOrder_NotRawCellsOrder()
    {
        var table = MakeStringTable("A", "B");
        var comments = new Dictionary<(int Row, int Col), string>();
        // RawRow.Cells 的插入顺序与列顺序相反，仍应按 table 列名匹配
        var raw = new RawRow(
            1,
            new Dictionary<string, string>(StringComparer.Ordinal) { ["B"] = "b", ["A"] = "a" },
            []
        );

        InvokeAddRawRow(table, comments, raw);

        Assert.Equal("a", table.Rows[0][0]);
        Assert.Equal("b", table.Rows[0][1]);
    }

    [Fact]
    public void AddRawRow_PropagatesCommentsIntoSharedMap()
    {
        var table = MakeStringTable("A");
        var comments = new Dictionary<(int Row, int Col), string>();
        var raw = new RawRow(
            3,
            new Dictionary<string, string>(StringComparer.Ordinal) { ["A"] = "x" },
            new Dictionary<(int Row, int Col), string> { [(3, 1)] = "备注" }
        );

        InvokeAddRawRow(table, comments, raw);

        Assert.Single(comments);
        Assert.Equal("备注", comments[(3, 1)]);
    }

    // ─────────────────────────────────────────────────────────────────
    //  SearchSnapshots — 不区分大小写子串 + 500 上限 + 1-based 行列
    // ─────────────────────────────────────────────────────────────────

    [Fact]
    public void SearchSnapshots_CaseInsensitiveSubstringMatch()
    {
        var data = new string[,]
        {
            { "Apple", "banana" },
            { "APPLE PIE", "cherry" },
        };
        var targets = MakeTargets(("File.xlsx", "Sheet1", data));

        var results = InvokeSearchSnapshots(targets, "apple");

        Assert.Equal(2, results.Count);
        // 1-based 行/列 + Excel 列名
        Assert.Equal((1, 1, "A"), (results[0].Row, results[0].Col, results[0].ColumnName));
        Assert.Equal((2, 1, "A"), (results[1].Row, results[1].Col, results[1].ColumnName));
        Assert.Equal("Apple", results[0].Value);
        Assert.Equal("APPLE PIE", results[1].Value);
    }

    [Fact]
    public void SearchSnapshots_ReportsFileAndSheetAndExcelColumnName()
    {
        var data = new string[,]
        {
            { "x", "y", "target" }, // col index 2 → Excel "C"
        };
        var targets = MakeTargets(("Book.xlsx", "MySheet", data));

        var results = InvokeSearchSnapshots(targets, "target");

        var hit = Assert.Single(results);
        Assert.Equal("Book.xlsx", hit.FileName);
        Assert.Equal("MySheet", hit.SheetName);
        Assert.Equal(1, hit.Row);
        Assert.Equal(3, hit.Col);
        Assert.Equal("C", hit.ColumnName);
    }

    [Fact]
    public void SearchSnapshots_NoMatch_ReturnsEmpty()
    {
        var data = new string[,]
        {
            { "foo", "bar" },
        };
        var targets = MakeTargets(("F.xlsx", "S", data));

        var results = InvokeSearchSnapshots(targets, "zzz");

        Assert.Empty(results);
    }

    [Fact]
    public void SearchSnapshots_CapsAtFiveHundredResults()
    {
        // 600 行 × 1 列全部命中 → 上限 500
        var data = new string[600, 1];
        for (var r = 0; r < 600; r++)
            data[r, 0] = "hit";
        var targets = MakeTargets(("F.xlsx", "S", data));

        var results = InvokeSearchSnapshots(targets, "hit");

        Assert.Equal(500, results.Count);
    }

    [Fact]
    public void SearchSnapshots_SpansMultipleTargetsInOrder()
    {
        var targets = MakeTargets(
            (
                "A.xlsx",
                "S1",
                new string[,]
                {
                    { "k" },
                }
            ),
            (
                "B.xlsx",
                "S2",
                new string[,]
                {
                    { "k" },
                }
            )
        );

        var results = InvokeSearchSnapshots(targets, "k");

        Assert.Equal(2, results.Count);
        Assert.Equal("A.xlsx", results[0].FileName);
        Assert.Equal("B.xlsx", results[1].FileName);
    }

    [Fact]
    public void SearchSnapshots_Item_KnownKeywordCountIsStable()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var rows = store.RowCount;
        var cols = store.ColumnCount;
        var data = new string[rows, cols];
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < cols; c++)
            data[r, c] = store.GetCell(r, c) ?? string.Empty;
        var targets = MakeTargets(("Item.xlsx", "Item", data));

        // 实测基线：关键字 "11010001" 在 Item.xlsx 出现 2 次，均在第 5 行（1-based）——
        // B5（id 列，Col 2）和 J5（Col 10，被别处引用）。搜索按行优先、列递增顺序返回，
        // 故 B5 在前、J5 在后。此为当前行为快照，P3.2 改完后须一致。
        var results = InvokeSearchSnapshots(targets, "11010001");

        Assert.Equal(2, results.Count);
        Assert.All(results, hit => Assert.Equal(5, hit.Row)); // 均在第 5 行（1-based）
        Assert.All(results, hit => Assert.Equal("11010001", hit.Value));

        // 第一命中 B5：id 列
        Assert.Equal(2, results[0].Col); // 1-based → Excel B
        Assert.Equal("B", results[0].ColumnName);
        // 第二命中 J5：同值出现在第 10 列
        Assert.Equal(10, results[1].Col); // 1-based → Excel J
        Assert.Equal("J", results[1].ColumnName);

        output.WriteLine(
            $"[SearchSnapshots] Item.xlsx '11010001' → {results.Count} hits: "
                + string.Join(
                    ", ",
                    results.Select(hit => $"{hit.ColumnName}{hit.Row}='{hit.Value}'")
                )
        );
    }

    // ─────────────────────────────────────────────────────────────────
    //  CellEditRecord + Undo/Redo 的 DataTable 状态迁移契约
    // ─────────────────────────────────────────────────────────────────

    [Fact]
    public void CellEditRecord_CarriesRowColOldNew()
    {
        var record = MakeCellEditRecord(3, 4, "old", "new");

        Assert.Equal(3, ReadRecordMember<int>(record, "Row"));
        Assert.Equal(4, ReadRecordMember<int>(record, "Col"));
        Assert.Equal("old", ReadRecordMember<object?>(record, "OldValue"));
        Assert.Equal("new", ReadRecordMember<string>(record, "NewValue"));
    }

    /// <summary>
    /// 特征化 MainWindow 的撤销契约：编辑压栈 (Row,Col,OldValue,NewValue)；撤销时
    /// 把单元格写回 OldValue，并把撤销前的当前值压入 RedoStack；重做时写 NewValue
    /// 并压回 UndoStack。此处用与生产代码同构的 DataTable + Stack&lt;CellEditRecord&gt;
    /// 复现该状态迁移，锁定 P3.2 切到 ColumnStore 后必须保持一致的可观察结果。
    /// </summary>
    [Fact]
    public void UndoRedo_Contract_RestoresOldValueThenRedoRestoresNewValue()
    {
        var table = MakeStringTable("A", "B");
        var row = table.NewRow();
        row[0] = "orig";
        row[1] = "keep";
        table.Rows.Add(row);

        var undo = new Stack<object>();
        var redo = new Stack<object>();

        // ── 用户编辑 A0: orig → edited（等价于 CellEditEnding 压栈）
        var oldValue = table.Rows[0][0];
        table.Rows[0][0] = "edited";
        undo.Push(MakeCellEditRecord(0, 0, oldValue, "edited"));
        redo.Clear();

        Assert.Equal("edited", table.Rows[0][0]);
        Assert.Single(undo);
        Assert.Empty(redo);

        // ── 撤销：pop undo，写回 OldValue，push 当前值到 redo
        var undoRec = undo.Pop();
        var undoR = ReadRecordMember<int>(undoRec, "Row");
        var undoC = ReadRecordMember<int>(undoRec, "Col");
        var currentBeforeUndo = table.Rows[undoR][undoC];
        table.Rows[undoR][undoC] = ReadRecordMember<object?>(undoRec, "OldValue");
        redo.Push(
            MakeCellEditRecord(undoR, undoC, currentBeforeUndo, currentBeforeUndo?.ToString() ?? "")
        );

        Assert.Equal("orig", table.Rows[0][0]); // 恢复原值
        Assert.Empty(undo);
        Assert.Single(redo);

        // ── 重做：pop redo，写 NewValue，push 当前值到 undo
        var redoRec = redo.Pop();
        var redoR = ReadRecordMember<int>(redoRec, "Row");
        var redoC = ReadRecordMember<int>(redoRec, "Col");
        var currentBeforeRedo = table.Rows[redoR][redoC];
        table.Rows[redoR][redoC] = ReadRecordMember<string>(redoRec, "NewValue");
        undo.Push(
            MakeCellEditRecord(redoR, redoC, currentBeforeRedo, currentBeforeRedo?.ToString() ?? "")
        );

        Assert.Equal("edited", table.Rows[0][0]); // 重做回到编辑值
        Assert.Single(undo);
        Assert.Empty(redo);

        // 相邻列不受影响
        Assert.Equal("keep", table.Rows[0][1]);
    }

    [Fact]
    public void UndoRedo_Contract_NewEditClearsRedoStack()
    {
        var redo = new Stack<object>();
        redo.Push(MakeCellEditRecord(0, 0, "a", "b"));
        Assert.Single(redo);

        // 任意新编辑发生时清空 redo（MainWindow.CellEditEnding: RedoStack.Clear()）
        redo.Clear();

        Assert.Empty(redo);
    }

    // ─────────────────────────────────────────────────────────────────
    //  反射辅助（沿用 OoxmlLazyReaderTests 加载 dll 反射私有成员的约定）
    // ─────────────────────────────────────────────────────────────────

    private static string InvokeGetExcelColumnName(int col) =>
        (string)
            MainWindowType
                .GetMethod("GetExcelColumnName", BindingFlags.NonPublic | BindingFlags.Static)!
                .Invoke(null, [col])!;

    private static List<(
        string Name,
        DataTable Table,
        Dictionary<(int Row, int Col), string> Comments,
        int TotalRows,
        int LoadedRows
    )> InvokeBuildAllSheetsLazy(string path)
    {
        var raw = MainWindowType
            .GetMethod("BuildAllSheetsLazy", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [path]);
        var list =
            new List<(string, DataTable, Dictionary<(int Row, int Col), string>, int, int)>();
        foreach (var tuple in (IEnumerable)raw!)
        {
            var type = tuple.GetType();
            list.Add(
                (
                    (string)type.GetField("Item1")!.GetValue(tuple)!,
                    (DataTable)type.GetField("Item2")!.GetValue(tuple)!,
                    (Dictionary<(int Row, int Col), string>)
                        type.GetField("Item3")!.GetValue(tuple)!,
                    (int)type.GetField("Item4")!.GetValue(tuple)!,
                    (int)type.GetField("Item5")!.GetValue(tuple)!
                )
            );
        }

        return list;
    }

    private static void InvokeAddRawRow(
        DataTable table,
        Dictionary<(int Row, int Col), string> comments,
        RawRow raw
    ) =>
        MainWindowType
            .GetMethod("AddRawRow", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [table, comments, raw]);

    private static List<string> ReadSheetNamesViaReflection(string path)
    {
        var readerType = LoadXlsxEditorType("NumDesTools.XlsxEditor.OoxmlLazyReader");
        var raw = readerType
            .GetMethod("ReadSheetNames", BindingFlags.Public | BindingFlags.Static)!
            .Invoke(null, [path]);
        return (List<string>)raw!;
    }

    private List<SearchHit> InvokeSearchSnapshots(object targets, string keyword)
    {
        var raw = MainWindowType
            .GetMethod("SearchSnapshots", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [targets, keyword]);
        var hits = new List<SearchHit>();
        foreach (var item in (IEnumerable)raw!)
        {
            hits.Add(
                new SearchHit(
                    ReadRecordMember<string>(item, "FileName"),
                    ReadRecordMember<string>(item, "SheetName"),
                    ReadRecordMember<int>(item, "Row"),
                    ReadRecordMember<int>(item, "Col"),
                    ReadRecordMember<string>(item, "Value"),
                    ReadRecordMember<string>(item, "ColumnName")
                )
            );
        }

        return hits;
    }

    /// <summary>
    /// 构造 SearchSnapshots 需要的强类型 targets：
    /// <c>List&lt;(string FileName, string SheetName, string[,] Data)&gt;</c>。
    /// 反射调用要求实参类型与形参精确一致，故直接用值元组列表（编译期已是正确类型）。
    /// </summary>
    private static List<(string FileName, string SheetName, string[,] Data)> MakeTargets(
        params (string FileName, string SheetName, string[,] Data)[] items
    ) => [.. items];

    private static object MakeCellEditRecord(int row, int col, object? oldValue, string newValue)
    {
        var type = LoadXlsxEditorType("NumDesTools.XlsxEditor.CellEditRecord");
        return Activator.CreateInstance(type, row, col, oldValue, newValue)!;
    }

    private static T ReadRecordMember<T>(object target, string memberName)
    {
        var value =
            target.GetType().GetProperty(memberName)?.GetValue(target)
            ?? target.GetType().GetField(memberName)?.GetValue(target);
        if (value is null)
            return default!;
        return (T)value;
    }

    private static DataTable MakeStringTable(params string[] columnNames)
    {
        var table = new DataTable();
        foreach (var name in columnNames)
            table.Columns.Add(name, typeof(string));
        return table;
    }

    private static Type LoadXlsxEditorType(string fullName)
    {
        var assemblyPath = Path.GetFullPath(
            Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "NumDesTools.XlsxEditor",
                "bin",
                "Debug",
                "net9.0-windows",
                "NumDesTools.XlsxEditor.dll"
            )
        );
        var assembly = Assembly.LoadFrom(assemblyPath);
        return assembly.GetType(fullName, throwOnError: true)!;
    }

    /// <summary>SearchResultItem 的可测试镜像（原类型 private nested record，反射读字段）。</summary>
    private sealed record SearchHit(
        string FileName,
        string SheetName,
        int Row,
        int Col,
        string Value,
        string ColumnName
    );
}
