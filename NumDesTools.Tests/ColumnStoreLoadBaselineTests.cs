using System.Collections;
using System.Reflection;
using NumDesTools.XlsxEditor;
using OfficeOpenXml;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// P3.2 新加载路径的基线测试：锁定 <c>MainWindow.BuildStoresFromExcel</c>（用
/// <see cref="ColumnStoreExcelLoader"/> 一次性构建 <see cref="ColumnStore"/>）的可观察行为。
///
/// <para>
/// 这批测试是旧 <see cref="MainWindowBehaviorBaselineTests"/> 中 <c>BuildAllSheetsLazy</c>（DataTable
/// 首屏加载）4 个测试的<b>等价替代</b>——旧测试断言的是 DataTable 结构/值，新测试断言 ColumnStore
/// 结构/值，验证的是<b>同一件事</b>（加载 Item.xlsx 后的行数/列数/列名/单元格值）。旧测试仍保留并通过
/// （<c>BuildAllSheetsLazy</c> 方法未删，作为纯逻辑仍被锁定），但它已不是 UI 的加载路径；UI 现在走
/// <c>BuildStoresFromExcel</c>，故此处补齐等价基线。
/// </para>
///
/// <para>旧测试 → 新测试 对应关系：</para>
/// <list type="bullet">
///   <item><description><c>BuildAllSheetsLazy_Item_ProducesAtLeastOneSheet</c> → <see cref="BuildStoresFromExcel_Item_ProducesAtLeastOneSheet"/></description></item>
///   <item><description><c>BuildAllSheetsLazy_Item_FirstSheet_HasExpectedShape</c>（DataTable.Rows.Count/Columns.Count/列名）→ <see cref="BuildStoresFromExcel_Item_FirstSheet_HasExpectedShape"/>（Store.RowCount/ColumnCount/ColumnNames）</description></item>
///   <item><description><c>BuildAllSheetsLazy_Item_FirstSheet_FirstRowsMatchEpPlus</c> → <see cref="BuildStoresFromExcel_Item_FirstSheet_MatchesEpPlus"/>（改为跨越全表抽样，因新路径一次性全量加载，不再有 200 行首屏上限）</description></item>
///   <item><description><c>BuildAllSheetsLazy_Item_FirstSheet_KnownLiteralValues</c> → <see cref="BuildStoresFromExcel_Item_FirstSheet_KnownLiteralValues"/></description></item>
/// </list>
///
/// <para>
/// **关键差异（诚实说明）**：旧路径首屏只加载 200 行（<c>LoadedRows==200</c>），后台再逐行灌满；
/// 新路径 <c>BuildStoresFromExcel</c> 一次性把全表读进 Store（<c>Store.RowCount==TotalRows==65105</c>），
/// 没有"首屏行数"概念。因此旧的 <c>Assert.Equal(200, first.LoadedRows)</c> 在新实现下不适用——
/// 新实现的等价断言是"加载后 Store 立即含全部 65105 行"，见 <see cref="BuildStoresFromExcel_Item_FirstSheet_HasExpectedShape"/>。
/// 这不是回归，是 P3 消灭"首屏 200 + 后台逐行"卡顿根因的预期结果。
/// </para>
/// </summary>
public sealed class ColumnStoreLoadBaselineTests(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    // 行数动态取自 EPPlus dimension（真实文件被上游刷表增行，硬编码基线会脆断）；列数 85 稳定但也动态取。
    private static readonly int ItemFirstSheetTotalRows = ReadDimension().Rows;
    private static readonly int ItemFirstSheetTotalCols = ReadDimension().Cols;

    private static (int Rows, int Cols) ReadDimension()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var dim = package.Workbook.Worksheets[0].Dimension;
        return (dim.End.Row, dim.End.Column);
    }

    private static readonly Type MainWindowType = LoadXlsxEditorType(
        "NumDesTools.XlsxEditor.MainWindow"
    );

    static ColumnStoreLoadBaselineTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    [Fact]
    public void BuildStoresFromExcel_Item_ProducesAtLeastOneSheet()
    {
        var sheets = InvokeBuildStoresFromExcel(ItemPath);

        Assert.NotEmpty(sheets);
        // 跳过 # 前缀后的 sheet 数应与 OoxmlLazyReader 报告一致（子集：还跳过空 sheet）
        var expectedNames = ReadSheetNamesViaReflection(ItemPath)
            .Where(name => !name.StartsWith('#'))
            .ToList();
        var builtNames = sheets.Select(s => s.Name).ToList();

        output.WriteLine(
            $"[BuildStoresFromExcel] sheets={sheets.Count}: {string.Join(", ", builtNames)}"
        );

        Assert.All(builtNames, name => Assert.Contains(name, expectedNames));
    }

    [Fact]
    public void BuildStoresFromExcel_Item_FirstSheet_HasExpectedShape()
    {
        var sheets = InvokeBuildStoresFromExcel(ItemPath);
        var first = sheets[0];
        var store = first.Store;

        // 列数 = dimension 报告的真实列数（等价旧 first.Table.Columns.Count == 85）
        Assert.Equal(ItemFirstSheetTotalCols, store.ColumnCount);
        // 新路径一次性全量加载：Store.RowCount == TotalRows == 65105（等价旧 first.TotalRows == 65105）。
        // 旧路径 first.Table.Rows.Count == first.LoadedRows == 200（首屏）；新路径无首屏概念，一次到位。
        Assert.Equal(ItemFirstSheetTotalRows, store.RowCount);
        Assert.Equal(ItemFirstSheetTotalRows, first.TotalRows);

        // 列名走 Excel 列名规则：第 1 列 A、第 26 列 Z、第 27 列 AA、末列 CG（等价旧 Table.Columns[i].ColumnName）
        Assert.Equal("A", store.ColumnNames[0]);
        Assert.Equal("Z", store.ColumnNames[25]);
        Assert.Equal("AA", store.ColumnNames[26]);
        Assert.Equal("CG", store.ColumnNames[ItemFirstSheetTotalCols - 1]);

        output.WriteLine(
            $"[BuildStoresFromExcel] first sheet '{first.Name}': "
                + $"RowCount={store.RowCount}, ColumnCount={store.ColumnCount}, "
                + $"firstCol='{store.ColumnNames[0]}', lastCol='{store.ColumnNames[^1]}'"
        );
    }

    [Fact]
    public void BuildStoresFromExcel_Item_FirstSheet_MatchesEpPlus()
    {
        var sheets = InvokeBuildStoresFromExcel(ItemPath);
        var store = sheets[0].Store;

        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[0];

        // Store 是 0-based 行/列；EPPlus 是 1-based。新路径一次性全量加载，故可抽样到很靠后的行
        // （旧测试只能抽到前 200 行首屏内）。
        (int Row, int Col)[] samples =
        [
            (0, 0), // Excel(1,1) = "#"
            (1, 1), // Excel(2,2) = "id"
            (4, 1), // Excel(5,2) = 首个数据行 id
            (99, 1), // Excel(100,2)
            (199, 1), // Excel(200,2)
            (1000, 1), // Excel(1001,2) — 远超旧首屏 200 行上限，验证全量加载
            (ItemFirstSheetTotalRows - 1, 1), // 末行
        ];

        foreach (var (row, col) in samples)
        {
            var expected = sheet.Cells[row + 1, col + 1].Value?.ToString() ?? string.Empty;
            var actual = store.GetCell(row, col) ?? string.Empty;
            Assert.Equal(expected, actual);
        }
    }

    [Fact]
    public void BuildStoresFromExcel_Item_FirstSheet_KnownLiteralValues()
    {
        var sheets = InvokeBuildStoresFromExcel(ItemPath);
        var store = sheets[0].Store;

        // 与旧 BuildAllSheetsLazy_..._KnownLiteralValues 完全相同的纯 ASCII/数字锚点
        Assert.Equal("#", store.GetCell(0, 0));
        Assert.Equal("id", store.GetCell(1, 1));
        Assert.Equal("11010001", store.GetCell(4, 1));
        Assert.Equal("13010504", store.GetCell(99, 1));
    }

    // ─────────────────────────────────────────────────────────────────
    //  编辑写回 + 撤销/重做（ColumnStore + RowView 版，等价旧
    //  MainWindowBehaviorBaselineTests.UndoRedo_Contract_...，验证 P3.2 切换后同一状态迁移）
    // ─────────────────────────────────────────────────────────────────

    /// <summary>
    /// 特征化 P3.2 编辑写回链路：DataGrid 列 <c>Binding "[col]"</c> → <see cref="RowView"/> 索引器 set →
    /// <see cref="ColumnStore.SetCell"/>（标脏）。等价旧 DataTable 路径里 <c>DataRowView[col]=value</c>。
    /// </summary>
    [Fact]
    public void RowViewIndexerSet_WritesBackToStore_AndMarksDirty()
    {
        var store = ColumnStore.Create(["A", "B"], initialRowCapacity: 1);
        store.AppendRow();
        store.SetCellQuiet(0, 0, "orig"); // 加载态：不脏
        store.SetCellQuiet(0, 1, "keep");
        Assert.False(store.IsDirty(0, 0));

        var view = new RowView(store, 0);

        // 编辑（等价 CellEditEnding: view[colIndex] = newValue）
        view[0] = "edited";

        Assert.Equal("edited", store.GetCell(0, 0));
        Assert.True(store.IsDirty(0, 0)); // 写回后标脏（P4 增量写回依赖）
        Assert.Equal("keep", store.GetCell(0, 1)); // 邻列不受影响
    }

    /// <summary>
    /// 撤销/重做契约在 ColumnStore 下等价成立：撤销写回 OldValue，重做写 NewValue，
    /// 用与生产代码同构的 <c>Stack&lt;(Row,Col,Old,New)&gt;</c> + ColumnStore 复现（对应旧
    /// UndoRedo_Contract_RestoresOldValueThenRedoRestoresNewValue 里的 DataTable 版本；
    /// CellEditRecord 为 internal，此处用等价元组承载相同字段）。
    /// </summary>
    [Fact]
    public void UndoRedo_Contract_OverColumnStore_RestoresOldThenNew()
    {
        var store = ColumnStore.Create(["A", "B"], initialRowCapacity: 1);
        store.AppendRow();
        store.SetCellQuiet(0, 0, "orig");
        store.SetCellQuiet(0, 1, "keep");

        var undo = new Stack<(int Row, int Col, string? Old, string New)>();
        var redo = new Stack<(int Row, int Col, string? Old, string New)>();

        // 编辑 A0: orig → edited（压栈）
        var oldValue = store.GetCell(0, 0);
        store.SetCell(0, 0, "edited");
        undo.Push((0, 0, oldValue, "edited"));
        redo.Clear();
        Assert.Equal("edited", store.GetCell(0, 0));

        // 撤销：写回 OldValue，压 redo（等价 OnUndoClick）
        var u = undo.Pop();
        var curBeforeUndo = store.GetCell(u.Row, u.Col);
        store.SetCell(u.Row, u.Col, u.Old);
        redo.Push((u.Row, u.Col, curBeforeUndo, curBeforeUndo ?? string.Empty));
        Assert.Equal("orig", store.GetCell(0, 0));
        Assert.Single(redo);
        Assert.Empty(undo);

        // 重做：写 NewValue，压 undo（等价 OnRedoClick）
        var r = redo.Pop();
        var curBeforeRedo = store.GetCell(r.Row, r.Col);
        store.SetCell(r.Row, r.Col, r.New);
        undo.Push((r.Row, r.Col, curBeforeRedo, curBeforeRedo ?? string.Empty));
        Assert.Equal("edited", store.GetCell(0, 0));
        Assert.Single(undo);
        Assert.Empty(redo);

        Assert.Equal("keep", store.GetCell(0, 1));
    }

    /// <summary>
    /// 【P4 WF1 行为变更】P2/P3.2 时 <see cref="ColumnStore.InsertRow"/>/<see cref="ColumnStore.DeleteRow"/>
    /// 直接 <c>_dirty.Clear()</c>（此测试原名 <c>StructuralOps_ClearDirtyTracking_KnownDesign</c>，断言"清空脏跟踪"）。
    /// P4 为支持"只写脏数据"的增量写回，把该行为改成 <b>remap 脏行号</b>——脏标记随行移动保留，
    /// 不再丢失（避免"编辑几格→再增删行→那几格漏写"）。故此测试更新为断言新的 remap 行为，
    /// 并新增 <see cref="ColumnStore.StructureChanged"/>（保存路径据此在结构变更后 fallback 全量写）。
    /// 原"清脏"断言已作废，因为它锁定的正是本次要修的 bug。
    /// </summary>
    [Fact]
    public void StructuralOps_RemapDirtyTracking_AndSetStructureChanged()
    {
        var store = ColumnStore.Create(["A"], initialRowCapacity: 4);
        store.AppendRow(); // row0
        store.AppendRow(); // row1
        store.SetCell(1, 0, "x"); // 脏格 (1,0)
        Assert.True(store.IsDirty(1, 0));
        Assert.False(store.StructureChanged);

        // 在 row0 插入 → (1,0) 下移到 (2,0)，脏标记保留（不再被清空）
        store.InsertRow(0);
        Assert.True(store.StructureChanged);
        Assert.False(store.IsDirty(1, 0));
        Assert.True(store.IsDirty(2, 0)); // remap 生效，脏标记跟着行走

        // 删除 row0 → (2,0) 上移回 (1,0)，脏标记仍保留
        store.DeleteRow(0);
        Assert.True(store.IsDirty(1, 0));

        // ClearDirty（保存成功后调）清脏 + 重置 StructureChanged
        store.ClearDirty();
        Assert.Empty(store.DirtyCells);
        Assert.False(store.StructureChanged);
    }

    // ─────────────────────────────────────────────────────────────────
    //  反射辅助
    // ─────────────────────────────────────────────────────────────────

    private static List<(string Name, ColumnStore Store, int TotalRows)> InvokeBuildStoresFromExcel(
        string path
    )
    {
        var raw = MainWindowType
            .GetMethod("BuildStoresFromExcel", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, [path]);
        var list = new List<(string, ColumnStore, int)>();
        foreach (var tuple in (IEnumerable)raw!)
        {
            var type = tuple.GetType();
            list.Add(
                (
                    (string)type.GetField("Item1")!.GetValue(tuple)!,
                    (ColumnStore)type.GetField("Item2")!.GetValue(tuple)!,
                    (int)type.GetField("Item4")!.GetValue(tuple)!
                )
            );
        }

        return list;
    }

    private static List<string> ReadSheetNamesViaReflection(string path)
    {
        var readerType = LoadXlsxEditorType("NumDesTools.XlsxEditor.OoxmlLazyReader");
        var raw = readerType
            .GetMethod("ReadSheetNames", BindingFlags.Public | BindingFlags.Static)!
            .Invoke(null, [path]);
        return (List<string>)raw!;
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
}
