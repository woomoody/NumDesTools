using NumDesTools.XlsxEditor;
using OfficeOpenXml;

namespace NumDesTools.Tests;

/// <summary>
/// ColumnStoreExcelLoader 读取集成测试：用真实大表 Item.xlsx（~6.5万行 × 85 列）验证
/// "读 xlsx → 产出正确 ColumnStore" 这条路径。抽查单元格值与 EPPlus 交叉验证，不凭感觉断言。
/// EPPlus 的 <c>cell.Value?.ToString()</c> 正是既有 DataTable 加载路径采用的文本化方式（见 status.md），
/// 用它做期望值可直接反映"ColumnStore 是否忠实还原旧路径的可见文本"。
/// 沿用 <see cref="OoxmlLazyReaderTests"/> 的约定：直接依赖真实文件路径存在。
/// </summary>
public sealed class ColumnStoreExcelLoaderTests
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    static ColumnStoreExcelLoaderTests() =>
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

    [Fact]
    public void Load_Item_MatchesExpectedShape()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);

        // 行/列数动态取自 EPPlus dimension（真实文件会被上游刷表增行，硬编码基线会脆断）。
        // 断言强度不变：ColumnStore 必须与 EPPlus 报告的形状逐字吻合。
        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[0];
        Assert.Equal(sheet.Dimension.End.Row, store.RowCount);
        Assert.Equal(sheet.Dimension.End.Column, store.ColumnCount);
        Assert.Equal("A", store.ColumnNames[0]);
    }

    [Fact]
    public void Load_Item_SampledCells_MatchEpPlus()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);

        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[0];

        // (storeRow0Based, storeCol0Based) → EPPlus 是 1-based，行列各 +1
        (int Row, int Col)[] samples =
        [
            (4, 1), // Excel(5,2) = 11010001（数字，首个数据行的 id）
            (4, 2), // Excel(5,3) = 引导蘑菇1级（中文字符串）
            (99, 1), // Excel(100,2) = 13010504
            (999, 1), // Excel(1000,2)
            (29999, 1), // Excel(30000,2) = 7616068834（10 位大整数）
            (29999, 2), // Excel(30000,3) 中文，避免硬编码乱码，动态取 EPPlus 期望值
            (65104, 1), // 末行 Excel(65105,2) = 17130112
            (1, 1), // Excel(2,2) = id（列名行，字符串 "id"）
            (0, 0), // Excel(1,1) = #（标题行第一列）
            (65104, 82), // Excel(65105,83) 末尾列区（验证 jagged 扩列后可访问）
        ];

        foreach (var (row, col) in samples)
        {
            var expected = sheet.Cells[row + 1, col + 1].Value?.ToString();
            var actual = store.GetCell(row, col);
            Assert.Equal(NormalizeEmptyToNull(expected), actual);
        }
    }

    [Fact]
    public void Load_Item_KnownLiteralValues()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);

        // 硬编码若干与编码无关（纯 ASCII/数字）的已知值，作为独立于 EPPlus 的第二重锚点。
        // 低行号锚点稳定（上游刷表在文件尾部增行，前面数据行不动）；末行值随行数变化，
        // 故末行 id 值改为与 EPPlus 交叉验证（不硬编码绝对行号，文件会增行）。
        Assert.Equal("11010001", store.GetCell(4, 1));
        Assert.Equal("13010504", store.GetCell(99, 1));
        Assert.Equal("7616068834", store.GetCell(29999, 1));
        Assert.Equal("id", store.GetCell(1, 1));
        Assert.Equal("#", store.GetCell(0, 0));

        using var package = new ExcelPackage(new FileInfo(ItemPath));
        var sheet = package.Workbook.Worksheets[0];
        var lastRow = store.RowCount - 1;
        Assert.Equal(sheet.Cells[lastRow + 1, 2].Value?.ToString(), store.GetCell(lastRow, 1));
    }

    [Fact]
    public void Load_Item_InternPool_ReusesReferences()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath);

        // 列 0（Excel A 列）前 4 行都是 "#"，其余多为空——收集该列所有非空引用，
        // 断言等值内容全部是同一个引用（驻留在加载路径 SetCellQuiet 里生效）。
        string? firstHash = null;
        var distinctRefsForHash = new List<object>();
        for (var row = 0; row < store.RowCount; row++)
        {
            var value = store.GetCell(row, 0);
            if (value != "#")
            {
                continue;
            }

            firstHash ??= value;
            if (!distinctRefsForHash.Any(existing => ReferenceEquals(existing, value)))
            {
                distinctRefsForHash.Add(value);
            }
        }

        Assert.NotNull(firstHash);
        // 所有 "#" 只对应一个引用，证明没有为每个重复值重新分配字符串
        Assert.Single(distinctRefsForHash);
    }

    private static string? NormalizeEmptyToNull(string? value) =>
        string.IsNullOrEmpty(value) ? null : value;
}
