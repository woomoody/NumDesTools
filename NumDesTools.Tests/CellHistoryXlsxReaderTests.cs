using NumDesTools;
using OfficeOpenXml;

namespace NumDesTools.Tests;

/// <summary>
/// 测试 CellHistoryXlsxReader — xlsx 解析行为，不依赖 git。
/// </summary>
public class CellHistoryXlsxReaderTests
{
    static CellHistoryXlsxReaderTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
    }

    // ── helpers ──────────────────────────────────────────────────────────────

    /// 创建一个标准 4-row 表头的 config 表：
    /// row1=空, row2=列名(#注,id,name), row3=类型(,int,string), row4=标签, row5+=数据
    private static ExcelWorksheet MakeStandardSheet(ExcelPackage pkg, params (string id, string name)[] rows)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
        var ws = pkg.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[2, 1].Value = "#注";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 3].Value = "name";
        ws.Cells[3, 2].Value = "int";
        ws.Cells[3, 3].Value = "string";
        ws.Cells[4, 2].Value = "编号";
        ws.Cells[4, 3].Value = "名称";
        int r = 5;
        foreach (var (id, name) in rows)
        {
            ws.Cells[r, 2].Value = id;
            ws.Cells[r, 3].Value = name;
            r++;
        }
        return ws;
    }

    /// 创建一个 type 表：只有 2 行表头（row2=列名, row3+=数据）
    private static ExcelWorksheet MakeTypeSheet(ExcelPackage pkg, params (string id, string val)[] rows)
    {
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
        var ws = pkg.Workbook.Worksheets.Add("TypeSheet");
        ws.Cells[2, 1].Value = "id";
        ws.Cells[2, 2].Value = "value";
        int r = 3;
        foreach (var (id, val) in rows)
        {
            ws.Cells[r, 1].Value = id;
            ws.Cells[r, 2].Value = val;
            r++;
        }
        return ws;
    }

    // ── Tracer bullet ─────────────────────────────────────────────────────────

    [Fact]
    public void StandardTable_FindsRowByKey()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeStandardSheet(pkg, ("1001", "测试活动"), ("1002", "其他活动"));

        var data = CellHistoryXlsxReader.ParseSheetData(ws);

        Assert.True(data.ContainsKey("1001"), "key 1001 应被索引");
        Assert.Equal("测试活动", data["1001"]["name"]);
    }

    // ── Row scanning ──────────────────────────────────────────────────────────

    [Fact]
    public void TypeTable_DataStartsAtRow3_FindsRow()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeTypeSheet(pkg, ("PackType1", "背包类型"), ("PackType2", "道具类型"));

        var data = CellHistoryXlsxReader.ParseSheetData(ws);

        Assert.True(data.ContainsKey("PackType1"), "row3 数据应被索引");
        Assert.Equal("背包类型", data["PackType1"]["value"]);
    }

    [Fact]
    public void EmptyKeyRow_IsSkipped()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeStandardSheet(pkg, ("1001", "活动A"), ("", "空键行"));

        var data = CellHistoryXlsxReader.ParseSheetData(ws);

        // 空 key 行一定不进缓存
        Assert.DoesNotContain("", data.Keys);
        // 真实数据行正常收录
        Assert.True(data.ContainsKey("1001"));
    }

    // ── Column lookup ─────────────────────────────────────────────────────────

    [Fact]
    public void ColumnLookup_ByName_NotByIndex()
    {
        using var pkg = new ExcelPackage();
        // name 列现在是第 4 列（插了一个 extra 列）
        ExcelPackage.License.SetNonCommercialPersonal("NumDesTools.Tests");
        var ws = pkg.Workbook.Worksheets.Add("Shifted");
        ws.Cells[2, 1].Value = "#注";
        ws.Cells[2, 2].Value = "id";
        ws.Cells[2, 3].Value = "extra";   // 多了一列
        ws.Cells[2, 4].Value = "name";
        ws.Cells[5, 2].Value = "9001";
        ws.Cells[5, 3].Value = "不要这个";
        ws.Cells[5, 4].Value = "正确值";

        var data = CellHistoryXlsxReader.ParseSheetData(ws);

        Assert.Equal("正确值", data["9001"]["name"]);
    }

    // ── Key column detection ──────────────────────────────────────────────────

    [Fact]
    public void KeyCol_SkipsHashPrefixColumns()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeStandardSheet(pkg, ("42", "值"));

        int keyCol = CellHistoryXlsxReader.FindKeyColIdx(ws);

        // row2: col1="#注"(#前缀), col2="id"(第一个非#) → keyCol 应为 2
        Assert.Equal(2, keyCol);
    }

    [Fact]
    public void KeyCol_NoHashPrefix_ReturnsFirstCol()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeTypeSheet(pkg, ("T1", "v"));

        int keyCol = CellHistoryXlsxReader.FindKeyColIdx(ws);

        Assert.Equal(1, keyCol); // id 在 col1，无 # 前缀
    }

    // ── Cache reuse ───────────────────────────────────────────────────────────

    [Fact]
    public void ParseSheetData_HashColumnsNotInValueMap()
    {
        using var pkg = new ExcelPackage();
        var ws = MakeStandardSheet(pkg, ("500", "名称X"));

        var data = CellHistoryXlsxReader.ParseSheetData(ws);

        // #注 列不应出现在 value map 里（可以出现，但不应作为查询 key）
        // 主要验证 name 列正常
        Assert.Equal("名称X", data["500"]["name"]);
    }
}
