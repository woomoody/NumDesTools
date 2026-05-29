using System.Collections.Generic;
using System.Text.RegularExpressions;
using NumDesTools;
using NumDesTools.Config;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// 测试 PubMetToExcelFunc 中三个自检方法的行为。
/// </summary>
public class DataCheckTests
{
    // ── Helper ──────────────────────────────────────────────────────────────

    /// <summary>
    /// 构造 MiniExcel Query 风格的行列表。
    /// 结构：row[0]=注释行, row[1]=字段名行, row[2]=类型行, row[3]=中文名行, row[4+]=数据行。
    /// cols 字典的 key = 列字母（A/B/C…），与 MiniExcel 默认 dynamic 行为一致。
    /// </summary>
    private static List<dynamic> MakeRows(
        Dictionary<string, object> fieldNames,   // row[1]: 字段名
        Dictionary<string, object> fieldTypes,   // row[2]: 类型
        Dictionary<string, object> chineseNames, // row[3]: 中文名（行[3]["A"] 须为 "#"）
        params Dictionary<string, object>[] dataRows
    )
    {
        var rows = new List<dynamic>
        {
            new Dictionary<string, object> { ["A"] = "#comment" }, // row[0] 注释
            fieldNames,
            fieldTypes,
            chineseNames,
        };
        foreach (var dr in dataRows)
            rows.Add(dr);
        return rows;
    }

    /// <summary>快速构造单 sheet 最小合法表头（A="#", B 字段名非空）。</summary>
    private static (
        Dictionary<string, object> fields,
        Dictionary<string, object> types,
        Dictionary<string, object> labels
    ) MinHeader(string bFieldName = "id", string bType = "int")
    {
        var fields = new Dictionary<string, object> { ["A"] = "#", ["B"] = bFieldName, ["C"] = "val" };
        var types  = new Dictionary<string, object> { ["A"] = "#", ["B"] = "int", ["C"] = bType };
        var labels = new Dictionary<string, object> { ["A"] = "#", ["B"] = "ID", ["C"] = "值" };
        return (fields, types, labels);
    }

    private static Dictionary<string, object> DataRow(string a, string b, string c = "") =>
        new() { ["A"] = a, ["B"] = b, ["C"] = c };

    // 预编译默认括号规则，供 CheckValueFormat 测试复用
    private static readonly List<(string, string, Regex, Regex)> DefaultCoupleRegexes =
        PubMetToExcelFunc.BuildCoupleRegexes(new[]
        {
            new GlobalVariable.CoupleKey("[", "]"),
            new GlobalVariable.CoupleKey("{", "}"),
            new GlobalVariable.CoupleKey("\"", "\""),
        });

    private static readonly List<string> NoNormalChars  = new();
    private static readonly List<string> NoSpecialChars = new();

    // ── CheckRepeatValue ────────────────────────────────────────────────────

    [Fact]
    public void CheckRepeatValue_DuplicateB_ReturnsBothRows()
    {
        var (fields, types, labels) = MinHeader();
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "100"),
            DataRow("#", "200"),
            DataRow("#", "100")  // 重复
        );

        var result = PubMetToExcelFunc.CheckRepeatValue(rows, "Sheet1");

        Assert.Equal(2, result.Count);
        Assert.All(result, r => Assert.Equal("100", r.Item1));
        Assert.All(result, r => Assert.Equal("数据重复", r.Item5));
    }

    [Fact]
    public void CheckRepeatValue_NoDuplicate_ReturnsEmpty()
    {
        var (fields, types, labels) = MinHeader();
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "100"),
            DataRow("#", "200"),
            DataRow("#", "300")
        );

        var result = PubMetToExcelFunc.CheckRepeatValue(rows, "Sheet1");

        Assert.Empty(result);
    }

    [Fact]
    public void CheckRepeatValue_FirstColumnNotHash_Skips()
    {
        var fields = new Dictionary<string, object> { ["A"] = "#", ["B"] = "id", ["C"] = "val" };
        var types  = new Dictionary<string, object> { ["A"] = "#", ["B"] = "int", ["C"] = "int" };
        // 中文名行（row[3]）A 列非 "#" → CheckRepeatValue 判定 dataRows[0]["A"] != "#" 跳过
        var labels = new Dictionary<string, object> { ["A"] = "x", ["B"] = "ID", ["C"] = "值" };
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "100"),
            DataRow("#", "100")  // 有重复，但因跳过不应报告
        );

        var result = PubMetToExcelFunc.CheckRepeatValue(rows, "Sheet1");

        Assert.Empty(result);
    }

    [Fact]
    public void CheckRepeatValue_FewerThanFourRows_ReturnsEmpty()
    {
        var rows = new List<dynamic>
        {
            new Dictionary<string, object> { ["A"] = "#", ["B"] = "id" },
            new Dictionary<string, object> { ["A"] = "#", ["B"] = "int" },
            new Dictionary<string, object> { ["A"] = "#", ["B"] = "ID" },
        };

        var result = PubMetToExcelFunc.CheckRepeatValue(rows, "Sheet1");

        Assert.Empty(result);
    }

    // ── CheckValueFormat ────────────────────────────────────────────────────

    [Fact]
    public void CheckValueFormat_NormalCharMatch_NonString_ReportsDoubleComma()
    {
        var (fields, types, labels) = MinHeader(bType: "int");
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "1,,2")  // val 列含双逗号
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   new[] { ",," },
            specialChars:  NoSpecialChars,
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Single(result);
        Assert.Equal("多逗号或中文逗号", result[0].Item5);
    }

    [Fact]
    public void CheckValueFormat_SpecialCharMatch_NonString_ReportsLessComma()
    {
        var (fields, types, labels) = MinHeader(bType: "int");
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "][")   // val 列含 ][
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   NoNormalChars,
            specialChars:  new[] { "][" },
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Single(result);
        Assert.Equal("少逗号", result[0].Item5);
    }

    [Fact]
    public void CheckValueFormat_MismatchedBrackets_ReportsBracketError()
    {
        var (fields, types, labels) = MinHeader(bType: "int[]");
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "[1,2")  // 缺 ]
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   NoNormalChars,
            specialChars:  NoSpecialChars,
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Single(result);
        Assert.Equal("括号问题", result[0].Item5);
    }

    [Fact]
    public void CheckValueFormat_OddDoubleQuotes_ReportsQuoteError()
    {
        var (fields, types, labels) = MinHeader(bType: "int");
        var rows = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "\"hello")  // 一个双引号
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   NoNormalChars,
            specialChars:  NoSpecialChars,
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Single(result);
        Assert.Equal("双引号问题", result[0].Item5);
    }

    [Fact]
    public void CheckValueFormat_StringType_SkipsCharChecks()
    {
        var fields = new Dictionary<string, object> { ["A"] = "#", ["B"] = "id", ["C"] = "name" };
        var types  = new Dictionary<string, object> { ["A"] = "#", ["B"] = "int", ["C"] = "string" };
        var labels = new Dictionary<string, object> { ["A"] = "#", ["B"] = "ID", ["C"] = "名字" };
        var rows   = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "hello,,world")  // 双逗号，但 string 类型应跳过
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   new[] { ",," },
            specialChars:  NoSpecialChars,
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Empty(result);
    }

    [Fact]
    public void CheckValueFormat_FirstColumnNotHash_Skips()
    {
        var fields = new Dictionary<string, object> { ["A"] = "#", ["B"] = "id", ["C"] = "val" };
        var types  = new Dictionary<string, object> { ["A"] = "#", ["B"] = "int", ["C"] = "int" };
        // 中文名行 A 列非 "#"
        var labels = new Dictionary<string, object> { ["A"] = "x", ["B"] = "ID", ["C"] = "值" };
        var rows   = MakeRows(
            fields, types, labels,
            DataRow("#", "1", "1,,2")
        );

        var result = PubMetToExcelFunc.CheckValueFormat(
            rows, "Sheet1",
            normalChars:   new[] { ",," },
            specialChars:  NoSpecialChars,
            coupleRegexes: DefaultCoupleRegexes
        );

        Assert.Empty(result);
    }

    // ── BuildCoupleRegexes ──────────────────────────────────────────────────

    [Fact]
    public void BuildCoupleRegexes_MatchesLeftAndRight_Correctly()
    {
        var regexes = PubMetToExcelFunc.BuildCoupleRegexes(new[]
        {
            new GlobalVariable.CoupleKey("[", "]"),
        });

        Assert.Single(regexes);
        var (left, right, leftRx, rightRx) = regexes[0];
        Assert.Equal("[", left);
        Assert.Equal("]", right);
        Assert.Single(leftRx.Matches("[1,2]"));
        Assert.Single(rightRx.Matches("[1,2]"));
        Assert.Equal(3, leftRx.Matches("[[1],[2]]").Count);   // [ 出现 3 次
        Assert.Equal(3, rightRx.Matches("[[1],[2]]").Count);  // ] 出现 3 次
    }

    [Fact]
    public void BuildCoupleRegexes_RegexEscapesSpecialChars()
    {
        var regexes = PubMetToExcelFunc.BuildCoupleRegexes(new[]
        {
            new GlobalVariable.CoupleKey("(", ")"),
        });

        var (_, _, leftRx, rightRx) = regexes[0];
        // "(" 是正则特殊字符，必须被转义才能正确匹配字面量
        Assert.Single(leftRx.Matches("(a,b)"));
        Assert.Single(rightRx.Matches("(a,b)"));
    }
}
