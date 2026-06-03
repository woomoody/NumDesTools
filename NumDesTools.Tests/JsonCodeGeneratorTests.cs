using NumDesTools.ExcelToLua;

namespace NumDesTools.Tests;

/// <summary>
/// H20/EXTRA2: Cell2JsonValue / Cell2JsonValue2 对 REWARD / LUA_TABLE / ANY 类型
/// 不应返回空字符串，否则导出配置静默丢失数据。
/// </summary>
public class JsonCodeGeneratorTests
{
    // ── 辅助：构造只有 key + 单个字段的 SheetData ──────────────────────────

    private static SheetData MakeSheet(int fieldType, string cellValue)
    {
        var sd = new SheetData(0, 0);
        // 字段 0 = key（INT 类型），字段 1 = 待测字段
        sd.AddField(0, FieldTypeDefine.INT, "id", "", false, "");
        sd.AddField(1, fieldType, "val", "", false, "");

        var row = new RowData();
        row.AddCell(0, 0, "1"); // key
        row.AddCell(0, 1, cellValue); // 待测值
        sd.rows.Add(row);
        return sd;
    }

    // ── ToJsonCode（Cell2JsonValue）─────────────────────────────────────────

    [Fact]
    public void ToJsonCode_RewardField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.REWARD, "[[1,100],[2,50]]");
        var json = JsonCodeGenerator.ToJsonCode(sd);
        // val 字段不应产生空字符串值
        Assert.DoesNotContain("\"val\":\"\"", json);
        // 原始数值应透传
        Assert.Contains("100", json);
    }

    [Fact]
    public void ToJsonCode_LuaTableField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.LUA_TABLE, "{key=1,val=2}");
        var json = JsonCodeGenerator.ToJsonCode(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("{key=1,val=2}", json);
    }

    [Fact]
    public void ToJsonCode_AnyField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.ANY, "someValue");
        var json = JsonCodeGenerator.ToJsonCode(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("someValue", json);
    }

    [Fact]
    public void ToJsonCode_RewardArrayField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.REWARD_ARRAY, "[[[1,100]],[[2,50]]]");
        var json = JsonCodeGenerator.ToJsonCode(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("100", json);
    }

    // ── RechargeToJson（Cell2JsonValue2）────────────────────────────────────

    [Fact]
    public void RechargeToJson_RewardField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.REWARD, "[[1,100]]");
        var json = JsonCodeGenerator.RechargeToJson(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("100", json);
    }

    [Fact]
    public void RechargeToJson_LuaTableField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.LUA_TABLE, "{a=1}");
        var json = JsonCodeGenerator.RechargeToJson(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("{a=1}", json);
    }

    [Fact]
    public void RechargeToJson_AnyField_EmitsValue()
    {
        var sd = MakeSheet(FieldTypeDefine.ANY, "42");
        var json = JsonCodeGenerator.RechargeToJson(sd);
        Assert.DoesNotContain("\"val\":\"\"", json);
        Assert.Contains("42", json);
    }
}
