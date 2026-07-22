using System.Collections.Generic;
using NumDesTools.Battle;
using Xunit;

namespace NumDesTools.Tests;

public class LteDataUnitTests
{
    // baseDic 辅助：按 LTE【基础】列顺序构造一行数据
    // [0]=数据编号, [1]=资源编号, [2]=图片编号, [3]=首次出现,
    // [4]=唯一代号, [5]=代号, [6]=当前包装, [7]=级别, [8]=类型
    private static List<string> MakeBaseRow(
        string id,
        string prefabId = "",
        string iconId = "",
        string firstMap = "",
        string onlyName = "",
        string name = "",
        string package = "",
        string level = "",
        string type = ""
    ) => new() { id, prefabId, iconId, firstMap, onlyName, name, package, level, type };

    // -------- FixFieldData --------

    [Fact]
    public void FixFieldData_FindByName_地标路径_返回正确fixData()
    {
        var baseDic = new Dictionary<string, List<string>>
        {
            ["100010001"] = MakeBaseRow("100010001", name: "A1", package: "pkgA", type: "地标"),
        };

        var (fixData, findData, id) = LteData.FixFieldData("A1", "2", "地标", baseDic);

        Assert.Equal("100010001", id);
        Assert.Equal("100010001", findData);
        Assert.Contains("14,100010001", fixData);
    }

    [Fact]
    public void FixFieldData_NameEmpty_唯一代号回退_返回正确id()
    {
        // name列（[5]）为空，唯一代号列（[4]）有值 "1L-A3"
        var baseDic = new Dictionary<string, List<string>>
        {
            ["100010003"] = MakeBaseRow(
                "100010003",
                onlyName: "1L-A3",
                name: "",
                package: "pkgA",
                type: "链"
            ),
        };

        var (fixData, _, id) = LteData.FixFieldData("1L-A3", "", "链", baseDic);

        Assert.Equal("100010003", id);
    }

    [Fact]
    public void FixFieldData_链类型_返回正确fixData()
    {
        var baseDic = new Dictionary<string, List<string>>
        {
            ["100010001"] = MakeBaseRow("100010001", name: "链A", package: "pkgB", type: "链"),
        };

        var (fixData, _, id) = LteData.FixFieldData("链A", "", "链", baseDic);

        Assert.Equal("100010001", id);
        Assert.Contains("7,100010001", fixData);
    }

    // -------- SetDic / GetDic 联合测试 --------

    [Fact]
    public void SetDic_ThenGetDic_返回存入的链列表()
    {
        // dep="100010001" → Set(dep, 2, "00") = "100010000"（key in strDic）
        // idList 含 100010001..100010003，其中 +1=100010001, +2=100010002, +3=100010003
        var exportWildcardDyData = new Dictionary<string, string>
        {
            ["dep"] = "100010001",
            ["链长"] = "3",
        };
        var strDic = new Dictionary<string, Dictionary<string, List<string>>>();
        var idList = new List<string> { "100010001", "100010002", "100010003" };

        LteCore.SetDic(exportWildcardDyData, strDic, "myDic", "dep", "2", "00", "链长", idList);

        Assert.True(strDic.ContainsKey("myDic"));

        // GetDic：itemKey = 某个在 linkList 里的 id，val 必须在 dependsValueList 中
        // linkList 包含 100010001（i=0, fixWildcardValue+1=100010001，在 idList 里）
        var dy2 = new Dictionary<string, string> { ["itemKey"] = "100010001" };
        var res = LteCore.GetDic(strDic, dy2, "myDic", "itemKey", "2", "00");
        Assert.False(string.IsNullOrEmpty(res));
    }
}
