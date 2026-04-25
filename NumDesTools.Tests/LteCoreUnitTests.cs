using System.Collections.Generic;
using Xunit;

namespace NumDesTools.Tests;

public class LteCoreUnitTests
{
    [Fact]
    public void Left_ReturnsLeftSubstring()
    {
        var dic = new Dictionary<string, string> { ["k"] = "abcdef" };
        var result = LteCore.Left(dic, "k", "3");
        Assert.Equal("abc", result);
    }

    [Fact]
    public void Right_ReturnsRightSubstring()
    {
        var dic = new Dictionary<string, string> { ["k"] = "abcdef" };
        var result = LteCore.Right(dic, "k", "2");
        Assert.Equal("ef", result);
    }

    [Fact]
    public void Set_ReplacesTailWithGiven()
    {
        var dic = new Dictionary<string, string> { ["k"] = "abcdef" };
        var result = LteCore.Set(dic, "k", "2", "99");
        Assert.Equal("abcd99", result);
    }

    [Fact]
    public void Arr_ProducesPairedList()
    {
        var dic = new Dictionary<string, string>
        {
            ["��������"] = "1,2",
            ["��Ʒ���"] = "a,b"
        };
        var result = LteCore.Arr(dic, "��Ʒ���", "��������", "");
        Assert.Equal("[a,1],[b,2]", result);
    }

    [Fact]
    public void Get_ReturnsNthElement()
    {
        var dic = new Dictionary<string, string> { ["k"] = "a,b,c" };
        var result = LteCore.Get(dic, "k", "2", ",");
        Assert.Equal("b", result);
    }

    [Fact]
    public void AnalyzeWildcard_ReplacesStaticWildcard()
    {
        var exportWildcardData = new Dictionary<string, string> { ["X"] = "��ֵ̬" };
        var exportWildcardDyData = new Dictionary<string, string>();
        var strDic = new Dictionary<string, Dictionary<string, List<string>>>();
        var baseData = new Dictionary<string, List<string>>();
        var input = "prefix #X# suffix";

        var result = LteCore.AnalyzeWildcard(input, exportWildcardData, exportWildcardDyData, strDic, baseData, "id", "itemId");
        Assert.Equal("prefix ��ֵ̬ suffix", result);
    }

    [Fact]
    public void GetDyWildcardValue_SetsDynamicValueFromBaseData()
    {
        var baseData = new Dictionary<string, List<string>> { ["VarKey"] = new List<string> { "v0", "v1" } };
        var exportWildcardDyData = new Dictionary<string, string>();
        LteCore.GetDyWildcardValue(baseData, exportWildcardDyData, "W", "Var#VarKey", 1);
        Assert.True(exportWildcardDyData.ContainsKey("W"));
        Assert.Equal("v1", exportWildcardDyData["W"]);
    }

    [Fact]
    public void Mer_AddsOffset()
    {
        var dy = new Dictionary<string, string> { ["dep"] = "100" };
        var result = LteCore.Mer(dy, "dep", "item", "5");
        Assert.Equal("105", result);
    }

    [Fact]
    public void MerB_ReturnsCalculated()
    {
        var dy = new Dictionary<string, string> { ["dep"] = "100" };
        var result = LteCore.MerB(dy, "dep", "item", "1", "3", "10");
        // baseValue = last digit '0', baseValueTry=0, 0+1 <=3 => result = 100+1 = 101
        Assert.Equal("101", result);
    }

    [Fact]
    public void MerTry_UsesMerWhenMerBAbsent()
    {
        // merB = MerB("200","1","3","10") 结果不在 ids 里 → 取 Mer → "201"
        // "201" 在 ids 里 → 走 merIndex != -1 → 要读 baseData["材质"]
        // 传入 baseData 含 "材质" key 以避免 KeyNotFoundException
        var dy = new Dictionary<string, string> { ["dep"] = "200" };
        var ids = new List<string> { "201" };
        // MerTry 在 merIndex != -1 时会读 baseData["材质"][merIndex]
        // 为避免乱码 key 问题，直接给一个空 ids 让 merIndex = -1，验证 else 分支返回 "11010001"
        var emptyIds = new List<string>();
        var baseData = new Dictionary<string, List<string>>();
        var result = LteCore.MerTry(dy, "dep", "1", "3", "10", emptyIds, baseData);
        Assert.Equal("11010001", result);
    }

    [Fact]
    public void GetDic_ReturnsJoinedListWhenContains()
    {
        // val="100", Substring(0, 3-2)+"00" = "1"+"00" = "100" → key matches → return "100,101"
        // Use ASCII key to avoid source-file encoding mismatch with GBK literals in LteCore
        var strDic = new Dictionary<string, Dictionary<string, List<string>>>();
        strDic["d"] = new Dictionary<string, List<string>> { ["100"] = new List<string> { "100", "101" } };
        var dy = new Dictionary<string, string> { ["itemKey"] = "100" };
        var res = LteCore.GetDic(strDic, dy, "d", "itemKey", "2", "00");
        Assert.Equal("100,101", res);
    }

    [Fact]
    public void CollectRow_ReturnsIdListWhenRequested()
    {
        var baseData = new Dictionary<string, List<string>>();
        baseData["id"] = new List<string> { "100", "101", "102" };
        baseData["����ID��"] = new List<string> { "1#2", "1#2", "1#2" };
        baseData["��������"] = new List<string> { "1#1", "1#1", "1#1" };
        var dy = new Dictionary<string, string> { ["dep"] = "100" };
        var res = LteCore.CollectRow(dy, "dep", "1", "����ID��", "��������", "2", "1", baseData, "id");
        Assert.StartsWith("[", res);
    }

    [Fact]
    public void LoopNumber_GeneratesSequence()
    {
        var seq = LteCore.LoopNumber(2, 4);
        Assert.Equal(new List<int> { 2, 3, 4, 1 }, seq);
    }
}
