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
            ["消耗量组"] = "1,2",
            ["物品编号"] = "a,b"
        };
        var result = LteCore.Arr(dic, "物品编号", "消耗量组", "");
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
        var exportWildcardData = new Dictionary<string, string> { ["X"] = "静态值" };
        var exportWildcardDyData = new Dictionary<string, string>();
        var strDic = new Dictionary<string, Dictionary<string, List<string>>>();
        var baseData = new Dictionary<string, List<string>>();
        var input = "prefix #X# suffix";

        var result = LteCore.AnalyzeWildcard(input, exportWildcardData, exportWildcardDyData, strDic, baseData, "id", "itemId");
        Assert.Equal("prefix 静态值 suffix", result);
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
        var dy = new Dictionary<string, string> { ["dep"] = "200" };
        var ids = new List<string> { "201" };
        var result = LteCore.MerTry(dy, "dep", "1", "3", "10", ids);
        Assert.Equal(LteCore.Mer(dy, "dep", string.Empty, "1"), result);
    }

    [Fact]
    public void GetDic_ReturnsJoinedListWhenContains()
    {
        var strDic = new Dictionary<string, Dictionary<string, List<string>>>();
        strDic["d"] = new Dictionary<string, List<string>> { ["10A"] = new List<string> { "100", "101" } };
        var dy = new Dictionary<string, string> { ["物品编号"] = "100" };
        var res = LteCore.GetDic(strDic, dy, "d", "物品编号", "2", "00");
        Assert.Equal("100,101", res);
    }

    [Fact]
    public void CollectRow_ReturnsIdListWhenRequested()
    {
        var baseData = new Dictionary<string, List<string>>();
        baseData["id"] = new List<string> { "100", "101", "102" };
        baseData["消耗ID组"] = new List<string> { "1#2", "1#2", "1#2" };
        baseData["消耗量组"] = new List<string> { "1#1", "1#1", "1#1" };
        var dy = new Dictionary<string, string> { ["dep"] = "100" };
        var res = LteCore.CollectRow(dy, "dep", "1", "消耗ID组", "消耗量组", "2", "1", baseData, "id");
        Assert.StartsWith("[", res);
    }

    [Fact]
    public void LoopNumber_GeneratesSequence()
    {
        var seq = LteCore.LoopNumber(2, 4);
        Assert.Equal(new List<int> { 2, 3, 4, 1 }, seq);
    }
}
