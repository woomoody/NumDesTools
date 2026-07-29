using NumDesTools.UI;

namespace NumDesTools.Tests;

/// <summary>
/// MahAppsHelper 资源合并去重逻辑的单测。
///
/// 背景（见 NumDesOutput/analysis/wpfui-migration-optimization-report-2026-07-29.md P1）：
/// EnsureInitialized() 里"资源是否已合并"的判断全靠 ResourceDictionary.Source URI 字符串匹配，
/// 之前完全没有测试覆盖，只能靠"改代码→重启 Excel→人工看"验证，导致多轮回归
/// （黑窗口→CTP 污染→窗口打不开→CTP 还原）反复出现且互相掩盖。
///
/// 这里把判断逻辑抽成纯函数 <see cref="MahAppsHelper.IsResourceMerged"/>，
/// 接受 Uri? 序列（而非 ResourceDictionary），不需要真实 STA Application
/// 也不会触发 WPF 的资源加载 IO，就能验证"合并顺序/去重"这类纯逻辑的正确性。
/// </summary>
public class MahAppsHelperResourceMergeTests
{
    [Fact]
    public void IsResourceMerged_EmptyCollection_ReturnsFalse()
    {
        var sources = Array.Empty<Uri?>();

        var result = MahAppsHelper.IsResourceMerged(sources, "Wpf.Ui");

        Assert.False(result);
    }

    [Fact]
    public void IsResourceMerged_MatchingUriFragmentPresent_ReturnsTrue()
    {
        Uri?[] sources = [new Uri("http://example/Wpf.Ui/Resources/Theme/Dark.xaml")];

        var result = MahAppsHelper.IsResourceMerged(sources, "Wpf.Ui");

        Assert.True(result);
    }

    [Fact]
    public void IsResourceMerged_NoMatchingFragment_ReturnsFalse()
    {
        Uri?[] sources = [new Uri("http://example/MahApps.Metro/Styles/Controls.xaml")];

        var result = MahAppsHelper.IsResourceMerged(sources, "Wpf.Ui");

        Assert.False(result);
    }

    [Fact]
    public void IsResourceMerged_NullSourceInCollection_DoesNotThrow()
    {
        // ResourceDictionary 可以没有 Source（inline 定义的资源字典），
        // 这种情况下 d.Source 是 null，之前的写法用了 null 传播（?.），
        // 这里验证 IsResourceMerged 同等处理 null 元素不崩、不误判为已合并。
        Uri?[] sources = [null];

        var result = MahAppsHelper.IsResourceMerged(sources, "Wpf.Ui");

        Assert.False(result);
    }

    [Fact]
    public void IsResourceMerged_ExactUriMatch_ReturnsTrue()
    {
        // MahApps 资源用的是精确 URI 匹配（不是 Contains 子串），
        // 验证 IsResourceMerged 同时支持精确匹配这种用法（exactMatch: true）。
        const string uri = "http://example/MahApps.Metro/Styles/Fonts.xaml";
        Uri?[] sources = [new Uri(uri)];

        var result = MahAppsHelper.IsResourceMerged(sources, uri, exactMatch: true);

        Assert.True(result);
    }

    [Fact]
    public void IsResourceMerged_ExactMatchWithDifferentUri_ReturnsFalse()
    {
        const string uri = "http://example/MahApps.Metro/Styles/Fonts.xaml";
        const string otherUri = "http://example/MahApps.Metro/Styles/Controls.xaml";
        Uri?[] sources = [new Uri(otherUri)];

        var result = MahAppsHelper.IsResourceMerged(sources, uri, exactMatch: true);

        Assert.False(result);
    }

    [Fact]
    public void IsResourceMerged_ExactMatchDoesNotMatchSubstring()
    {
        // exactMatch=true 时不能像 Contains 模式那样被子串命中——
        // 回归防护：确保 exactMatch 分支真的走的是 == 而不是不小心又变成 Contains。
        const string fullUri = "http://example/MahApps.Metro/Styles/Fonts.xaml";
        const string fragment = "MahApps.Metro";
        Uri?[] sources = [new Uri(fullUri)];

        var result = MahAppsHelper.IsResourceMerged(sources, fragment, exactMatch: true);

        Assert.False(result);
    }
}
