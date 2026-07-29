using NumDesTools.UI;

namespace NumDesTools.Tests;

/// <summary>
/// MahAppsHelper 资源合并失败消息格式化的单测。
///
/// 背景（见 NumDesOutput/analysis/wpfui-migration-optimization-report-2026-07-29.md P1 第 3 条）：
/// EnsureInitialized() 里 4 处 catch (Exception ex) 之前只调 WriteDebugLog 写入
/// Documents/workspace/xlsx-editor-wpfui-init.log，这个文件没人会主动去看，导致
/// 一次真实的配置错误（"Could not find a theme matching 'Dark.Steel'"）被静默吞掉，
/// 只能靠事后手动 cat 日志文件才发现。
///
/// 复核后发现：这 4 处 catch 全部是真正的失败路径——"资源已存在"的情况在进入 try 之前
/// 已经被 IsResourceMerged 判断过滤掉了，不会走到这些 catch。所以本函数里没有真正
/// "可忽略"的异常分支，报告要求的"区分"落地为：所有失败统一格式化后既写调试日志文件，
/// 也升级到 PluginLog（应用内 PluginLogWindow 可见），不再只靠一个隐蔽文件。
///
/// FormatResourceMergeFailure 是这条链路里唯一的纯函数部分（消息格式化），
/// 实际的双路写入（文件 + PluginLog）在 EnsureInitialized 里调用，属于副作用，不在此单测范围。
/// </summary>
public class MahAppsHelperFailureFormattingTests
{
    [Fact]
    public void FormatResourceMergeFailure_IncludesContextAndExceptionMessage()
    {
        var ex = new InvalidOperationException("boom");

        var result = MahAppsHelper.FormatResourceMergeFailure("wpfui resource merge", ex);

        Assert.Contains("wpfui resource merge", result);
        Assert.Contains("boom", result);
    }

    [Fact]
    public void FormatResourceMergeFailure_IncludesExceptionTypeName()
    {
        // 异常类型名（如 UriFormatException）比只看 Message 更容易定位是"配置写错了"
        // 还是"真的代码 bug"，报告里提到的 ArgumentException("Could not find a theme...")
        // 就是靠类型名 + 消息才能快速定位。
        var ex = new ArgumentException("Could not find a theme matching \"Dark.Steel\"");

        var result = MahAppsHelper.FormatResourceMergeFailure("MahApps ThemeManager", ex);

        Assert.Contains(nameof(ArgumentException), result);
        Assert.Contains("Could not find a theme matching", result);
    }

    [Fact]
    public void FormatResourceMergeFailure_DifferentContextsProduceDifferentMessages()
    {
        var ex = new Exception("same exception");

        var result1 = MahAppsHelper.FormatResourceMergeFailure("context A", ex);
        var result2 = MahAppsHelper.FormatResourceMergeFailure("context B", ex);

        Assert.NotEqual(result1, result2);
        Assert.Contains("context A", result1);
        Assert.Contains("context B", result2);
    }
}
