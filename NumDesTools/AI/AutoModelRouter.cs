using System.Text.RegularExpressions;

namespace NumDesTools.AI;

public static class AutoModelRouter
{
    public const string AutoModelName = "🤖 自动";

    public static readonly string[] DefaultModelList =
    {
        AutoModelName,
        "claude-sonnet-5",
        "claude-opus-4-8",
        "deepseek-v4-flash",
        "gemini-3.1-flash-image-preview",
        "gpt-5.5",
        "kimi-k2.6",
    };

    private static readonly (Regex Pattern, string Model, string Reason)[] _rules =
    [
        (
            new Regex(@"翻译|translate|英文|中文互转", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "deepseek-v4-flash",
            "翻译任务"
        ),
        (
            new Regex(@"生成图片|画.{0,4}图|图片", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "gemini-3.1-flash-image-preview",
            "图片生成"
        ),
        (
            new Regex(
                @"设计.{0,4}系统|架构.{0,4}方案|留存.{0,4}分析|数值.{0,4}系统|前提.{0,4}矛盾",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            ),
            "claude-opus-4-8",
            "复杂推理/架构设计"
        ),
        (
            new Regex(
                @"代码|编程|C#|Python|Lua|脚本|算法|重构",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            ),
            "claude-sonnet-5",
            "代码任务"
        ),
        (
            new Regex(@"分析|统计|汇总|趋势|对比|报告", RegexOptions.IgnoreCase | RegexOptions.Compiled),
            "deepseek-v4-flash",
            "数据分析"
        ),
        (
            new Regex(
                @"Excel.{0,4}操作|读取|写入|公式|单元格",
                RegexOptions.IgnoreCase | RegexOptions.Compiled
            ),
            "deepseek-v4-flash",
            "Excel操作"
        ),
    ];

    public static (string Model, string Reason) Route(string input)
    {
        foreach (var (pattern, model, reason) in _rules)
        {
            if (pattern.IsMatch(input))
                return (model, reason);
        }

        return ("deepseek-v4-flash", "通用任务");
    }
}
