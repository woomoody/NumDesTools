using Microsoft.Office.Interop.Excel;
using NumDesTools.Config;

namespace NumDesTools;

/// <summary>
/// 静态服务定位器——在 AutoOpen 时由 NumDesAddIn.Init() 注入，
/// 其他模块通过此类访问核心依赖，避免直接耦合 NumDesAddIn 类型。
/// </summary>
public static class AppServices
{
    public static Application App { get; private set; } = null!;
    public static GlobalVariable GlobalValue { get; private set; } = null!;
    public static AppConfig Config { get; private set; } = null!;

    internal static void Init(Application app, GlobalVariable globalValue, AppConfig config)
    {
        App = app;
        GlobalValue = globalValue;
        Config = config;
    }
}
