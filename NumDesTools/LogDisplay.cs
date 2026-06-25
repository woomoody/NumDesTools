namespace NumDesTools;

/// <summary>
/// 兼容旧调用点的静态代理。所有 LogDisplay.RecordLine / Show / Hide
/// 调用无需修改，内部路由到 PluginLog（文件+内存）和 PluginLogWindow（UI）。
/// </summary>
internal static class LogDisplay
{
    /// <summary>
    /// 记录一行日志。兼容调用方自带 [DateTime.Now] 前缀的格式：
    ///   LogDisplay.RecordLine($"[{DateTime.Now}] , {msg}")
    ///   LogDisplay.RecordLine("[{0}] , {1}", DateTime.Now, msg)
    /// </summary>
    public static void RecordLine(string format, params object[] args) =>
        PluginLog.RecordLine(format, args);

    /// <summary>打开日志窗口（若已打开则激活）。</summary>
    public static void Show() => UI.PluginLogWindow.EnsureOpen();

    /// <summary>关闭日志窗口（若已打开）。</summary>
    public static void Hide() => UI.PluginLogWindow.CloseWindow();
}
