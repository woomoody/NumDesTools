using System.Text.Json;
using System.Text.Json.Serialization;

namespace NumDesTools.ConflictResolver;

/// <summary>
/// Rust TUI 混合架构的启动/回读逻辑：序列化 FileDiff → 起子进程 conflict-tui.exe → 读结果合并回 diff。
/// 从 NumDesTools.Scanner 提到 Core，供 Scanner（--conflict）和 WPF 插件（Ribbon）共用，不重复一份。
/// 不含任何 UI/控制台依赖——调用方（Spectre 版打印，WPF 版弹 MessageBox）自己决定怎么呈现结果/报错。
/// </summary>
public static class RustTuiLauncher
{
    /// <summary>找 Rust conflict-tui.exe。优先调用方自己 exe 所在目录旁（发布后跟着走），
    /// 其次 tools\conflict-tui\target\release\（开发环境兜底，脱离源码树后这条路径就没了）。</summary>
    public static string? FindRustTuiExe()
    {
        var candidates = new[]
        {
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "conflict-tui.exe"),
            Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "tools",
                "conflict-tui",
                "target",
                "release",
                "conflict-tui.exe"
            ),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".local",
                "bin",
                "conflict-tui.exe"
            ),
        };
        foreach (var p in candidates)
        {
            var full = Path.GetFullPath(p);
            if (File.Exists(full))
                return full;
        }
        return null;
    }

    /// <summary>走 Rust TUI：序列化 FileDiff → 起子进程 → 读 result.json 合并 selections。
    /// 返回 true=用户确认（diff 已通过 ApplySelections 就地合并选择）；false=取消/失败（error 有值时是失败）。</summary>
    public static bool TryResolve(
        FileDiff diff,
        string rustTuiPath,
        string? oursLabel,
        string? theirsLabel,
        out string? error
    )
    {
        error = null;
        var tempDir = Path.Combine(Path.GetTempPath(), "NumDesConflictTui");
        Directory.CreateDirectory(tempDir);
        var diffPath = Path.Combine(tempDir, $"{Guid.NewGuid():N}_diff.json");
        var resultPath = diffPath.Replace("diff.json", "result.json");

        try
        {
            // 序列化（UTF8 无 BOM）
            var dto = diff.ToDto(oursLabel, theirsLabel);
            var json = dto.ToJson();
            File.WriteAllText(diffPath, json, new System.Text.UTF8Encoding(false));

            // UseShellExecute=false：直接继承父进程控制台/窗口站，不需要另开窗口。
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = rustTuiPath,
                Arguments = $"\"{diffPath}\"",
                UseShellExecute = false,
            };
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc is null)
            {
                error = "启动 Rust TUI 失败";
                return false;
            }
            proc.WaitForExit();

            if (proc.ExitCode != 0)
                return false; // 用户按 q 放弃，不算错误

            if (!File.Exists(resultPath))
            {
                error = "Rust TUI 未生成 result.json";
                return false;
            }
            var resultJson = File.ReadAllText(resultPath);
            var result = JsonSerializer.Deserialize<SelectionResultDto>(
                resultJson,
                new JsonSerializerOptions
                {
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                    Converters = { new JsonStringEnumConverter() },
                }
            );
            if (result is null || !result.Confirmed)
                return false;

            diff.ApplySelections(result);
            return true;
        }
        catch (Exception ex)
        {
            error = $"Rust TUI 调用失败：{ex.Message}";
            return false;
        }
        finally
        {
            if (Environment.GetEnvironmentVariable("NUMDES_KEEP_TUI_TEMP") != "1")
            {
                try
                {
                    if (File.Exists(diffPath))
                        File.Delete(diffPath);
                    if (File.Exists(resultPath))
                        File.Delete(resultPath);
                }
                catch { }
            }
        }
    }
}
