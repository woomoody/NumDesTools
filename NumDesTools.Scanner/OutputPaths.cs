using System.Text.Json;

namespace NumDesTools.Scanner;

/// <summary>
/// 统一输出目录管理。根目录读 NumDesGlobalKey.json 的 OutputRootPath，
/// 读不到则 fallback 到 Documents\NumDesOutput。
/// 以后新功能需要写文件：直接用 OutputPaths.Reports / Analysis / Misc，
/// 没有合适子目录时在此处新增一个属性。
/// </summary>
public static class OutputPaths
{
    private static string? _root;

    public static string Root => _root ??= ResolveRoot();

    public static string Reports => Ensure(Path.Combine(Root, "reports"));

    public static string Analysis => Ensure(Path.Combine(Root, "analysis"));

    public static string Misc => Ensure(Path.Combine(Root, "misc"));

    /// <summary>提交 OutputRootPath 下的变更到本地 git 仓库。</summary>
    public static void GitCommit(string message)
    {
        try
        {
            Run("git", $"-C \"{Root}\" add -A");
            Run(
                "git",
                $"-C \"{Root}\" diff --cached --quiet || git -C \"{Root}\" commit -m \"{message.Replace("\"", "\\\"")}\""
            );
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[OutputPaths] git commit 失败（非致命）: {ex.Message}");
        }
    }

    private static string ResolveRoot()
    {
        var keyJson = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesGlobalKey.json"
        );
        if (File.Exists(keyJson))
        {
            try
            {
                var doc = JsonDocument.Parse(File.ReadAllText(keyJson));
                if (
                    doc.RootElement.TryGetProperty("OutputRootPath", out var val)
                    && val.ValueKind == JsonValueKind.String
                    && !string.IsNullOrWhiteSpace(val.GetString())
                )
                    return val.GetString()!;
            }
            catch { }
        }
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesOutput"
        );
    }

    private static string Ensure(string dir)
    {
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Run(string file, string args)
    {
        var psi = new System.Diagnostics.ProcessStartInfo(file, args)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var p = System.Diagnostics.Process.Start(psi);
        p?.WaitForExit(10_000);
    }
}
