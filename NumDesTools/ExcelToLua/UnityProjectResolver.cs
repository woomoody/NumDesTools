using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;
using NumDesTools;

namespace NumDesTools.ExcelToLua;

/// <summary>
/// 推断并缓存“Excel 表目录(BasePath) → Unity 项目根”的映射。
/// 替代插件端原硬编码 "Code/Assets/..."：别人项目目录不叫 Code 时不再写飞。
/// 首次(或 BasePath 变/路径失效)弹 FolderBrowserDialog 让用户确认一次，结果按 BasePath 缓存到本地 JSON，
/// 同 BasePath 以后直接命中。对齐 Unity 端 Application.dataPath 的“自动指向本项目”语义。
/// </summary>
static class UnityProjectResolver
{
    // 缓存存进 NumDesGlobalKey.json 的 "UnityProjectMap" 字段（复用现有配置，不新开文件）

    public static string Normalize(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return "";
        }
        return Path.GetFullPath(path.TrimEnd('/', '\\'))
            .Replace("\\", "/")
            .TrimEnd('/')
            .ToLowerInvariant();
    }

    public static bool IsUnityProject(string dir) =>
        !string.IsNullOrWhiteSpace(dir)
        && Directory.Exists(dir)
        && Directory.Exists(Path.Combine(dir, "Assets"))
        && Directory.Exists(Path.Combine(dir, "ProjectSettings"));

    public static string Resolve(string basePath = null)
    {
        basePath ??= NumDesAddIn.BasePath;
        if (string.IsNullOrWhiteSpace(basePath))
        {
            return null;
        }

        var key = Normalize(basePath);
        var map = LoadMap();

        // 缓存命中且仍是 Unity 项目 → 直接用
        if (map.TryGetValue(key, out var cached) && IsUnityProject(cached))
        {
            return cached;
        }

        // 推断候选：从 BasePath 往上逐层，在父目录的兄弟里找 Unity 项目
        var candidates = FindUnityProjectCandidates(basePath);

        // 弹窗让用户确认（预填候选或 BasePath 父目录）
        var chosen = PromptForUnityRoot(candidates, basePath);
        if (chosen is null || !IsUnityProject(chosen))
        {
            // 用户取消或选了非 Unity 目录：不缓存、返回失败，EnsureUnityRoot 会中止导表
            return null;
        }

        map[key] = chosen;
        SaveMap(map);
        return chosen;
    }

    internal static List<string> FindUnityProjectCandidates(string basePath)
    {
        var candidates = new List<string>();
        var dir = Path.GetFullPath(basePath);
        for (int i = 0; i < 5 && dir is not null; i++)
        {
            var parent = Path.GetDirectoryName(dir);
            if (parent is null)
            {
                break;
            }
            try
            {
                foreach (var sibling in Directory.GetDirectories(parent))
                {
                    if (IsUnityProject(sibling) && !candidates.Contains(sibling))
                    {
                        candidates.Add(sibling);
                    }
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            dir = parent;
        }
        return candidates;
    }

    static string PromptForUnityRoot(List<string> candidates, string basePath)
    {
        using var dlg = new FolderBrowserDialog();
        dlg.Description = candidates.Count switch
        {
            0 => "未找到 Unity 项目根，请选择 Unity 项目根目录（须含 Assets 和 ProjectSettings）",
            1 => "已推断 Unity 项目根，请确认：",
            _ => $"找到 {candidates.Count} 个 Unity 项目，请确认目标项目根：",
        };

        var initial =
            candidates.Count > 0
                ? candidates[0]
                : Path.GetDirectoryName(Path.GetFullPath(basePath));
        if (initial is not null && Directory.Exists(initial))
        {
            dlg.SelectedPath = initial;
        }

        var ownerHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
        var result =
            ownerHandle != IntPtr.Zero
                ? dlg.ShowDialog(new WindowWrapper(ownerHandle))
                : dlg.ShowDialog();

        return result == DialogResult.OK ? dlg.SelectedPath : null;
    }

    static Dictionary<string, string> LoadMap()
    {
        try
        {
            var json = NumDesAddIn.GlobalValue?.Value.GetValueOrDefault("UnityProjectMap", "");
            return string.IsNullOrEmpty(json)
                ? new Dictionary<string, string>()
                : JsonSerializer.Deserialize<Dictionary<string, string>>(json)
                    ?? new Dictionary<string, string>();
        }
        catch
        {
            return new Dictionary<string, string>();
        }
    }

    static void SaveMap(Dictionary<string, string> map)
    {
        try
        {
            NumDesAddIn.GlobalValue.Value["UnityProjectMap"] = JsonSerializer.Serialize(map);
            NumDesAddIn.GlobalValue.SaveConfig();
        }
        catch
        {
            // 缓存写失败不影响本次导出
        }
    }

    sealed class WindowWrapper(IntPtr handle) : IWin32Window
    {
        public IntPtr Handle => handle;
    }
}
