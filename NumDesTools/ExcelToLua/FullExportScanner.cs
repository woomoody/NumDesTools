using System.IO;
using System.Diagnostics;

namespace NumDesTools.ExcelToLua;

/// <summary>
/// 全表导出辅助：扫描 Excels 根下 3 个子目录（Localizations/Tables/UIs）所有 Excel，
/// 清空 Tables 输出目录（保留 NonOutputTable 子文件夹），git add 导出产物。
/// 纯 IO + 进程调用，不依赖 Excel COM，可单测。
/// </summary>
public static class FullExportScanner
{
    private static readonly string[] _subDirs = { "Localizations", "Tables", "UIs" };

    /// <summary>扫描 excelsRoot 下 Localizations/Tables/UIs 三目录的所有 .xls/.xlsx 文件，跳过 # / ~ 前缀的隐藏/WIP 表。</summary>
    public static List<string> ScanAllExcels(string excelsRoot)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(excelsRoot) || !Directory.Exists(excelsRoot))
            return result;

        foreach (var sub in _subDirs)
        {
            var dir = Path.Combine(excelsRoot, sub);
            if (!Directory.Exists(dir))
                continue;
            foreach (var f in Directory.EnumerateFiles(dir, "*.xls", SearchOption.AllDirectories))
            {
                if (IsExportable(f))
                    result.Add(f);
            }
            foreach (var f in Directory.EnumerateFiles(dir, "*.xlsx", SearchOption.AllDirectories))
            {
                if (IsExportable(f) && !result.Contains(f))
                    result.Add(f);
            }
        }
        return result;
    }

    /// <summary>对齐 GitExportSelectWindow.IsExportable：跳过文件名或路径段含 # 或 ~ 的隐藏/WIP 表。</summary>
    static bool IsExportable(string file)
    {
        var name = Path.GetFileNameWithoutExtension(file);
        if (name.Contains('#') || name.Contains('~'))
            return false;
        // 路径段含 # / ~（如 #Hidden 目录）也跳过
        var rel = file.Replace('\\', '/');
        return !rel.Contains("/#") && !rel.Contains("/~");
    }

    /// <summary>
    /// 清空 Tables 输出目录下的 .txt 文件（lua 导出产物），保留 .meta（Unity GUID，导出器不重写）+ NonOutputTable 子文件夹。
    /// 死文件 meta 由 PruneOrphanMetas 在导出后单独剪枝。
    /// </summary>
    public static void CleanTablesOutput(string tablesOutputDir)
    {
        if (string.IsNullOrWhiteSpace(tablesOutputDir) || !Directory.Exists(tablesOutputDir))
            return;

        foreach (var f in Directory.EnumerateFiles(tablesOutputDir, "*.txt", SearchOption.TopDirectoryOnly))
        {
            try { File.Delete(f); }
            catch { /* 单文件删失败不阻塞 */ }
        }
    }

    /// <summary>
    /// 剪孤儿 meta：枚举 Tables 顶层 *.meta，若同名 txt（去 .meta 后缀）不存在=该表已死（Excel 没了）→ 删 meta。
    /// 活跃表 txt 已在导出阶段写回，其 meta 全程不动（GUID 保留）。
    /// </summary>
    public static void PruneOrphanMetas(string tablesOutputDir)
    {
        if (string.IsNullOrWhiteSpace(tablesOutputDir) || !Directory.Exists(tablesOutputDir))
            return;

        foreach (var meta in Directory.EnumerateFiles(tablesOutputDir, "*.meta", SearchOption.TopDirectoryOnly))
        {
            // xxx.lua.txt.meta → 对应 txt = xxx.lua.txt（去末尾 .meta）
            var txt = meta.Substring(0, meta.Length - ".meta".Length);
            if (!File.Exists(txt))
            {
                try { File.Delete(meta); }
                catch { /* 单文件删失败不阻塞 */ }
            }
        }
    }

    /// <summary>
    /// git -C unityRoot add -A Assets/LuaScripts/Tables Assets/LuaScripts/Localizations
    /// 标记 Tables（删除+新增）和 Localizations（导出产物）给 git。返回是否成功。
    /// </summary>
    public static bool GitAddTablesAndLocalizations(string unityRoot)
    {
        if (string.IsNullOrWhiteSpace(unityRoot) || !Directory.Exists(unityRoot))
            return false;

        try
        {
            RunGit(unityRoot, "add -A Assets/LuaScripts/Tables Assets/LuaScripts/Localizations");
            return true;
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[FullExportScanner] git add 失败: {ex.Message}");
            return false;
        }
    }

    static void RunGit(string workDir, string args)
    {
        var psi = new ProcessStartInfo("git", args)
        {
            WorkingDirectory = workDir,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var p = Process.Start(psi);
        if (p != null && !p.WaitForExit(30_000))
            p.Kill();
    }
}
