using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using NumDesTools.ExcelToLua;

namespace NumDesTools;

/// <summary>
/// lua 反推得到的一条列变更事件（来自导出产物 lua.txt 的 git 历史）。
/// OldVal 为 null = 该提交创建了行；NewVal 为 null = 该提交删除了行。
/// </summary>
internal sealed record LuaHistoryEvent(
    string Sha,
    string Date,
    string Author,
    string Msg,
    string? OldVal,
    string? NewVal
);

/// <summary>
/// "谁的锅"加速：用导出产物 lua.txt 的 git 历史反推单元格列变更。
///
/// 原理：xlsx 是二进制，逐版本取一格要解包整份 xlsx（深行 4~8s/提交）；
/// lua.txt 是 Excel 导出的纯文本，每行数据占一行，git pickaxe(-G) 原生行级追踪极快，
/// 且一次覆盖全部历史（不受 xlsx 500 提交窗口限制）。
///
/// 映射关系（导出机制 ExcelExporter/LuaCodeGenerator 的逆过程）：
///   主表 {表名}.lua.txt：拆表时只存 [行键] = _sub_table_id 索引；非拆表则 [行键] = {...} 全量。
///   子表 {表名}_{_sub_table_id}.lua.txt：该行键的整行数据，每行一条。
///
/// 已知限制（设计取舍）：lua 是"导表快照"，滞后于 xlsx 的未导表改动。
/// 因此调用方只把本结果用于"较早历史"，最近改动仍以 xlsx 扫描为准（见 QueryHistoryStreaming）。
/// 任何环节失败都返回 null，由调用方回退纯 xlsx 路径。
/// </summary>
internal static class CellGitHistoryLuaIndex
{
    /// <summary>
    /// 尝试 lua 反推，返回该单元格的列变更历史（新→旧）。失败/未导出返回 null。
    /// </summary>
    public static List<LuaHistoryEvent>? TryQueryHistory(
        string absFilePath,
        string rowKey,
        string colName,
        out string reason
    )
    {
        reason = "";
        try
        {
            var unityRoot = UnityProjectResolver.TryResolveCached(absFilePath);
            if (string.IsNullOrEmpty(unityRoot))
            {
                reason = "无已缓存的 Unity 项目根";
                return null;
            }

            var baseName = Path.GetFileNameWithoutExtension(absFilePath);
            var mainLua = Path.Combine(
                unityRoot,
                "Assets",
                "LuaScripts",
                "Tables",
                $"{baseName}.lua.txt"
            );
            if (!File.Exists(mainLua))
            {
                reason = $"主表 lua 不存在: {Path.GetFileName(mainLua)}";
                return null;
            }

            var targetFile = ResolveTargetLuaFile(mainLua, baseName, rowKey);
            if (targetFile == null)
            {
                reason = "该行未导出到 lua（或非标准结构）";
                return null;
            }

            var relPath = Path.GetRelativePath(unityRoot, targetFile).Replace('\\', '/');

            // pickaxe：找所有改过该行的提交（全分支）
            var pickaxePattern = BuildRowPattern(rowKey);
            var logOut = RunGit(
                unityRoot,
                $"log --all --format=\"%H|%ai|%an|%s\" -G\"{pickaxePattern}\" -- \"{relPath}\""
            );

            var commits = new List<(string sha, string date, string author, string msg)>();
            foreach (var line in logOut.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var p = line.Trim('"').Split('|', 4);
                if (p.Length >= 4 && p[0].Trim().Length >= 8)
                    commits.Add((p[0].Trim(), p[1].Trim(), p[2].Trim(), p[3].Trim()));
            }
            if (commits.Count == 0)
            {
                reason = "pickaxe 未找到改过该行的提交";
                return null;
            }

            var events = new List<LuaHistoryEvent>();
            foreach (var (sha, date, author, msg) in commits)
            {
                var diff = RunGit(unityRoot, $"diff {sha}^ {sha} -- \"{relPath}\"");
                if (string.IsNullOrWhiteSpace(diff)) // 首提交无父
                    diff = RunGit(unityRoot, $"show {sha} --format= -- \"{relPath}\"");

                string? oldVal = null,
                    newVal = null;
                foreach (var dl in diff.Split('\n'))
                {
                    if (dl.Length == 0 || (dl[0] != '-' && dl[0] != '+'))
                        continue;
                    if (!RowLineMatches(dl, rowKey))
                        continue;
                    var v = ExtractFieldValue(dl, colName);
                    if (dl[0] == '-')
                        oldVal = v;
                    else
                        newVal = v;
                }

                // 只记录本列真实变化（含创建/删除）；其它列改动导致的行 diff 在此被滤掉
                if (!string.Equals(oldVal, newVal, StringComparison.Ordinal))
                    events.Add(new LuaHistoryEvent(sha, date, author, msg, oldVal, newVal));
            }

            reason = $"命中 {commits.Count} 个行级提交，本列变更 {events.Count} 次";
            return events;
        }
        catch (Exception ex)
        {
            reason = $"异常: {ex.Message}";
            return null;
        }
    }

    // ── 目标文件定位（主表索引 → 分片）────────────────────────────────────────

    private static string? ResolveTargetLuaFile(string mainLua, string baseName, string rowKey)
    {
        var dir = Path.GetDirectoryName(mainLua)!;
        foreach (var line in File.ReadLines(mainLua))
        {
            if (!RowLineMatches(line, rowKey))
                continue;

            // 拆表索引行：  [行键] = 1101   → 数据在 {baseName}_1101.lua.txt
            var m = Regex.Match(line, @"=\s*(\d+)\s*,?\s*$");
            if (m.Success)
            {
                var shard = Path.Combine(dir, $"{baseName}_{m.Groups[1].Value}.lua.txt");
                return File.Exists(shard) ? shard : null;
            }

            // 非拆表内联行：[行键] = { ... } → 数据就在主表
            if (line.Contains("= {", StringComparison.Ordinal))
                return mainLua;

            return null;
        }
        return null;
    }

    // ── 行/字段解析 ───────────────────────────────────────────────────────────

    private static bool RowLineMatches(string line, string rowKey)
    {
        // 数字键 [11010002] 或字符串键 ["xxx"]
        return line.Contains($"[{rowKey}]", StringComparison.Ordinal)
            || line.Contains($"[\"{rowKey}\"]", StringComparison.Ordinal);
    }

    private static string BuildRowPattern(string rowKey)
    {
        var escaped = Regex.Escape(rowKey);
        return char.IsDigit(rowKey, 0) ? $"\\[{escaped}\\]" : $"\\[\"{escaped}\"\\]";
    }

    /// <summary>
    /// 从一行 `[键] = { a=1, col=v, b={1,2} }` 中取 colName 的原始 lua 值片段。
    /// 列缺失（默认值被导出省略）返回 null。
    /// </summary>
    private static string? ExtractFieldValue(string line, string colName)
    {
        var needle = colName + " = ";
        var idx = line.IndexOf(needle, StringComparison.Ordinal);
        if (idx < 0)
            return null;
        var i = idx + needle.Length;

        var sb = new StringBuilder();
        if (i < line.Length && line[i] == '"') // 字符串：读到未转义的闭引号
        {
            sb.Append('"');
            i++;
            while (i < line.Length)
            {
                var c = line[i];
                if (c == '\\' && i + 1 < line.Length)
                {
                    sb.Append(c).Append(line[i + 1]);
                    i += 2;
                    continue;
                }
                sb.Append(c);
                i++;
                if (c == '"')
                    break;
            }
            return sb.ToString();
        }

        if (i < line.Length && line[i] == '{') // 表/数组：配平花括号
        {
            var depth = 0;
            var start = i;
            while (i < line.Length)
            {
                if (line[i] == '{')
                    depth++;
                else if (line[i] == '}' && --depth == 0)
                {
                    i++;
                    break;
                }
                i++;
            }
            return line[start..i];
        }

        // 裸值：读到 ',' 或 '}'
        var end = i;
        while (end < line.Length && line[end] != ',' && line[end] != '}')
            end++;
        var raw = line[i..end].Trim();
        return raw.Length == 0 ? null : raw;
    }

    // ── git 子进程 ────────────────────────────────────────────────────────────

    private static string RunGit(string workDir, string arguments)
    {
        using var proc = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = FindGitExe(),
                Arguments = arguments,
                WorkingDirectory = workDir,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            },
        };
        proc.Start();
        var stdout = proc.StandardOutput.ReadToEnd();
        proc.StandardError.ReadToEnd();
        proc.WaitForExit(30_000);
        return stdout;
    }

    private static string? _gitExe;

    private static string FindGitExe()
    {
        if (_gitExe != null)
            return _gitExe;
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
        {
            try
            {
                var p = Path.Combine(dir.Trim(), "git.exe");
                if (File.Exists(p))
                    return _gitExe = p;
            }
            catch { }
        }
        return _gitExe = "git";
    }
}
