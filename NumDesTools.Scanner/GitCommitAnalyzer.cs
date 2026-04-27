using LibGit2Sharp;
using System.Text;
using System.Text.RegularExpressions;

namespace NumDesTools.Scanner;

/// <summary>
/// 从飞书工作项的人工输入（标题、描述、评论）中提取 git commit SHA，
/// 用 LibGit2Sharp 获取提交摘要（改了哪些文件 + 提交说明），追加到 AI 分析中。
/// </summary>
public static class GitCommitAnalyzer
{
    // 匹配 7-40 位十六进制 SHA（排除纯数字的活动ID）
    private static readonly Regex RxSha = new(
        @"\b([0-9a-f]{7,40})\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // 配置表 / 代码 git 仓库路径，按顺序搜索
    private static readonly string[] RepoPaths =
    [
        @"C:\M1Work\public",
        @"C:\M1Work\Code",
    ];

    /// <summary>
    /// 从文本（标题 + 描述 + 人工评论）中提取 commit SHA 并分析。
    /// 返回 null 表示没有找到任何有效提交。
    /// </summary>
    public static string? Analyze(string title, string desc, List<string> humanComments)
    {
        var humanText = BuildHumanText(title, desc, humanComments);
        var shas = ExtractShas(humanText);
        if (shas.Count == 0) return null;

        var sb = new StringBuilder();
        foreach (var sha in shas)
        {
            var info = TryGetCommitInfo(sha);
            if (info == null) continue;
            if (sb.Length > 0) sb.AppendLine();
            sb.AppendLine(info);
        }

        return sb.Length == 0 ? null : sb.ToString().TrimEnd();
    }

    // 剥离描述中的 AI 段（--- 分隔线之后），只保留人工部分
    private static string BuildHumanText(string title, string desc, List<string> comments)
    {
        var sepIdx = desc.IndexOf("\n---", StringComparison.Ordinal);
        var humanDesc = sepIdx > 0 ? desc[..sepIdx] : desc;
        if (humanDesc.TrimStart().StartsWith("[AI"))
            humanDesc = string.Empty;

        var parts = new List<string> { title, humanDesc };
        parts.AddRange(comments.Where(c => !c.TrimStart().StartsWith("[AI")));
        return string.Join(" ", parts);
    }

    private static List<string> ExtractShas(string text)
    {
        return RxSha.Matches(text)
            .Select(m => m.Groups[1].Value.ToLowerInvariant())
            .Distinct()
            .Where(s => s.Length >= 7 && !s.All(char.IsDigit))
            .ToList();
    }

    private static string? TryGetCommitInfo(string sha)
    {
        GlobalSettings.SetOwnerValidation(false);

        foreach (var repoPath in RepoPaths)
        {
            if (!Directory.Exists(repoPath) || !Repository.IsValid(repoPath)) continue;

            try
            {
                using var repo = new Repository(repoPath);
                var commit = repo.Lookup<Commit>(sha);
                if (commit == null) continue;

                var sb = new StringBuilder();
                var repoName = Path.GetFileName(repoPath);
                sb.AppendLine($"[{commit.Sha[..8]}] {repoName} · {commit.Author.When.LocalDateTime:yyyy-MM-dd HH:mm} · {commit.Author.Name}");
                sb.AppendLine($"  说明: {commit.MessageShort}");

                // 列出与父提交的差异文件
                var parent = commit.Parents.FirstOrDefault();
                if (parent != null)
                {
                    var changes = repo.Diff.Compare<TreeChanges>(parent.Tree, commit.Tree);
                    var files = changes
                        .Select(c => $"{DiffTypeChar(c.Status)} {c.Path}")
                        .Take(20)
                        .ToList();
                    if (files.Count > 0)
                        sb.AppendLine("  改动文件: " + string.Join(", ", files));
                    if (changes.Count() > 20)
                        sb.AppendLine($"  ... 共 {changes.Count()} 个文件");
                }

                return sb.ToString().TrimEnd();
            }
            catch { /* 该仓库找不到此 SHA，继续下一个 */ }
        }
        return null;
    }

    private static char DiffTypeChar(ChangeKind kind) => kind switch
    {
        ChangeKind.Added    => '+',
        ChangeKind.Deleted  => '-',
        ChangeKind.Modified => '~',
        ChangeKind.Renamed  => '→',
        _                   => '?',
    };
}
