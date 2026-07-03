using LibGit2Sharp;

namespace NumDesTools;

internal class SvnGitTools
{
    public static List<string> GitDiffFileCount(string path) =>
        GitDiffAndStagedFiles(path, workdirOnly: true);

    public static List<string> GitDiffAndStagedFiles(string path, bool workdirOnly = false)
    {
        string repoPath = FindGitRoot(path);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var fileList = new List<string>();
        using var repo = new Repository(repoPath);
        var status = repo.RetrieveStatus();

        foreach (var item in status)
        {
            bool isWorkdir =
                item.State == FileStatus.ModifiedInWorkdir || item.State == FileStatus.NewInWorkdir;
            bool isStaged =
                !workdirOnly
                && (
                    (item.State & FileStatus.ModifiedInIndex) != 0
                    || (item.State & FileStatus.NewInIndex) != 0
                    || (item.State & FileStatus.DeletedFromIndex) != 0
                    || (item.State & FileStatus.RenamedInIndex) != 0
                );

            if (isWorkdir || isStaged)
            {
                string fullPath = Path.Combine(repoPath, item.FilePath);
                if (seen.Add(fullPath))
                    fileList.Add(fullPath);
            }
        }

        return fileList;
    }

    // 返回最近 N 条提交的摘要（sha7、作者、时间、message首行）
    public record CommitInfo(
        string Sha,
        string ShortSha,
        string Author,
        DateTime When,
        string Message
    );

    public static List<CommitInfo> GetCommitList(string repoPath, int count = 30)
    {
        repoPath = FindGitRoot(repoPath) ?? repoPath;
        using var repo = new Repository(repoPath);
        return repo
            .Commits.QueryBy(
                new CommitFilter
                {
                    SortBy = CommitSortStrategies.Time,
                    IncludeReachableFrom = repo.Head,
                }
            )
            .Take(count)
            .Select(c => new CommitInfo(
                c.Sha,
                c.Sha[..7],
                c.Author.Name,
                c.Author.When.LocalDateTime,
                c.MessageShort
            ))
            .ToList();
    }

    // 获取某个提交涉及的所有文件（只返回当前仍存在的 xlsx/xls）
    public static List<string> GetCommitFiles(string repoPath, string commitSha)
    {
        repoPath = FindGitRoot(repoPath) ?? repoPath;
        using var repo = new Repository(repoPath);
        var commit =
            repo.Lookup<Commit>(commitSha)
            ?? throw new ArgumentException($"找不到提交 {commitSha}");
        var parent = commit.Parents.FirstOrDefault();
        if (parent is null)
            return [];

        var diff = repo.Diff.Compare<TreeChanges>(parent.Tree, commit.Tree);
        var result = new List<string>();
        foreach (var change in diff)
        {
            if (change.Status == ChangeKind.Deleted)
                continue;
            var fullPath = Path.Combine(
                repoPath,
                change.Path.Replace('/', Path.DirectorySeparatorChar)
            );
            if (File.Exists(fullPath))
                result.Add(fullPath);
        }
        return result;
    }

    // 获取指定作者最近 N 次提交中涉及的所有文件（去重），只返回当前仍存在的文件
    public static List<string> GetRecentAuthorCommitFiles(
        string repoPath,
        string authorName,
        int commitCount
    )
    {
        repoPath = FindGitRoot(repoPath) ?? repoPath;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var fileList = new List<string>();
        using var repo = new Repository(repoPath);

        var commits = repo
            .Commits.QueryBy(
                new CommitFilter
                {
                    SortBy = CommitSortStrategies.Time,
                    IncludeReachableFrom = repo.Head,
                }
            )
            .Where(c => c.Author.Name.Contains(authorName, StringComparison.OrdinalIgnoreCase))
            .Take(commitCount)
            .ToList();

        foreach (var commit in commits)
        {
            var parent = commit.Parents.FirstOrDefault();
            if (parent is null)
                continue;

            var diff = repo.Diff.Compare<TreeChanges>(parent.Tree, commit.Tree);
            foreach (var change in diff)
            {
                if (change.Status == ChangeKind.Deleted)
                    continue;
                var fullPath = Path.Combine(
                    repoPath,
                    change.Path.Replace('/', Path.DirectorySeparatorChar)
                );
                if (File.Exists(fullPath) && seen.Add(fullPath))
                    fileList.Add(fullPath);
            }
        }

        return fileList;
    }

    public static string FindGitRoot(string startPath)
    {
        var gitDir = Repository.Discover(startPath);
        if (gitDir is null)
            return null;
        // Repository.Discover 固定返回带尾部分隔符的路径("...\.git\")，Directory.GetParent
        // 对带尾部分隔符的路径只会去掉斜杠返回原路径本身，不会真的往上一级——不去掉这个
        // 尾部分隔符，FindGitRoot 返回的是 ".git" 目录，不是仓库根目录。
        var trimmed = gitDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var root = Directory.GetParent(trimmed)?.FullName;
        return root;
    }

    public static bool IsFileModified(string filePath)
    {
        var repoPath = FindGitRoot(filePath);
        if (repoPath is null)
        {
            return false;
        }

        using var repo = new Repository(repoPath);
        var status = repo.RetrieveStatus(filePath);
        // 使用按位与检查是否包含 ModifiedInWorkdir 或 ModifiedInIndex
        return (status & FileStatus.ModifiedInWorkdir) != 0
            || (status & FileStatus.ModifiedInIndex) != 0;
    }

    public static (string Name, string Email) GetGitUserInfo()
    {
        string configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".gitconfig"
        );

        if (!File.Exists(configPath))
            return (null, null);

        var config = Configuration.BuildFrom(configPath);
        return (config.Get<string>("user.name")?.Value, config.Get<string>("user.email")?.Value);
    }

    public static (TimeSpan Delta, DateTime LastCommit) GetLastCommitDelta(
        string authorName,
        string repoPath = "."
    )
    {
        try
        {
            // 确保仓库路径有效
            if (!Repository.IsValid(repoPath))
            {
                throw new ArgumentException($"提供的路径 '{repoPath}' 不是有效的 Git 仓库。");
            }

            // 禁用所有权验证
            GlobalSettings.SetOwnerValidation(false);

            using (var repo = new Repository(repoPath))
            {
                // 查询指定作者的最后一次提交
                var lastCommit = repo
                    .Commits.QueryBy(
                        new CommitFilter
                        {
                            SortBy = CommitSortStrategies.Time,
                            IncludeReachableFrom = repo.Refs, // 查询所有分支（类似 --all）
                        }
                    )
                    .FirstOrDefault(commit => commit.Author.Name.Contains(authorName)); // 根据需求调整匹配逻辑

                if (lastCommit is null)
                {
                    throw new Exception($"在仓库中未找到作者 '{authorName}' 的提交记录。");
                }

                TimeSpan delta = DateTime.Now - lastCommit.Author.When.DateTime;
                return (delta, lastCommit.Author.When.DateTime);
            }
        }
        catch (Exception ex)
        {
            // 处理异常，例如仓库无效、未找到提交等
            throw new Exception($"使用 LibGit2Sharp 查询失败: {ex.Message}", ex);
        }
    }
}
