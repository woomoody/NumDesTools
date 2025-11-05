using LibGit2Sharp;

namespace NumDesTools;

internal class SvnGitTools
{
    public static List<string> GitDiffFileCount(string path)
    {
        string repoPath = FindGitRoot(path);
        var fileList = new List<string>();
        using var repo = new Repository(repoPath);
        var status = repo.RetrieveStatus();

        foreach (var item in status)
        {
            if (
                item.State == FileStatus.ModifiedInWorkdir
                || item.State == FileStatus.NewInWorkdir
            )
            {
                string fullPath = Path.Combine(repoPath, item.FilePath);
                fileList.Add(fullPath);
            }
        }

        return fileList;
    }

    public static string FindGitRoot(string startPath)
    {
        var directory = new DirectoryInfo(startPath);
        while (directory != null && !Directory.Exists(Path.Combine(directory.FullName, ".git")))
        {
            directory = directory.Parent;
        }
        return directory?.FullName;
    }

    public static bool IsFileModified(string filePath)
    {
        var repoPath = FindGitRoot(filePath);
        if (repoPath == null)
        {
            return false;
        }

        using var repo = new Repository(repoPath);
        var status = repo.RetrieveStatus(filePath);
        // 使用按位与检查是否包含 ModifiedInWorkdir 或 ModifiedInIndex
        return (status & FileStatus.ModifiedInWorkdir) != 0 || (status & FileStatus.ModifiedInIndex) != 0;
    }
    public static (string Name, string Email) GetGitUserInfo()
    {
        string configPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".gitconfig");

        if (!File.Exists(configPath))
            return (null, null);

        var config = Configuration.BuildFrom(configPath);
        return (
            config.Get<string>("user.name")?.Value,
            config.Get<string>("user.email")?.Value
        );
    }
    public static (TimeSpan Delta, DateTime LastCommit) GetLastCommitDelta(string authorName, string repoPath = ".")
    {
        try
        {
            // 确保仓库路径有效
            if (!Repository.IsValid(repoPath))
            {
                throw new ArgumentException($"提供的路径 '{repoPath}' 不是有效的 Git 仓库。");
            }

            using (var repo = new Repository(repoPath))
            {
                // 查询指定作者的最后一次提交
                var lastCommit = repo.Commits.QueryBy(new CommitFilter
                {
                    SortBy = CommitSortStrategies.Time,
                    IncludeReachableFrom = repo.Refs // 查询所有分支（类似 --all）
                }).FirstOrDefault(commit => commit.Author.Name.Contains(authorName)); // 根据需求调整匹配逻辑

                if (lastCommit == null)
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
