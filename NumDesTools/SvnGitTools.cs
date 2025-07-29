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

    static string FindGitRoot(string startPath)
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

}
