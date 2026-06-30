using System.Windows;
using NumDesTools;
using NumDesTools.ConflictResolver;
using NumDesTools.UI;
using OfficeOpenXml;

namespace NumDesTools.Scanner;

/// <summary>
/// --conflict-manager 入口：全功能 Excel 冲突管理器（不依赖 Excel 进程）。
/// 从 cwd 向上找 .git 目录，调 ExcelConflictEntry.RunConflictManager。
/// </summary>
internal static class ConflictManager
{
    public static int Run(string[] args)
    {

        // 优先用 --git-root 显式传入（lazygit 用 {{.RepoPath}}），否则从 cwd 向上查找
        string? gitRoot = null;
        int idx = Array.IndexOf(args, "--git-root");
        if (idx >= 0 && idx + 1 < args.Length)
            gitRoot = args[idx + 1];

        gitRoot ??= SvnGitTools.FindGitRoot(Environment.CurrentDirectory);

        if (gitRoot == null)
        {
            Console.Error.WriteLine("错误：当前目录或其父目录不在 Git 仓库中。");
            return 2;
        }

        var thread = new Thread(() =>
        {
            MahAppsHelper.EnsureInitialized();
            ExcelConflictEntry.RunConflictManager(gitRoot, skipHash: true);
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        // WPF 关闭后检查冲突是否全部解决，是则自动提交
        try
        {
            using var repo = new LibGit2Sharp.Repository(gitRoot);
            if (!repo.Index.Conflicts.Any())
            {
                var mergeMsgPath = Path.Combine(gitRoot, ".git", "MERGE_MSG");
                var msg = File.Exists(mergeMsgPath)
                    ? File.ReadAllText(mergeMsgPath)
                    : "解决 xlsx 合并冲突（NumDesTools）";

                var sig = repo.Config.BuildSignature(DateTimeOffset.Now);
                repo.Commit(msg, sig, sig, new LibGit2Sharp.CommitOptions());
                Console.WriteLine("[NumDesTools.Scanner] 冲突已全部解决，已自动提交。");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[NumDesTools.Scanner] 自动提交失败（可手动提交）: {ex.Message}");
        }

        return 0;
    }
}
