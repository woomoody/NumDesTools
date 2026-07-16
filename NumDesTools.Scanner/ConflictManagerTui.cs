using LibGit2Sharp;
using NumDesTools;
using NumDesTools.ConflictResolver;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// --conflict-manager-tui 入口：终端版全功能 Excel 冲突管理器。
/// 与 --conflict-manager（WPF 版）共用发现/blob提取/diff/写回引擎，
/// 仅把 WPF 的 picker + 解决窗口换成 Spectre.Console 交互，逐个解决直到冲突清空后自动提交。
/// </summary>
internal static class ConflictManagerTui
{
    private const string QuitChoice = "[退出]";

    public static int Run(string[] args)
    {
        string? gitRoot = null;
        int idx = Array.IndexOf(args, "--git-root");
        if (idx >= 0 && idx + 1 < args.Length)
            gitRoot = args[idx + 1];

        gitRoot ??= SvnGitTools.FindGitRoot(Environment.CurrentDirectory);

        if (gitRoot == null)
        {
            AnsiConsole.MarkupLine("[red]错误：当前目录或其父目录不在 Git 仓库中。[/]");
            return 2;
        }

        const bool skipHash = true; // # 前缀的 xlsm/txt 冲突不进 picker，始终以对方版本为准（与 WPF 版一致）

        while (true)
        {
            List<string> allXlsx;
            try
            {
                using var repo = new Repository(gitRoot);
                var allConflicted = repo
                    .Index.Conflicts.Select(c => c.Ours?.Path ?? c.Theirs?.Path ?? string.Empty)
                    .Where(p => !string.IsNullOrEmpty(p))
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();

                var autoAccept = allConflicted
                    .Where(p =>
                        p.EndsWith(".xll", StringComparison.OrdinalIgnoreCase)
                        || (skipHash && Path.GetFileName(p).Contains('#'))
                    )
                    .ToList();
                foreach (var p in autoAccept)
                    ExcelConflictEntry.AutoAcceptTheirs(repo, gitRoot, p);

                allXlsx = allConflicted
                    .Where(p => p.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .Where(p => !skipHash || !Path.GetFileName(p).Contains('#'))
                    .ToList();
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]读取 Git 状态失败：{ex.Message}[/]");
                return 2;
            }

            if (allXlsx.Count == 0)
                break;

            var choices = allXlsx.Append(QuitChoice).ToList();
            var chosen = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title($"[yellow]{allXlsx.Count} 个 xlsx 仍有冲突，选一个解决：[/]")
                    .PageSize(15)
                    .AddChoices(choices)
            );
            if (chosen == QuitChoice)
            {
                AnsiConsole.MarkupLine("[dim]已退出，冲突未全部解决。[/]");
                return 1;
            }

            var workingFilePath = Path.Combine(
                gitRoot,
                chosen.Replace('/', Path.DirectorySeparatorChar)
            );

            ExcelConflictEntry.ConflictBlobResult? blobs;
            try
            {
                blobs = ExcelConflictEntry.ExtractConflictBlobs(gitRoot, chosen);
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]提取冲突版本失败：{ex.Message}[/]");
                continue;
            }
            if (blobs == null)
            {
                AnsiConsole.MarkupLine($"[red]在 Index 中找不到冲突条目：{chosen}[/]");
                continue;
            }

            var resolved = ConflictTui.ResolveInteractive(
                blobs.Value.OursPath,
                blobs.Value.TheirsPath,
                blobs.Value.BasePath,
                outPath: workingFilePath,
                gitAdd: true,
                oursLabel: blobs.Value.OursLabel,
                theirsLabel: blobs.Value.TheirsLabel
            );

            // "无差异"或写回成功时 ResolveInteractive 返回 true 但可能未做 git add（无差异分支）→ 补做
            if (resolved)
                ExcelConflictEntry.FinishGitAdd(gitRoot, chosen);

            AnsiConsole.WriteLine();
        }

        // 全部解决后自动提交（与 ConflictManager.cs 的 WPF 版一致）
        try
        {
            using var repo = new Repository(gitRoot);
            if (!repo.Index.Conflicts.Any())
            {
                var mergeMsgPath = Path.Combine(gitRoot, ".git", "MERGE_MSG");
                var msg = File.Exists(mergeMsgPath)
                    ? File.ReadAllText(mergeMsgPath)
                    : "解决 xlsx 合并冲突（NumDesTools）";

                var sig = repo.Config.BuildSignature(DateTimeOffset.Now);
                repo.Commit(msg, sig, sig, new CommitOptions());
                AnsiConsole.MarkupLine("[green]✓ 冲突已全部解决，已自动提交。[/]");
            }
        }
        catch (Exception ex)
        {
            AnsiConsole.MarkupLine($"[yellow]自动提交失败（可手动提交）: {ex.Message}[/]");
        }

        return 0;
    }
}
