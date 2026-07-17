using LibGit2Sharp;
using NumDesTools;
using NumDesTools.ConflictResolver;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// --conflict-manager-tui 入口：终端版全功能 Excel 冲突管理器。
/// 与 --conflict-manager（WPF 版）共用发现/blob提取/diff/写回/批量自动解决引擎
/// （<see cref="ExcelConflictEntry.BatchAutoResolve"/>），仅把 WPF 的 picker + 解决窗口
/// 换成 Spectre.Console 交互，逐个解决直到冲突清空后自动提交。
/// </summary>
internal static class ConflictManagerTui
{
    private const string QuitChoice = "（退出）";
    private const string BulkAutoChoice = "⚡ 一键自动解决（扫描可预选文件）";

    public static int Run(string[] args)
    {
        ConflictTui.EnterAltScreen();
        try
        {
            return RunCore(args);
        }
        finally
        {
            ConflictTui.ExitAltScreen();
        }
    }

    private static int RunCore(string[] args)
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

            var choices = new List<string> { BulkAutoChoice };
            choices.AddRange(allXlsx);
            choices.Add(QuitChoice);

            var prompt = new SelectionPrompt<string>()
                .Title($"[yellow]{allXlsx.Count} 个 xlsx 仍有冲突，选择：[/]")
                .PageSize(15)
                .UseConverter(Markup.Escape)
                .AddChoices(choices);
            prompt.WrapAround = true;
            var chosen = AnsiConsole.Prompt(prompt);

            if (chosen == QuitChoice)
            {
                AnsiConsole.MarkupLine("[dim]已退出，冲突未全部解决。[/]");
                return 1;
            }

            if (chosen == BulkAutoChoice)
            {
                List<string> manual = [];
                List<string> errors = [];
                AnsiConsole
                    .Status()
                    .Start(
                        "正在批量自动解决...",
                        ctx =>
                        {
                            ctx.Spinner(Spinner.Known.Dots);
                            var progress = new Progress<(int done, int total, string file)>(p =>
                            {
                                if (p.total > 0)
                                    ctx.Status($"正在批量自动解决... {p.done}/{p.total}");
                            });
                            (manual, errors) = ExcelConflictEntry.BatchAutoResolve(
                                gitRoot,
                                allXlsx,
                                progress
                            );
                        }
                    );

                var autoCount = allXlsx.Count - manual.Count;
                AnsiConsole.MarkupLine(
                    $"[green]✓ 已自动解决 {autoCount} 个[/]"
                        + $"，[yellow]{manual.Count} 个仍需手动（双方均有改动）[/]"
                        + (errors.Count > 0 ? $"，[red]{errors.Count} 个出错[/]" : "")
                );
                foreach (var e in errors)
                    AnsiConsole.MarkupLine($"  [red]{Markup.Escape(e)}[/]");
                AnsiConsole.WriteLine();
                continue;
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
                AnsiConsole.MarkupLine(
                    $"[red]在 Index 中找不到冲突条目：{Markup.Escape(chosen)}[/]"
                );
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
