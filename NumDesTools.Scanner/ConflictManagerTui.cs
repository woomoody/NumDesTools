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
    private const string QuitChoice = "（退出）";
    private const string BulkResolvePrefix = "（一键解决无真冲突文件";

    /// <summary>一次分类结果：已提取好 blob + 算好 diff，避免选中后重复提取/diff。</summary>
    private readonly record struct Classified(
        string RelPath,
        FileDiff Diff,
        ExcelConflictEntry.ConflictBlobResult Blobs
    );

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

            // 提取 blob + diff 一次性分类：HasTrueConflict=false 的文件三方预选/新增删除默认值已经够用，
            // 不需要人工判断，可以一键接受；只有双方都改同一格的文件才值得进 TUI 逐行看。
            AnsiConsole.MarkupLine("[dim]正在分析各文件冲突情况...[/]");
            var classified = new List<Classified>();
            foreach (var relPath in allXlsx)
            {
                ExcelConflictEntry.ConflictBlobResult? blobs;
                try
                {
                    blobs = ExcelConflictEntry.ExtractConflictBlobs(gitRoot, relPath);
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine(
                        $"[red]提取冲突版本失败：{Markup.Escape(relPath)}: {ex.Message}[/]"
                    );
                    continue;
                }
                if (blobs == null)
                    continue;

                var diff = ExcelConflictDiffer.Diff(
                    blobs.Value.OursPath,
                    blobs.Value.TheirsPath,
                    blobs.Value.BasePath
                );
                classified.Add(new Classified(relPath, diff, blobs.Value));
            }

            var autoResolvable = classified.Where(c => !c.Diff.HasTrueConflict).ToList();
            var needManual = classified.Where(c => c.Diff.HasTrueConflict).ToList();

            var choices = new List<string>();
            var bulkChoiceLabel = $"{BulkResolvePrefix}，共 {autoResolvable.Count} 个）";
            if (autoResolvable.Count > 0)
                choices.Add(bulkChoiceLabel);
            choices.AddRange(needManual.Select(c => c.RelPath));
            choices.AddRange(autoResolvable.Select(c => c.RelPath));
            choices.Add(QuitChoice);

            var chosen = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title(
                        $"[yellow]{classified.Count} 个 xlsx 冲突[/]"
                            + $"（[green]{autoResolvable.Count} 个无真冲突[/] / [red]{needManual.Count} 个需手动[/]），选择："
                    )
                    .PageSize(15)
                    .UseConverter(Markup.Escape) // 冲突文件路径/占位选项里的方括号等字符会被误判为 Markup 标签，统一转义
                    .AddChoices(choices)
            );

            if (chosen == QuitChoice)
            {
                AnsiConsole.MarkupLine("[dim]已退出，冲突未全部解决。[/]");
                return 1;
            }

            if (chosen == bulkChoiceLabel)
            {
                foreach (var c in autoResolvable)
                {
                    var workingPath = Path.Combine(
                        gitRoot,
                        c.RelPath.Replace('/', Path.DirectorySeparatorChar)
                    );
                    ConflictApplier.Apply(c.Diff, workingPath, gitAdd: true);
                    ExcelConflictEntry.FinishGitAdd(gitRoot, c.RelPath);
                    AnsiConsole.MarkupLine(
                        $"[green]✓ 已自动解决（无真冲突）：{Markup.Escape(c.RelPath)}[/]"
                    );
                }
                AnsiConsole.WriteLine();
                continue;
            }

            var picked = classified.First(c => c.RelPath == chosen);
            var pickedWorkingPath = Path.Combine(
                gitRoot,
                picked.RelPath.Replace('/', Path.DirectorySeparatorChar)
            );

            var resolved = ConflictTui.ResolveInteractive(
                picked.Diff,
                outPath: pickedWorkingPath,
                gitAdd: true,
                oursLabel: picked.Blobs.OursLabel,
                theirsLabel: picked.Blobs.TheirsLabel
            );

            // "无差异"或写回成功时 ResolveInteractive 返回 true 但可能未做 git add（无差异分支）→ 补做
            if (resolved)
                ExcelConflictEntry.FinishGitAdd(gitRoot, picked.RelPath);

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
