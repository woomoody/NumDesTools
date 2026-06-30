using LibGit2Sharp;
using NumDesTools;
using NumDesTools.ConflictResolver;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// 终端交互式 xlsx 冲突解决器。
/// 用法：NumDesTools.Scanner --conflict &lt;ours.xlsx&gt; &lt;theirs.xlsx&gt; [base.xlsx]
///       NumDesTools.Scanner --conflict-add &lt;git-repo-root&gt;  （一键 git add 无冲突 xlsx）
///
/// 所有非 Same 行都进入交互队列：
///   Modified     → ↑↓光标移动，o/t 选格，O/T 整行，Enter 用默认值跳过，s/q 跳过/放弃
///   OnlyOurs     → Enter/s 用默认保留，o=保留，t=接受对方删除，q 放弃
///   OnlyTheirs   → Enter/s 用默认接受，t=接受，o=拒绝，q 放弃
/// </summary>
internal static class ConflictTui
{
    private const string KeyOurs = "o";
    private const string KeyTheirs = "t";
    private const string KeySkip = "s";
    private const string KeyQuit = "q";
    private const string KeyAllOurs = "O";
    private const string KeyAllTheirs = "T";

    // ── 入口：--conflict ────────────────────────────────────────────────────

    public static int Run(string[] args)
    {
        int idx = Array.IndexOf(args, "--conflict");
        if (idx < 0 || idx + 2 >= args.Length)
        {
            AnsiConsole.MarkupLine(
                "[red]用法:[/] --conflict <ours.xlsx> <theirs.xlsx> [base.xlsx]"
            );
            return 1;
        }

        OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");

        var oursPath = args[idx + 1];
        var theirsPath = args[idx + 2];
        var basePath =
            idx + 3 < args.Length && !args[idx + 3].StartsWith('-') ? args[idx + 3] : null;

        if (!File.Exists(oursPath))
        {
            AnsiConsole.MarkupLine($"[red]文件不存在:[/] {oursPath}");
            return 1;
        }
        if (!File.Exists(theirsPath))
        {
            AnsiConsole.MarkupLine($"[red]文件不存在:[/] {theirsPath}");
            return 1;
        }

        FileDiff diff = null!;
        AnsiConsole
            .Status()
            .Start(
                "正在比较文件...",
                ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    diff = ExcelConflictDiffer.Diff(oursPath, theirsPath, basePath);
                }
            );

        var allRows = diff
            .Sheets.SelectMany(s => s.Rows.Where(r => r.DiffType != RowDiffType.Same))
            .ToList();

        if (allRows.Count == 0)
        {
            AnsiConsole.MarkupLine("[green]✓ 无差异，文件内容一致。[/]");
            return 0;
        }

        int modifiedCount = allRows.Count(r => r.DiffType == RowDiffType.Modified);
        int onlyCount = allRows.Count - modifiedCount;
        AnsiConsole.MarkupLine(
            $"[yellow]发现 {allRows.Count} 行差异[/]"
                + $"  Modified=[cyan]{modifiedCount}[/]  仅一方=[cyan]{onlyCount}[/]"
        );
        AnsiConsole.MarkupLine(
            $"  [dim][[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方"
                + $"  Enter/[[{KeySkip}]]跳过(用默认)  [[{KeyQuit}]]放弃[/]"
        );
        AnsiConsole.WriteLine();

        for (int i = 0; i < allRows.Count; i++)
        {
            var row = allRows[i];
            int exitCode =
                row.DiffType == RowDiffType.Modified
                    ? ProcessModified(row, i + 1, allRows.Count)
                    : ProcessOnly(row, i + 1, allRows.Count);

            if (exitCode != 0)
            {
                AnsiConsole.MarkupLine("[red]已放弃，未写入任何文件。[/]");
                return 2;
            }
        }

        // ── 摘要 + 确认写回 ──────────────────────────────────────────────────
        AnsiConsole.WriteLine();
        RenderSummary(diff);

        var confirm = AnsiConsole.Confirm("确认写回 OURS 文件并执行 git add？", defaultValue: true);
        if (!confirm)
        {
            AnsiConsole.MarkupLine("[yellow]已取消，未写入任何文件。[/]");
            return 3;
        }

        string? gitLog = null;
        AnsiConsole
            .Status()
            .Start(
                "写回文件...",
                _ =>
                {
                    ConflictApplier.Apply(diff, oursPath, gitAdd: true);
                    gitLog = BuildGitAddLog(oursPath);
                }
            );

        AnsiConsole.MarkupLine($"[green]✓ 已写回并 git add：{Path.GetFileName(oursPath)}[/]");
        if (gitLog != null)
            AnsiConsole.MarkupLine($"  [dim]{Markup.Escape(gitLog)}[/]");
        return 0;
    }

    // ── 入口：--conflict-add（一键 git add 所有无冲突 xlsx）─────────────────

    public static int RunConflictAdd(string[] args)
    {
        int idx = Array.IndexOf(args, "--conflict-add");
        var repoRoot = idx >= 0 && idx + 1 < args.Length ? args[idx + 1] : Directory.GetCurrentDirectory();

        if (!Directory.Exists(Path.Combine(repoRoot, ".git")))
        {
            AnsiConsole.MarkupLine($"[red]非 git 仓库：[/]{repoRoot}");
            return 1;
        }

        List<string> added = [];
        List<string> skipped = [];

        AnsiConsole
            .Status()
            .Start(
                "扫描冲突状态...",
                ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    using var repo = new Repository(repoRoot);
                    foreach (var entry in repo.Index.Conflicts)
                    {
                        // 冲突入口有 Ancestor/Ours/Theirs，Ours.Path 是相对路径
                        var relPath = (entry.Ours ?? entry.Theirs)?.Path;
                        if (relPath == null)
                            continue;
                        if (!relPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                            && !relPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                            continue;
                        // 无冲突标记：git status 下 Stage == Both 的是真冲突，已被 TUI 处理过的会 stage clean
                        var absPath = Path.Combine(repoRoot, relPath.Replace('/', Path.DirectorySeparatorChar));
                        skipped.Add(relPath);
                        _ = absPath;
                    }

                    // 用 git status 找已暂存但未提交的（TUI 写回后 git add 过的）
                    foreach (var entry in repo.RetrieveStatus(new StatusOptions()))
                    {
                        if (!entry.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                            && !entry.FilePath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (entry.State.HasFlag(FileStatus.ModifiedInWorkdir)
                            && !entry.State.HasFlag(FileStatus.Conflicted))
                        {
                            repo.Index.Add(entry.FilePath);
                            repo.Index.Write();
                            added.Add(entry.FilePath);
                        }
                    }
                }
            );

        if (added.Count == 0 && skipped.Count == 0)
        {
            AnsiConsole.MarkupLine("[green]没有待处理的 xlsx 文件。[/]");
            return 0;
        }

        foreach (var f in added)
            AnsiConsole.MarkupLine($"[green]✓ git add:[/] {f}");
        foreach (var f in skipped)
            AnsiConsole.MarkupLine($"[yellow]⚠ 仍有冲突（跳过）：[/]{f}");

        AnsiConsole.MarkupLine(
            $"\n[bold]完成[/]  已 add {added.Count} 个  仍冲突 {skipped.Count} 个（需先用 --conflict 解决）"
        );
        return 0;
    }

    // ── Modified 行处理（光标模式）──────────────────────────────────────────

    private static int ProcessModified(RowConflict row, int current, int total)
    {
        int sel = FindIndex(row.Cells, c => !c.IsExplicit);
        if (sel < 0)
            sel = 0;

        while (true)
        {
            RenderModified(row, current, total, sel);
            var key = Console.ReadKey(intercept: true);

            if (key.Key == ConsoleKey.UpArrow)
            {
                if (sel > 0)
                    sel--;
                continue;
            }
            if (key.Key == ConsoleKey.DownArrow)
            {
                if (sel < row.Cells.Count - 1)
                    sel++;
                continue;
            }
            // Enter / s = 未选格全用默认（Ours）然后确认
            if (key.Key == ConsoleKey.Enter || key.KeyChar.ToString() == KeySkip)
            {
                foreach (var c in row.Cells.Where(c => !c.IsExplicit))
                {
                    c.Choice = ConflictChoice.Ours;
                    c.IsExplicit = true;
                }
                return 0;
            }

            switch (key.KeyChar.ToString())
            {
                case KeyOurs:
                {
                    row.Cells[sel].Choice = ConflictChoice.Ours;
                    row.Cells[sel].IsExplicit = true;
                    if (row.IsResolved)
                        return 0;
                    var next = FindIndex(row.Cells, c => !c.IsExplicit);
                    if (next >= 0)
                        sel = next;
                    break;
                }
                case KeyTheirs:
                {
                    row.Cells[sel].Choice = ConflictChoice.Theirs;
                    row.Cells[sel].IsExplicit = true;
                    if (row.IsResolved)
                        return 0;
                    var next = FindIndex(row.Cells, c => !c.IsExplicit);
                    if (next >= 0)
                        sel = next;
                    break;
                }
                case KeyAllOurs:
                    row.SetAllCells(ConflictChoice.Ours);
                    return 0;

                case KeyAllTheirs:
                    row.SetAllCells(ConflictChoice.Theirs);
                    return 0;

                case KeyQuit:
                    return -1;
            }
        }
    }

    private static int FindIndex<T>(IEnumerable<T> source, Func<T, bool> predicate)
    {
        int i = 0;
        foreach (var item in source)
        {
            if (predicate(item))
                return i;
            i++;
        }
        return -1;
    }

    // ── OnlyOurs / OnlyTheirs 行处理 ────────────────────────────────────────

    private static int ProcessOnly(RowConflict row, int current, int total)
    {
        RenderOnly(row, current, total);
        while (true)
        {
            var key = Console.ReadKey(intercept: true);
            // Enter / s = 用默认值跳过
            if (key.Key == ConsoleKey.Enter || key.KeyChar.ToString() == KeySkip)
                return 0;

            switch (key.KeyChar.ToString())
            {
                case KeyOurs:
                case KeyAllOurs:
                    row.RowChoice = ConflictChoice.Ours;
                    return 0;

                case KeyTheirs:
                case KeyAllTheirs:
                    row.RowChoice = ConflictChoice.Theirs;
                    return 0;

                case KeyQuit:
                    return -1;
            }
        }
    }

    // ── 渲染 Modified ────────────────────────────────────────────────────────

    private static void RenderModified(RowConflict row, int current, int total, int sel)
    {
        AnsiConsole.Clear();
        var header =
            $"[bold]差异 {current}/{total}[/]  [yellow]Modified[/]  行 [cyan]{Markup.Escape(row.RowKey)}[/]";
        if (!string.IsNullOrEmpty(row.DisplayName))
            header += $"  [dim]{Markup.Escape(row.DisplayName)}[/]";
        AnsiConsole.MarkupLine(header);
        AnsiConsole.WriteLine();

        var table = new Table()
            .Border(TableBorder.Rounded)
            .AddColumn(new TableColumn("[bold]列名[/]"))
            .AddColumn(new TableColumn("[blue]我方 (OURS)[/]"))
            .AddColumn(new TableColumn("[yellow]对方 (THEIRS)[/]"))
            .AddColumn(new TableColumn("[bold]选择[/]"));

        for (int i = 0; i < row.Cells.Count; i++)
        {
            var cell = row.Cells[i];
            bool isCursor = i == sel;

            var choiceStr = cell.IsExplicit
                ? (cell.Choice == ConflictChoice.Ours ? "[blue]我方 ✓[/]" : "[yellow]对方 ✓[/]")
                : "[dim]待选(默认我方)[/]";

            var colName = isCursor
                ? $"[bold green]▶ {Markup.Escape(cell.ColName)}[/]"
                : Markup.Escape(cell.ColName);

            table.AddRow(
                colName,
                $"[blue]{Markup.Escape(cell.OursDisplay)}[/]",
                $"[yellow]{Markup.Escape(cell.TheirsDisplay)}[/]",
                choiceStr
            );
        }

        AnsiConsole.Write(table);
        AnsiConsole.WriteLine();

        var cur = row.Cells[sel];
        AnsiConsole.MarkupLine(
            $"当前：[bold green]{Markup.Escape(cur.ColName)}[/]  "
                + $"[blue]{Markup.Escape(cur.OursDisplay)}[/] vs [yellow]{Markup.Escape(cur.TheirsDisplay)}[/]"
        );
        AnsiConsole.MarkupLine(
            $"[dim]↑↓移动  [[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方  Enter/[[{KeySkip}]]跳过(默认我方)  [[{KeyQuit}]]放弃[/]"
        );
    }

    // ── 渲染 OnlyOurs / OnlyTheirs ───────────────────────────────────────────

    private static void RenderOnly(RowConflict row, int current, int total)
    {
        AnsiConsole.Clear();

        bool isOurs = row.DiffType == RowDiffType.OnlyOurs;
        var typeLabel = isOurs ? "[blue]仅我方[/]" : "[yellow]仅对方[/]";
        var badge = Markup.Escape(row.DiffTypeBadge);

        var header =
            $"[bold]差异 {current}/{total}[/]  {typeLabel}  行 [cyan]{Markup.Escape(row.RowKey)}[/]";
        if (!string.IsNullOrEmpty(row.DisplayName))
            header += $"  [dim]{Markup.Escape(row.DisplayName)}[/]";
        AnsiConsole.MarkupLine(header);
        AnsiConsole.MarkupLine($"  [dim]{badge}[/]");
        AnsiConsole.WriteLine();

        var src = isOurs ? row.OursFullRow : row.TheirsFullRow;
        if (src != null && row.AllColumns.Count > 0)
        {
            var table = new Table()
                .Border(TableBorder.Simple)
                .AddColumn(new TableColumn("[bold]列名[/]"))
                .AddColumn(new TableColumn(isOurs ? "[blue]我方值[/]" : "[yellow]对方值[/]"));

            foreach (var col in row.AllColumns)
            {
                if (!src.TryGetValue(col, out var v))
                    continue;
                var val = v?.ToString() ?? "(空)";
                var colored = isOurs
                    ? $"[blue]{Markup.Escape(val)}[/]"
                    : $"[yellow]{Markup.Escape(val)}[/]";
                table.AddRow(Markup.Escape(col), colored);
            }

            AnsiConsole.Write(table);
            AnsiConsole.WriteLine();
        }

        if (isOurs)
            AnsiConsole.MarkupLine(
                $"[dim][[{KeyOurs}/{KeyAllOurs}]]保留此行  [[{KeyTheirs}/{KeyAllTheirs}]]删除此行  Enter/[[{KeySkip}]]跳过(默认保留)  [[{KeyQuit}]]放弃[/]"
            );
        else
            AnsiConsole.MarkupLine(
                $"[dim][[{KeyTheirs}/{KeyAllTheirs}]]接受此行  [[{KeyOurs}/{KeyAllOurs}]]拒绝此行  Enter/[[{KeySkip}]]跳过(默认接受)  [[{KeyQuit}]]放弃[/]"
            );
    }

    // ── 摘要 ─────────────────────────────────────────────────────────────────

    private static void RenderSummary(FileDiff diff)
    {
        AnsiConsole.MarkupLine("[bold]处理摘要[/]");
        var table = new Table()
            .Border(TableBorder.Simple)
            .AddColumn("Sheet")
            .AddColumn("行 ID")
            .AddColumn("类型")
            .AddColumn("选择");

        foreach (var sheet in diff.Sheets)
        {
            foreach (var row in sheet.Rows.Where(r => r.DiffType != RowDiffType.Same))
            {
                string choiceStr;
                if (row.DiffType == RowDiffType.Modified)
                {
                    choiceStr = string.Join(
                        ", ",
                        row.Cells.Select(c =>
                            $"{c.ColName}={(c.Choice == ConflictChoice.Ours ? "我方" : "对方")}"
                        )
                    );
                }
                else
                {
                    bool keep =
                        row.DiffType == RowDiffType.OnlyOurs
                            ? row.RowChoice == ConflictChoice.Ours
                            : row.RowChoice == ConflictChoice.Theirs;
                    choiceStr = keep ? "✓ 保留" : "✗ 丢弃";
                }

                table.AddRow(
                    Markup.Escape(sheet.SheetName),
                    Markup.Escape(row.RowKey),
                    Markup.Escape(row.DiffTypeBadge),
                    Markup.Escape(choiceStr)
                );
            }
        }

        AnsiConsole.Write(table);
    }

    // ── 生成 git add 日志摘要 ────────────────────────────────────────────────

    private static string? BuildGitAddLog(string filePath)
    {
        try
        {
            var repoRoot = SvnGitTools.FindGitRoot(filePath);
            if (repoRoot == null)
                return null;
            using var repo = new Repository(repoRoot);
            var rel = Path.GetRelativePath(repoRoot, filePath).Replace('\\', '/');
            var status = repo.RetrieveStatus(rel);
            return $"git status: {status}  repo={Path.GetFileName(repoRoot)}";
        }
        catch
        {
            return null;
        }
    }
}
