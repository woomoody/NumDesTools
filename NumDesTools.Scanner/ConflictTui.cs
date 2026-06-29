using NumDesTools.ConflictResolver;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// 终端交互式 xlsx 冲突解决器。
/// 用法：NumDesTools.Scanner --conflict &lt;ours.xlsx&gt; &lt;theirs.xlsx&gt; [base.xlsx]
///
/// 工作流：
///   1. Diff 两个文件 → 收集所有冲突行
///   2. 逐行展示 OURS / THEIRS 对比表格，按键选择
///   3. 所有行处理完毕后调用 ConflictApplier 写回并 git add
/// </summary>
internal static class ConflictTui
{
    private const string KeyOurs = "o";
    private const string KeyTheirs = "t";
    private const string KeySkip = "s";
    private const string KeyQuit = "q";
    private const string KeyAllOurs = "O";
    private const string KeyAllTheirs = "T";

    public static int Run(string[] args)
    {
        // ── 参数解析 ─────────────────────────────────────────────────────────
        int idx = Array.IndexOf(args, "--conflict");
        if (idx < 0 || idx + 2 >= args.Length)
        {
            AnsiConsole.MarkupLine(
                "[red]用法:[/] --conflict <ours.xlsx> <theirs.xlsx> [base.xlsx]"
            );
            return 1;
        }

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

        // ── Diff ─────────────────────────────────────────────────────────────
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

        var conflictRows = diff
            .Sheets.SelectMany(s => s.Rows.Where(r => r.DiffType == RowDiffType.Modified))
            .ToList();

        var totalConflict = diff.TotalConflictRows;
        if (totalConflict == 0)
        {
            AnsiConsole.MarkupLine("[green]✓ 无冲突，文件内容一致。[/]");
            return 0;
        }

        AnsiConsole.MarkupLine(
            $"[yellow]发现 {totalConflict} 行差异[/]（Modified={conflictRows.Count}，其余为仅一方有的行已自动预选）"
        );
        AnsiConsole.MarkupLine(
            $"  [dim]{KeyOurs}=保留我方  {KeyTheirs}=接受对方  {KeySkip}=跳过  {KeyAllOurs}=本行全选我方  {KeyAllTheirs}=本行全选对方  {KeyQuit}=放弃退出[/]"
        );
        AnsiConsole.WriteLine();

        // ── 逐行处理 Modified 冲突 ────────────────────────────────────────────
        int done = 0;
        foreach (var row in conflictRows)
        {
            done++;
            RenderRowConflict(row, done, conflictRows.Count);

            bool rowDone = false;
            while (!rowDone)
            {
                var key = Console.ReadKey(intercept: true).KeyChar.ToString();
                switch (key)
                {
                    case KeyOurs:
                        // 逐格选我方（已有光标的格）
                        var pending = row.Cells.FirstOrDefault(c => !c.IsExplicit);
                        if (pending != null)
                        {
                            pending.Choice = ConflictChoice.Ours;
                            pending.IsExplicit = true;
                        }
                        if (row.IsResolved)
                            rowDone = true;
                        else
                            RenderRowConflict(row, done, conflictRows.Count);
                        break;

                    case KeyTheirs:
                        var pendingT = row.Cells.FirstOrDefault(c => !c.IsExplicit);
                        if (pendingT != null)
                        {
                            pendingT.Choice = ConflictChoice.Theirs;
                            pendingT.IsExplicit = true;
                        }
                        if (row.IsResolved)
                            rowDone = true;
                        else
                            RenderRowConflict(row, done, conflictRows.Count);
                        break;

                    case KeyAllOurs:
                        row.SetAllCells(ConflictChoice.Ours);
                        rowDone = true;
                        break;

                    case KeyAllTheirs:
                        row.SetAllCells(ConflictChoice.Theirs);
                        rowDone = true;
                        break;

                    case KeySkip:
                        // 未选格默认取我方，行标记已解决
                        foreach (var c in row.Cells.Where(c => !c.IsExplicit))
                        {
                            c.Choice = ConflictChoice.Ours;
                            c.IsExplicit = true;
                        }
                        rowDone = true;
                        break;

                    case KeyQuit:
                        AnsiConsole.MarkupLine("[red]已放弃，未写入任何文件。[/]");
                        return 2;
                }
            }

            AnsiConsole.MarkupLine($"  [green]✓ 已处理[/]");
        }

        // ── 摘要 + 确认 ───────────────────────────────────────────────────────
        AnsiConsole.WriteLine();
        RenderSummary(diff);

        var confirm = AnsiConsole.Confirm("确认写回 OURS 文件并执行 git add？", defaultValue: true);
        if (!confirm)
        {
            AnsiConsole.MarkupLine("[yellow]已取消，未写入任何文件。[/]");
            return 3;
        }

        // ── 写回 ─────────────────────────────────────────────────────────────
        AnsiConsole
            .Status()
            .Start(
                "写回文件...",
                _ =>
                {
                    ConflictApplier.Apply(diff, oursPath, gitAdd: true);
                }
            );

        AnsiConsole.MarkupLine($"[green]✓ 已写回并 git add：{Path.GetFileName(oursPath)}[/]");
        return 0;
    }

    // ── 渲染 ──────────────────────────────────────────────────────────────────

    private static void RenderRowConflict(RowConflict row, int current, int total)
    {
        AnsiConsole.Clear();
        var header = $"[bold]冲突 {current}/{total}[/]  行 [cyan]{row.RowKey}[/]";
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

        foreach (var cell in row.Cells)
        {
            var choiceStr = cell.IsExplicit
                ? (cell.Choice == ConflictChoice.Ours ? "[blue]我方[/]" : "[yellow]对方[/]")
                : "[dim]待选[/]";

            table.AddRow(
                Markup.Escape(cell.ColName),
                $"[blue]{Markup.Escape(cell.OursDisplay)}[/]",
                $"[yellow]{Markup.Escape(cell.TheirsDisplay)}[/]",
                choiceStr
            );
        }

        AnsiConsole.Write(table);
        AnsiConsole.WriteLine();

        var pendingCell = row.Cells.FirstOrDefault(c => !c.IsExplicit);
        if (pendingCell != null)
        {
            AnsiConsole.MarkupLine(
                $"当前待选：[bold]{Markup.Escape(pendingCell.ColName)}[/]  "
                    + $"[blue]{Markup.Escape(pendingCell.OursDisplay)}[/] vs [yellow]{Markup.Escape(pendingCell.TheirsDisplay)}[/]"
            );
        }

        AnsiConsole.MarkupLine(
            $"[dim][{KeyOurs}]我方  [{KeyTheirs}]对方  [{KeyAllOurs}]本行全我方  [{KeyAllTheirs}]本行全对方  [{KeySkip}]跳过  [{KeyQuit}]放弃[/]"
        );
    }

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
                var choiceStr =
                    row.DiffType == RowDiffType.Modified
                        ? string.Join(
                            ", ",
                            row.Cells.Select(c =>
                                $"{c.ColName}={(c.Choice == ConflictChoice.Ours ? "我方" : "对方")}"
                            )
                        )
                    : row.RowChoice == ConflictChoice.Ours ? "保留"
                    : "放弃";

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
}
