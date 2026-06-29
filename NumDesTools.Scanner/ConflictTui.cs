using NumDesTools.ConflictResolver;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// 终端交互式 xlsx 冲突解决器。
/// 用法：NumDesTools.Scanner --conflict &lt;ours.xlsx&gt; &lt;theirs.xlsx&gt; [base.xlsx]
///
/// 所有非 Same 行都进入交互队列：
///   Modified     → 逐格选 o/t，O/T 整行，s 跳过
///   OnlyOurs     → o=保留我方行  t=接受对方（删除）  s=默认保留
///   OnlyTheirs   → t=接受对方行  o=拒绝对方（丢弃）  s=默认接受
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

        // 所有非 Same 行都进入队列
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
            $"  [dim][[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方  [[{KeySkip}]]跳过  [[{KeyQuit}]]放弃[/]"
        );
        AnsiConsole.WriteLine();

        // ── 逐行处理 ─────────────────────────────────────────────────────────
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

            AnsiConsole.MarkupLine($"  [green]✓[/]");
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
            .Start("写回文件...", _ => ConflictApplier.Apply(diff, oursPath, gitAdd: true));

        AnsiConsole.MarkupLine($"[green]✓ 已写回并 git add：{Path.GetFileName(oursPath)}[/]");
        return 0;
    }

    // ── Modified 行处理（逐格选）────────────────────────────────────────────

    private static int ProcessModified(RowConflict row, int current, int total)
    {
        RenderModified(row, current, total);
        while (true)
        {
            var key = Console.ReadKey(intercept: true).KeyChar.ToString();
            switch (key)
            {
                case KeyOurs:
                {
                    var cell = row.Cells.FirstOrDefault(c => !c.IsExplicit);
                    if (cell != null)
                    {
                        cell.Choice = ConflictChoice.Ours;
                        cell.IsExplicit = true;
                    }
                    if (row.IsResolved)
                        return 0;
                    RenderModified(row, current, total);
                    break;
                }
                case KeyTheirs:
                {
                    var cell = row.Cells.FirstOrDefault(c => !c.IsExplicit);
                    if (cell != null)
                    {
                        cell.Choice = ConflictChoice.Theirs;
                        cell.IsExplicit = true;
                    }
                    if (row.IsResolved)
                        return 0;
                    RenderModified(row, current, total);
                    break;
                }
                case KeyAllOurs:
                    row.SetAllCells(ConflictChoice.Ours);
                    return 0;

                case KeyAllTheirs:
                    row.SetAllCells(ConflictChoice.Theirs);
                    return 0;

                case KeySkip:
                    // 未选格默认我方
                    foreach (var c in row.Cells.Where(c => !c.IsExplicit))
                    {
                        c.Choice = ConflictChoice.Ours;
                        c.IsExplicit = true;
                    }
                    return 0;

                case KeyQuit:
                    return -1;
            }
        }
    }

    // ── OnlyOurs / OnlyTheirs 行处理（整行选）───────────────────────────────

    private static int ProcessOnly(RowConflict row, int current, int total)
    {
        RenderOnly(row, current, total);
        while (true)
        {
            var key = Console.ReadKey(intercept: true).KeyChar.ToString();
            switch (key)
            {
                case KeyOurs:
                case KeyAllOurs:
                    row.RowChoice = ConflictChoice.Ours;
                    return 0;

                case KeyTheirs:
                case KeyAllTheirs:
                    row.RowChoice = ConflictChoice.Theirs;
                    return 0;

                case KeySkip:
                    // 保持 DefaultRowChoice 不变（已在 Differ 里设好）
                    return 0;

                case KeyQuit:
                    return -1;
            }
        }
    }

    // ── 渲染 Modified ────────────────────────────────────────────────────────

    private static void RenderModified(RowConflict row, int current, int total)
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

        foreach (var cell in row.Cells)
        {
            var choiceStr = cell.IsExplicit
                ? (cell.Choice == ConflictChoice.Ours ? "[blue]我方 ✓[/]" : "[yellow]对方 ✓[/]")
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

        var pending = row.Cells.FirstOrDefault(c => !c.IsExplicit);
        if (pending != null)
            AnsiConsole.MarkupLine(
                $"当前待选：[bold]{Markup.Escape(pending.ColName)}[/]  "
                    + $"[blue]{Markup.Escape(pending.OursDisplay)}[/] vs [yellow]{Markup.Escape(pending.TheirsDisplay)}[/]"
            );

        AnsiConsole.MarkupLine(
            $"[dim][[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方  [[{KeySkip}]]跳过(默认我方)  [[{KeyQuit}]]放弃[/]"
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

        // 显示该行的全部列值
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

        // 根据行类型给出有语义的按键说明
        if (isOurs)
            AnsiConsole.MarkupLine(
                $"[dim][[{KeyOurs}/{KeyAllOurs}]]保留此行  [[{KeyTheirs}/{KeyAllTheirs}]]删除此行(接受对方删除)  [[{KeySkip}]]跳过(默认保留)  [[{KeyQuit}]]放弃[/]"
            );
        else
            AnsiConsole.MarkupLine(
                $"[dim][[{KeyTheirs}/{KeyAllTheirs}]]接受此行  [[{KeyOurs}/{KeyAllOurs}]]拒绝此行  [[{KeySkip}]]跳过(默认接受)  [[{KeyQuit}]]放弃[/]"
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
}
