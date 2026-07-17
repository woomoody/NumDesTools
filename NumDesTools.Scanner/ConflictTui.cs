using LibGit2Sharp;
using NumDesTools;
using NumDesTools.ConflictResolver;
using Spectre.Console;
using Spectre.Console.Rendering;

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
    private const string KeyOurs = "o"; // 死代码（ProcessModified 等）残留引用，下版清理时一并删
    private const string KeyTheirs = "t";
    private const string KeySelect = "s";
    private const string KeyQuit = "q";
    private const string KeyAllOurs = "O";
    private const string KeyAllTheirs = "T";

    /// <summary>切到全屏替代屏幕缓冲区（vim/lazygit 同款 ANSI 转义），退出后原终端内容恢复。</summary>
    internal static void EnterAltScreen() => Console.Write("\x1b[?1049h\x1b[H");

    internal static void ExitAltScreen() => Console.Write("\x1b[?1049l");

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

        EnterAltScreen();
        try
        {
            return ResolveInteractive(
                oursPath,
                theirsPath,
                basePath,
                outPath: oursPath,
                gitAdd: true
            )
                ? 0
                : 2;
        }
        finally
        {
            ExitAltScreen();
        }
    }

    /// <summary>
    /// 交互式解决一个文件的冲突并写回。供 --conflict（CLI直传路径）和
    /// ConflictManagerTui（发现循环+git blob提取后传入临时路径）共用。
    /// 返回 true：无差异或用户确认写回成功；false：用户中途放弃/取消。
    /// </summary>
    internal static bool ResolveInteractive(
        string oursPath,
        string theirsPath,
        string? basePath,
        string outPath,
        bool gitAdd,
        string? oursLabel = null,
        string? theirsLabel = null
    )
    {
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

        return ResolveInteractive(diff, outPath, gitAdd, oursLabel, theirsLabel);
    }

    /// <summary>
    /// 同上，但接受调用方已经算好的 FileDiff（分类阶段就 diff 过一次时用这个重载，避免重复 diff）。
    /// </summary>
    internal static bool ResolveInteractive(
        FileDiff diff,
        string outPath,
        bool gitAdd,
        string? oursLabel = null,
        string? theirsLabel = null
    )
    {
        var allRows = diff
            .Sheets.SelectMany(s => s.Rows.Where(r => r.DiffType != RowDiffType.Same))
            .ToList();

        if (allRows.Count == 0)
        {
            AnsiConsole.MarkupLine("[green]✓ 无差异，文件内容一致。[/]");
            return true;
        }

        // 已经被三方预选（单边改动）/新增删除默认值覆盖的行，不用停下来让人看——直接用算好的默认值。
        // 只有真正"双方都改、选不出默认"的行才值得进向导逐格看，这是本来就该有的行为，之前漏掉了。
        var autoResolved = allRows.Where(r => r.IsResolved).ToList();
        var needsAttention = allRows.Where(r => !r.IsResolved).ToList();

        int modifiedCount = needsAttention.Count(r => r.DiffType == RowDiffType.Modified);
        int onlyCount = needsAttention.Count - modifiedCount;
        var oursTag = oursLabel != null ? $"（我方={oursLabel}）" : "";
        var theirsTag = theirsLabel != null ? $"（对方={theirsLabel}）" : "";

        if (autoResolved.Count > 0)
            AnsiConsole.MarkupLine(
                $"[dim]{autoResolved.Count} 行已被三方预选/新增删除默认值覆盖，直接采用默认值，不需要人工看。[/]"
            );

        if (needsAttention.Count == 0)
        {
            AnsiConsole.MarkupLine("[green]✓ 所有差异都已有默认值，无需人工判断。[/]");
        }
        else
        {
            AnsiConsole.MarkupLine(
                $"[yellow]{needsAttention.Count} 行需要人工判断[/]"
                    + $"  Modified=[cyan]{modifiedCount}[/]  仅一方=[cyan]{onlyCount}[/]"
                    + $"  [dim]{Markup.Escape(oursTag + theirsTag)}[/]"
            );
            AnsiConsole.MarkupLine(
                $"  [dim][[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方"
                    + $"  Enter/[[{KeySelect}]]跳过(用默认)  [[{KeyQuit}]]放弃[/]"
            );
        }
        AnsiConsole.WriteLine();

        // 整表一屏：把所有需要人工判断的冲突格拍平成一张表，集中显示，不用一个个选。
        // 鼠标精确点单元格选版本受终端字符网格精度限制（需真实终端逐格校准坐标），
        // 本版先做键盘：↑↓移光标，o/t 选版本（反色高亮），Enter 确认（未选用默认我方），q 放弃。
        if (!ProcessAllConflictsTable(needsAttention, oursLabel, theirsLabel))
            return false;

        // ── 摘要 + 确认写回 ──────────────────────────────────────────────────
        AnsiConsole.WriteLine();
        RenderSummary(diff);

        var confirm = AnsiConsole.Confirm("确认写回并执行 git add？", defaultValue: true);
        if (!confirm)
        {
            AnsiConsole.MarkupLine("[yellow]已取消，未写入任何文件。[/]");
            return false;
        }

        string? gitLog = null;
        AnsiConsole
            .Status()
            .Start(
                "写回文件...",
                _ =>
                {
                    ConflictApplier.Apply(diff, outPath, gitAdd: gitAdd);
                    if (gitAdd)
                        gitLog = BuildGitAddLog(outPath);
                }
            );

        AnsiConsole.MarkupLine($"[green]✓ 已写回并 git add：{Path.GetFileName(outPath)}[/]");
        if (gitLog != null)
            AnsiConsole.MarkupLine($"  [dim]{Markup.Escape(gitLog)}[/]");
        return true;
    }

    // ── 入口：--conflict-add（一键 git add 所有无冲突 xlsx）─────────────────

    public static int RunConflictAdd(string[] args)
    {
        int idx = Array.IndexOf(args, "--conflict-add");
        var repoRoot =
            idx >= 0 && idx + 1 < args.Length ? args[idx + 1] : Directory.GetCurrentDirectory();

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
                        if (
                            !relPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                            && !relPath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                        )
                            continue;
                        // 无冲突标记：git status 下 Stage == Both 的是真冲突，已被 TUI 处理过的会 stage clean
                        var absPath = Path.Combine(
                            repoRoot,
                            relPath.Replace('/', Path.DirectorySeparatorChar)
                        );
                        skipped.Add(relPath);
                        _ = absPath;
                    }

                    // 用 git status 找已暂存但未提交的（TUI 写回后 git add 过的）
                    foreach (var entry in repo.RetrieveStatus(new StatusOptions()))
                    {
                        if (
                            !entry.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                            && !entry.FilePath.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase)
                        )
                            continue;
                        if (
                            entry.State.HasFlag(FileStatus.ModifiedInWorkdir)
                            && !entry.State.HasFlag(FileStatus.Conflicted)
                        )
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

    // ── Modified 行处理（光标模式，键盘+鼠标）───────────────────────────────

    // BuildModifiedView 的固定竖直布局：Panel顶边框(含标题)=行0，表格顶边框=行1，
    // 表格表头=行2，分隔线=行3，第一条数据行=行4。渲染前必须把光标归零，这个偏移量才成立。
    private const int TableDataStartRow = 4;

    private static int ProcessModified(
        RowConflict row,
        int current,
        int total,
        string? oursLabel,
        string? theirsLabel
    )
    {
        int sel = FindIndex(row.Cells, c => !c.IsExplicit);
        if (sel < 0)
            sel = 0;

        int result = 0;
        try
        {
            // 必须先清屏再归零光标——上一行的画面可能比这一行长，只挪光标不清屏会把
            // 上一屏没覆盖到的残留内容叠在下面，"画面乱叠加"就是这个漏了清屏的锅
            AnsiConsole.Clear();
            Console.SetCursorPosition(0, 0);
        }
        catch { }

        AnsiConsole
            .Live(BuildModifiedView(row, current, total, sel, oursLabel, theirsLabel))
            .Start(ctx =>
            {
                ctx.Refresh(); // Live 默认要等到循环里第一次交互后才显示第一帧——这里强制先刷一次，否则画面"选了才出来"
                while (true)
                {
                    var (isKey, key, col, screenRow) = ConsoleMouseInput.ReadNext();
                    bool done = false;
                    bool mouseHandled = false;

                    if (!isKey)
                    {
                        // 鼠标左键点击：Y 落在某条数据行范围内则选中该格；X 落在窗口左半/右半决定选我方/对方
                        // （表格 Expand 铺满整个终端宽度，colName+我方两列大致占左半，对方+选择占右半，
                        //  不是逐列精确像素边界，是"点左边=我方、点右边=对方"的粗粒度命中，和 lazygit 类似工具的点击体验一致）
                        int cellIdx = screenRow - TableDataStartRow;
                        if (cellIdx >= 0 && cellIdx < row.Cells.Count)
                        {
                            sel = cellIdx;
                            bool clickedLeft = col < Console.WindowWidth / 2;
                            row.Cells[sel].Choice = clickedLeft
                                ? ConflictChoice.Ours
                                : ConflictChoice.Theirs;
                            row.Cells[sel].IsExplicit = true;
                            var next = FindIndex(row.Cells, c => !c.IsExplicit);
                            if (next >= 0)
                                sel = next;
                            mouseHandled = true;
                        }
                    }
                    else if (key.Key == ConsoleKey.UpArrow)
                    {
                        if (sel > 0)
                            sel--;
                    }
                    else if (key.Key == ConsoleKey.DownArrow)
                    {
                        if (sel < row.Cells.Count - 1)
                            sel++;
                    }
                    // Enter / s = 未选格全用默认（Ours）然后确认
                    else if (key.Key == ConsoleKey.Enter || key.KeyChar.ToString() == KeySelect)
                    {
                        foreach (var c in row.Cells.Where(c => !c.IsExplicit))
                        {
                            c.Choice = ConflictChoice.Ours;
                            c.IsExplicit = true;
                        }
                        result = 0;
                        done = true;
                    }
                    else if (!mouseHandled)
                    {
                        switch (key.KeyChar.ToString())
                        {
                            // o/t 只记录选择并把焦点移到下一个待选格；不会因为"刚好选完最后一格"就自动退出——
                            // 退出这一行必须显式按 Enter/s 确认，保留反悔空间（↑↓可以随时回去改已选格的选择）
                            case KeyOurs:
                                row.Cells[sel].Choice = ConflictChoice.Ours;
                                row.Cells[sel].IsExplicit = true;
                                var nextO = FindIndex(row.Cells, c => !c.IsExplicit);
                                if (nextO >= 0)
                                    sel = nextO;
                                break;

                            case KeyTheirs:
                                row.Cells[sel].Choice = ConflictChoice.Theirs;
                                row.Cells[sel].IsExplicit = true;
                                var nextT = FindIndex(row.Cells, c => !c.IsExplicit);
                                if (nextT >= 0)
                                    sel = nextT;
                                break;

                            case KeyAllOurs:
                                row.SetAllCells(ConflictChoice.Ours);
                                result = 0;
                                done = true;
                                break;

                            case KeyAllTheirs:
                                row.SetAllCells(ConflictChoice.Theirs);
                                result = 0;
                                done = true;
                                break;

                            case KeyQuit:
                                result = -1;
                                done = true;
                                break;
                        }
                    }

                    ctx.UpdateTarget(
                        BuildModifiedView(row, current, total, sel, oursLabel, theirsLabel)
                    );
                    ctx.Refresh();
                    if (done)
                        break;
                }
            });

        return result;
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

    /// <summary>总览列表里一行的展示文本。带 RowKey，避免不同行凑巧文案相同时 IndexOf 选错。</summary>
    private static string RowStatusLabel(RowConflict row)
    {
        var status =
            row.DiffType == RowDiffType.Modified
                ? row.IsResolved
                    ? $"✓已选({row.Cells.Count(c => c.Choice == ConflictChoice.Ours)}我方/{row.Cells.Count(c => c.Choice == ConflictChoice.Theirs)}对方)"
                    : $"未选 {row.Cells.Count(c => !c.IsExplicit)}/{row.Cells.Count} 格"
                : row.IsResolved
                    ? $"✓已选:{(row.RowChoice == ConflictChoice.Ours ? "我方" : "对方")}"
                    : "未选";

        var typeStr = row.DiffType switch
        {
            RowDiffType.Modified => "Modified",
            RowDiffType.OnlyOurs => "仅我方",
            RowDiffType.OnlyTheirs => "仅对方",
            _ => row.DiffType.ToString(),
        };

        return $"行 {row.RowKey}  {typeStr}  {status}"
            + (string.IsNullOrEmpty(row.DisplayName) ? "" : $"  {row.DisplayName}");
    }

    // ── OnlyOurs / OnlyTheirs 行处理 ────────────────────────────────────────

    private static int ProcessOnly(
        RowConflict row,
        int current,
        int total,
        string? oursLabel,
        string? theirsLabel
    )
    {
        int result = 0;
        try
        {
            // 必须先清屏再归零光标——上一行的画面可能比这一行长，只挪光标不清屏会把
            // 上一屏没覆盖到的残留内容叠在下面，"画面乱叠加"就是这个漏了清屏的锅
            AnsiConsole.Clear();
            Console.SetCursorPosition(0, 0);
        }
        catch { }

        AnsiConsole
            .Live(BuildOnlyView(row, current, total, oursLabel, theirsLabel))
            .Start(ctx =>
            {
                ctx.Refresh();
                while (true)
                {
                    var (isKey, key, _, _) = ConsoleMouseInput.ReadNext();
                    bool done = false;

                    if (!isKey)
                        continue; // 这个视图没有可点选的目标，鼠标事件直接忽略

                    // Enter / s = 用默认值跳过
                    if (key.Key == ConsoleKey.Enter || key.KeyChar.ToString() == KeySelect)
                    {
                        result = 0;
                        done = true;
                    }
                    else
                    {
                        switch (key.KeyChar.ToString())
                        {
                            case KeyOurs:
                            case KeyAllOurs:
                                row.RowChoice = ConflictChoice.Ours;
                                result = 0;
                                done = true;
                                break;

                            case KeyTheirs:
                            case KeyAllTheirs:
                                row.RowChoice = ConflictChoice.Theirs;
                                result = 0;
                                done = true;
                                break;

                            case KeyQuit:
                                result = -1;
                                done = true;
                                break;
                        }
                    }

                    if (done)
                    {
                        ctx.UpdateTarget(
                            BuildOnlyView(row, current, total, oursLabel, theirsLabel)
                        );
                        ctx.Refresh();
                        break;
                    }
                }
            });

        return result;
    }

    // ── 整表集中显示（所有未自动判定的冲突格一屏）──────────────────────────

    private readonly record struct AllConflictEntry(
        string RowKey,
        string ColName,
        string OursDisplay,
        string TheirsDisplay,
        CellConflict Cell
    );

    /// <summary>把 needsAttention 里每行的"未选格"拍平成一维列表——已三方预选的格不列（简略），只留真需人工的。</summary>
    private static List<AllConflictEntry> FlattenUnresolved(List<RowConflict> needsAttention)
    {
        var list = new List<AllConflictEntry>();
        foreach (var row in needsAttention)
        {
            if (row.DiffType != RowDiffType.Modified)
                continue; // OnlyOurs/OnlyTheirs 默认已解决，不在 needsAttention 里
            foreach (var cell in row.Cells.Where(c => !c.IsExplicit))
                list.Add(
                    new AllConflictEntry(
                        row.RowKey,
                        cell.ColName,
                        cell.OursDisplay,
                        cell.TheirsDisplay,
                        cell
                    )
                );
        }
        return list;
    }

    private static IRenderable BuildAllConflictsView(
        List<AllConflictEntry> entries,
        int cursorRow,
        int cursorCol,
        string? oursLabel,
        string? theirsLabel
    )
    {
        var pending = entries.Count(e => !e.Cell.IsExplicit);
        var title =
            $"[yellow]{pending}/{entries.Count} 个冲突格待选[/]"
            + (pending == 0 ? "  [green]全部已选，Enter 确认写回[/]" : "");

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Expand()
            .AddColumn(new TableColumn("[bold]#[/]"))
            .AddColumn(new TableColumn("[bold]行[/]"))
            .AddColumn(new TableColumn("[bold]列名[/]"))
            .AddColumn(new TableColumn("[blue]我方 (OURS)[/]"))
            .AddColumn(new TableColumn("[yellow]对方 (THEIRS)[/]"))
            .AddColumn(new TableColumn("[bold]已选[/]"));

        for (int i = 0; i < entries.Count; i++)
        {
            var e = entries[i];
            bool selected = e.Cell.IsExplicit;
            bool cursorHere = i == cursorRow;

            // 光标格用 [reverse] 反色高亮（vim 式）；已选非光标格用蓝/黄底块 + ✓
            string oursVal = (cursorHere, cursorCol, selected, e.Cell.Choice) switch
            {
                (true, 0, _, _) => $"[reverse] {Markup.Escape(e.OursDisplay)} ◀[/]",
                (_, _, true, ConflictChoice.Ours) =>
                    $"[bold black on blue] {Markup.Escape(e.OursDisplay)} ✓[/]",
                _ => $"[blue]{Markup.Escape(e.OursDisplay)}[/]",
            };
            string theirsVal = (cursorHere, cursorCol, selected, e.Cell.Choice) switch
            {
                (true, 1, _, _) => $"[reverse] {Markup.Escape(e.TheirsDisplay)} ◀[/]",
                (_, _, true, ConflictChoice.Theirs) =>
                    $"[bold black on yellow] {Markup.Escape(e.TheirsDisplay)} ✓[/]",
                _ => $"[yellow]{Markup.Escape(e.TheirsDisplay)}[/]",
            };
            string choiceStr = !selected
                ? "[dim]未选(默认我方)[/]"
                : (e.Cell.Choice == ConflictChoice.Ours ? "[blue]我方[/]" : "[yellow]对方[/]");
            string prefix = cursorHere ? "[bold green]▶[/] " : "  ";
            table.AddRow(
                $"{prefix}{i + 1}",
                Markup.Escape(e.RowKey),
                Markup.Escape(e.ColName),
                oursVal,
                theirsVal,
                choiceStr
            );
        }

        var cur = entries.Count > 0 ? entries[cursorRow] : default;
        var colWord = cursorCol == 0 ? "我方" : "对方";
        IRenderable curInfo =
            entries.Count > 0
                ? new Markup(
                    $"当前：[bold green]{Markup.Escape(cur.ColName)}[/]  光标在 [bold]{colWord}[/]列  按 [[{KeySelect}]] 选此版本"
                )
                : Text.Empty;
        var legend = BuildLegendLine(oursLabel, theirsLabel);
        var footer = new Markup(
            $"[dim]↑↓←→移动光标  [[{KeySelect}]]选当前列版本(默认我方)  Enter确认(未选默认我方)  [[{KeyQuit}]]放弃[/]"
        );

        var body = new Rows(table, Text.Empty, curInfo, Text.Empty, legend, Text.Empty, footer);
        return new Panel(body)
        {
            Header = new PanelHeader($" {title} "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Grey),
            Expand = true,
        };
    }

    /// <summary>整表主循环：↑↓←→移光标（vim 式二维：行=冲突格，列=我方/对方值），s 选当前光标列版本，Enter 确认，q 放弃。</summary>
    private static bool ProcessAllConflictsTable(
        List<RowConflict> needsAttention,
        string? oursLabel,
        string? theirsLabel
    )
    {
        var entries = FlattenUnresolved(needsAttention);
        if (entries.Count == 0)
            return true;

        int cursorRow = 0;
        int cursorCol = 0; // 0=我方值列, 1=对方值列；默认我方，未选按 s 即选我方
        int result = 0;
        try
        {
            AnsiConsole.Clear();
            Console.SetCursorPosition(0, 0);
        }
        catch { }

        AnsiConsole
            .Live(BuildAllConflictsView(entries, cursorRow, cursorCol, oursLabel, theirsLabel))
            .Start(ctx =>
            {
                ctx.Refresh();
                while (true)
                {
                    var key = Console.ReadKey(intercept: true);
                    bool done = false;

                    if (key.Key == ConsoleKey.UpArrow)
                    {
                        if (cursorRow > 0)
                            cursorRow--;
                    }
                    else if (key.Key == ConsoleKey.DownArrow)
                    {
                        if (cursorRow < entries.Count - 1)
                            cursorRow++;
                    }
                    else if (key.Key == ConsoleKey.LeftArrow)
                    {
                        cursorCol = 0;
                    }
                    else if (key.Key == ConsoleKey.RightArrow)
                    {
                        cursorCol = 1;
                    }
                    else if (key.Key == ConsoleKey.Enter)
                    {
                        foreach (var e in entries.Where(e => !e.Cell.IsExplicit))
                        {
                            e.Cell.Choice = ConflictChoice.Ours;
                            e.Cell.IsExplicit = true;
                        }
                        result = 0;
                        done = true;
                    }
                    else
                    {
                        switch (key.KeyChar.ToString())
                        {
                            // s = 选当前光标所在列的版本：cursorCol=0→我方，1→对方；选后自动下移到下一格
                            case KeySelect:
                                entries[cursorRow].Cell.Choice =
                                    cursorCol == 0 ? ConflictChoice.Ours : ConflictChoice.Theirs;
                                entries[cursorRow].Cell.IsExplicit = true;
                                if (cursorRow < entries.Count - 1)
                                    cursorRow++;
                                break;
                            case KeyQuit:
                                result = -1;
                                done = true;
                                break;
                        }
                    }

                    ctx.UpdateTarget(
                        BuildAllConflictsView(entries, cursorRow, cursorCol, oursLabel, theirsLabel)
                    );
                    ctx.Refresh();
                    if (done)
                        break;
                }
            });

        return result == 0;
    }

    // ── 构建 Modified 视图 ───────────────────────────────────────────────────

    private static IRenderable BuildModifiedView(
        RowConflict row,
        int current,
        int total,
        int sel,
        string? oursLabel,
        string? theirsLabel
    )
    {
        var title =
            $"差异 {current}/{total}  [yellow]Modified[/]  行 [cyan]{Markup.Escape(row.RowKey)}[/]";
        if (!string.IsNullOrEmpty(row.DisplayName))
            title += $"  [dim]{Markup.Escape(row.DisplayName)}[/]";

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Expand()
            .AddColumn(new TableColumn("[bold]列名[/]"))
            .AddColumn(new TableColumn("[blue]我方 (OURS)[/]"))
            .AddColumn(new TableColumn("[yellow]对方 (THEIRS)[/]"))
            .AddColumn(new TableColumn("[bold]选择[/]"));

        for (int i = 0; i < row.Cells.Count; i++)
        {
            var cell = row.Cells[i];
            bool isCursor = i == sel;

            // 选中的一方用反色底块强调，未选中的一方压暗——不用再去读"选择"那一列文字才能确认
            string oursVal,
                theirsVal,
                choiceStr;
            if (!cell.IsExplicit)
            {
                oursVal = $"[blue]{Markup.Escape(cell.OursDisplay)}[/]";
                theirsVal = $"[yellow]{Markup.Escape(cell.TheirsDisplay)}[/]";
                choiceStr = "[dim]待选(默认我方)[/]";
            }
            else if (cell.Choice == ConflictChoice.Ours)
            {
                oursVal = $"[bold black on blue] {Markup.Escape(cell.OursDisplay)} ✓[/]";
                theirsVal = $"[dim]{Markup.Escape(cell.TheirsDisplay)}[/]";
                choiceStr = "[bold blue]我方 ✓[/]";
            }
            else
            {
                oursVal = $"[dim]{Markup.Escape(cell.OursDisplay)}[/]";
                theirsVal = $"[bold black on yellow] {Markup.Escape(cell.TheirsDisplay)} ✓[/]";
                choiceStr = "[bold yellow]对方 ✓[/]";
            }

            var colName = isCursor
                ? $"[bold green]▶ {Markup.Escape(cell.ColName)}[/]"
                : Markup.Escape(cell.ColName);

            table.AddRow(colName, oursVal, theirsVal, choiceStr);
        }

        var cur = row.Cells[sel];
        var curInfo = new Markup(
            $"当前：[bold green]{Markup.Escape(cur.ColName)}[/]  "
                + $"[blue]{Markup.Escape(cur.OursDisplay)}[/] vs [yellow]{Markup.Escape(cur.TheirsDisplay)}[/]"
        );
        var footer = new Markup(
            $"[dim]↑↓移动(可回到已选格改选)  [[{KeyOurs}]]我方  [[{KeyTheirs}]]对方  [[{KeyAllOurs}]]整行我方  [[{KeyAllTheirs}]]整行对方  Enter/[[{KeySelect}]]确认此行(未选格默认我方)  [[{KeyQuit}]]放弃[/]"
        );

        var body = new Rows(
            table,
            Text.Empty,
            curInfo,
            Text.Empty,
            BuildLegendLine(oursLabel, theirsLabel),
            footer
        );
        return new Panel(body)
        {
            Header = new PanelHeader($" {title} "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Grey),
            Expand = true,
        };
    }

    /// <summary>我方(OURS)/对方(THEIRS)具体对应哪个分支/commit——常驻在每一屏的面板里，不是进入全屏前打印一次就再也看不到。</summary>
    private static IRenderable BuildLegendLine(string? oursLabel, string? theirsLabel)
    {
        var ours = oursLabel != null ? Markup.Escape(oursLabel) : "(未知)";
        var theirs = theirsLabel != null ? Markup.Escape(theirsLabel) : "(未知)";
        return new Markup($"[blue]我方(OURS)[/] = {ours}    [yellow]对方(THEIRS)[/] = {theirs}");
    }

    // ── 构建 OnlyOurs / OnlyTheirs 视图 ──────────────────────────────────────

    private static IRenderable BuildOnlyView(
        RowConflict row,
        int current,
        int total,
        string? oursLabel,
        string? theirsLabel
    )
    {
        bool isOurs = row.DiffType == RowDiffType.OnlyOurs;
        var typeLabel = isOurs ? "[blue]仅我方[/]" : "[yellow]仅对方[/]";
        var badge = Markup.Escape(row.DiffTypeBadge);

        var title = $"差异 {current}/{total}  {typeLabel}  行 [cyan]{Markup.Escape(row.RowKey)}[/]";
        if (!string.IsNullOrEmpty(row.DisplayName))
            title += $"  [dim]{Markup.Escape(row.DisplayName)}[/]";

        var parts = new List<IRenderable> { new Markup($"[dim]{badge}[/]"), Text.Empty };

        var src = isOurs ? row.OursFullRow : row.TheirsFullRow;
        if (src != null && row.AllColumns.Count > 0)
        {
            var table = new Table()
                .Border(TableBorder.Simple)
                .Expand()
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

            parts.Add(table);
            parts.Add(Text.Empty);
        }

        var footer = isOurs
            ? $"[dim][[{KeyOurs}/{KeyAllOurs}]]保留此行  [[{KeyTheirs}/{KeyAllTheirs}]]删除此行  Enter/[[{KeySelect}]]跳过(默认保留)  [[{KeyQuit}]]放弃[/]"
            : $"[dim][[{KeyTheirs}/{KeyAllTheirs}]]接受此行  [[{KeyOurs}/{KeyAllOurs}]]拒绝此行  Enter/[[{KeySelect}]]跳过(默认接受)  [[{KeyQuit}]]放弃[/]";
        parts.Add(new Markup(footer));
        parts.Add(Text.Empty);
        parts.Add(BuildLegendLine(oursLabel, theirsLabel));

        var body = new Rows(parts);
        return new Panel(body)
        {
            Header = new PanelHeader($" {title} "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(isOurs ? Color.Blue : Color.Yellow),
            Expand = true,
        };
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
