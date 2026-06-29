using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using NumDesTools.ExcelIndex;
using Spectre.Console;

[assembly: InternalsVisibleTo("NumDesTools.Tests")]

namespace NumDesTools.Scanner;

/// <summary>
/// 终端交互式 Excel 索引搜索（fzf 风格）。
/// 用法：NumDesTools.Scanner --search [--index &lt;path.json.gz&gt;]
///
/// 工作流：
///   1. 加载 excel_index_*.json（自动发现 Documents\NumDesTools 下的索引文件）
///   2. 用户实时输入关键词 → 动态过滤命中项
///   3. 上下键选择结果，Enter 打开文件，Tab 切换前缀/包含模式，Esc 退出
/// </summary>
internal static class SearchTui
{
    private const int MaxDisplayRows = 20;
    private static int PageSize => Math.Max(5, Math.Min(MaxDisplayRows, Console.WindowHeight - 6));

    public static int Run(string[] args)
    {
        // ── 参数解析 ─────────────────────────────────────────────────────────
        string? indexPath = null;
        int idxFlag = Array.IndexOf(args, "--index");
        if (idxFlag >= 0 && idxFlag + 1 < args.Length)
            indexPath = args[idxFlag + 1];

        // ── 自动发现索引 ─────────────────────────────────────────────────────
        string[] indexFiles;
        if (indexPath != null)
        {
            if (!File.Exists(indexPath))
            {
                AnsiConsole.MarkupLine($"[red]索引文件不存在：[/]{indexPath}");
                return 1;
            }
            indexFiles = [indexPath];
        }
        else
        {
            // 1. 从 cwd 向上找 Excels 目录，找到则按需构建索引
            var excelsRoot = FindExcelsRoot(Environment.CurrentDirectory);
            if (excelsRoot != null)
            {
                var autoIndexPath = ExcelSearchIndex.GetIndexPath(excelsRoot);
                var existing = ExcelSearchIndex.LoadFromDisk(autoIndexPath);
                if (existing == null || string.IsNullOrEmpty(existing.ExcelsRoot))
                {
                    AnsiConsole.MarkupLine(
                        $"[yellow]首次使用，正在构建索引...[/] ({Markup.Escape(excelsRoot)})"
                    );
                    ExcelSearchIndex? built = null;
                    AnsiConsole
                        .Progress()
                        .AutoClear(false)
                        .HideCompleted(false)
                        .Start(ctx =>
                        {
                            var task = ctx.AddTask("[green]扫描文件[/]", maxValue: 100);
                            var builder = new ExcelIndexBuilder(excelsRoot);
                            built = builder.Build(
                                existing,
                                new Progress<(int done, int total)>(p =>
                                {
                                    if (p.total > 0)
                                        task.Value = (double)p.done / p.total * 100;
                                })
                            );
                            task.Value = 100;
                        });
                    built!.BuildSortedKeys();
                    built.SaveToDisk(autoIndexPath);
                    AnsiConsole.MarkupLine(
                        $"[green]✓ 索引构建完成[/]  {built.Exact.Count} 个值  {built.Files.Count} 个文件"
                    );
                }
                indexFiles = [autoIndexPath];
            }
            else
            {
                // 2. fallback：扫 Documents\NumDesTools 下已有索引
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "NumDesTools"
                );
                indexFiles = Directory.Exists(dir)
                    ? Directory.GetFiles(dir, "excel_index_*.json", SearchOption.TopDirectoryOnly)
                    : [];
                if (indexFiles.Length == 0)
                {
                    AnsiConsole.MarkupLine(
                        "[red]未找到索引文件，且当前目录下也没有 Excels 子目录。[/]"
                    );
                    AnsiConsole.MarkupLine(
                        "[dim]请在含 Excels/ 的 git 仓库目录下运行，或先在 Excel 插件里执行「构建搜索索引」[/]"
                    );
                    return 1;
                }
            }
        }

        // ── 加载索引 ──────────────────────────────────────────────────────────
        ExcelSearchIndex? idx = null;
        AnsiConsole
            .Status()
            .Start(
                "加载索引...",
                ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    foreach (var f in indexFiles)
                    {
                        var loaded = ExcelSearchIndex.LoadFromDisk(f);
                        if (loaded == null)
                            continue;
                        if (idx == null)
                        {
                            idx = loaded;
                        }
                        else
                        {
                            foreach (var (k, hits) in loaded.Exact)
                            {
                                if (!idx.Exact.TryGetValue(k, out var list))
                                    idx.Exact[k] = hits;
                                else
                                    list.AddRange(hits);
                            }
                            int fileOffset = idx.Files.Count;
                            int sheetOffset = idx.Sheets.Count;
                            idx.Files.AddRange(loaded.Files);
                            idx.Sheets.AddRange(loaded.Sheets);
                            foreach (var (fid, sid) in loaded.AllSheets)
                                idx.AllSheets.Add((fid + fileOffset, sid + sheetOffset));
                        }
                    }
                    idx?.RebuildLookups();
                    idx?.BuildSortedKeys();
                }
            );

        if (idx == null)
        {
            AnsiConsole.MarkupLine("[red]索引加载失败。[/]");
            return 1;
        }

        AnsiConsole.MarkupLine(
            $"[green]索引就绪[/]  {idx.Exact.Count:N0} 个值 · {idx.Files.Count} 个文件  "
                + $"[dim](构建于 {idx.BuiltAt:yyyy-MM-dd HH:mm})[/]"
        );
        AnsiConsole.MarkupLine(
            "[dim]输入关键词实时搜索 | ↑↓/PgUp/PgDn 翻页 | Enter 打开文件 | Tab 切换模式 | Esc 退出[/]"
        );
        AnsiConsole.WriteLine();

        // ── 主循环 ────────────────────────────────────────────────────────────
        var query = string.Empty;
        var selectedIdx = 0;
        var usePrefix = false;
        List<SearchResult> results = [];
        string? statusMsg = null;

        while (true)
        {
            Render(query, results, selectedIdx, usePrefix, statusMsg);
            statusMsg = null;

            var key = Console.ReadKey(intercept: true);

            switch (key.Key)
            {
                case ConsoleKey.Escape:
                    AnsiConsole.Clear();
                    AnsiConsole.MarkupLine("[dim]已退出搜索。[/]");
                    return 0;

                case ConsoleKey.Enter:
                    if (results.Count > 0 && selectedIdx < results.Count)
                        statusMsg = OpenFile(idx, results[selectedIdx]);
                    continue;

                case ConsoleKey.Tab:
                    usePrefix = !usePrefix;
                    selectedIdx = 0;
                    results = DoSearch(idx, query, usePrefix);
                    continue;

                case ConsoleKey.UpArrow:
                    if (selectedIdx > 0)
                        selectedIdx--;
                    continue;

                case ConsoleKey.DownArrow:
                    if (selectedIdx < results.Count - 1)
                        selectedIdx++;
                    continue;

                case ConsoleKey.PageUp:
                    selectedIdx = Math.Max(0, selectedIdx - PageSize);
                    continue;

                case ConsoleKey.PageDown:
                    selectedIdx = Math.Min(results.Count - 1, selectedIdx + PageSize);
                    continue;

                case ConsoleKey.Backspace:
                    if (query.Length > 0)
                        query = query[..^1];
                    break;

                default:
                    if (!char.IsControl(key.KeyChar))
                        query += key.KeyChar;
                    break;
            }

            selectedIdx = 0;
            results = DoSearch(idx, query, usePrefix);
        }
    }

    // ── 打开文件 ──────────────────────────────────────────────────────────────

    private static string OpenFile(ExcelSearchIndex idx, SearchResult r)
    {
        var root = idx.ExcelsRoot;
        if (string.IsNullOrEmpty(root))
            return "[red]索引未含根目录信息，请在 Excel 插件里重新「构建搜索索引」后再试。[/]";

        var absPath = Path.Combine(root, r.RelPath.Replace('/', Path.DirectorySeparatorChar));

        if (!File.Exists(absPath))
            return $"[red]文件不存在：[/]{Markup.Escape(absPath)}";

        var cellAddr = $"{ColToLetter(r.Col)}{r.Row}";

        try
        {
            var safePath = absPath.Replace("'", "''");
            var safeSheet = r.Sheet.Replace("'", "''");
            var ps =
                $"$xl = New-Object -ComObject Excel.Application; "
                + "$xl.Visible = $true; "
                + $"$wb = $xl.Workbooks.Open('{safePath}'); "
                + "try { "
                + $"$ws = $wb.Sheets.Item('{safeSheet}'); "
                + "$ws.Activate(); "
                + $"$ws.Range('{cellAddr}').Select(); "
                + $"$xl.ActiveWindow.ScrollRow = $ws.Range('{cellAddr}').Row "
                + "} catch {}";
            Process.Start(
                new ProcessStartInfo("pwsh", $"-NoProfile -WindowStyle Hidden -Command \"{ps}\"")
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                }
            );
            return $"[green]已打开[/] {Markup.Escape(Path.GetFileName(absPath))}  "
                + $"Sheet:[cyan]{Markup.Escape(r.Sheet)}[/]  "
                + $"[yellow]{cellAddr}[/]  "
                + $"值:[yellow]{Markup.Escape(r.Value)}[/]";
        }
        catch (Exception ex)
        {
            return $"[red]打开失败：[/]{Markup.Escape(ex.Message)}";
        }
    }

    private static string ColToLetter(int col)
    {
        var result = string.Empty;
        while (col > 0)
        {
            col--;
            result = (char)('A' + col % 26) + result;
            col /= 26;
        }
        return result;
    }

    // ── 搜索 ──────────────────────────────────────────────────────────────────

    internal static List<SearchResult> DoSearch(ExcelSearchIndex idx, string query, bool usePrefix)
    {
        if (string.IsNullOrEmpty(query))
            return [];

        List<CellHit> hits;

        if (usePrefix)
        {
            var sorted = idx.SortedKeys;
            hits = [];
            if (sorted != null)
            {
                int lo = 0,
                    hi = sorted.Length;
                while (lo < hi)
                {
                    int mid = (lo + hi) / 2;
                    if (string.Compare(sorted[mid], query, StringComparison.OrdinalIgnoreCase) < 0)
                        lo = mid + 1;
                    else
                        hi = mid;
                }
                for (int i = lo; i < sorted.Length && hits.Count < 500; i++)
                {
                    if (!sorted[i].StartsWith(query, StringComparison.OrdinalIgnoreCase))
                        break;
                    if (idx.Exact.TryGetValue(sorted[i], out var h))
                        hits.AddRange(h);
                }
            }
        }
        else
        {
            hits = idx.SearchByContains(query, StringComparison.OrdinalIgnoreCase, maxCap: 500);
        }

        var seen = new HashSet<(int, int, int)>();
        var results = new List<SearchResult>(Math.Min(hits.Count, 300));
        foreach (var h in hits)
        {
            if (!seen.Add((h.FileId, h.SheetId, h.Row)))
                continue;
            var file = h.FileId < idx.Files.Count ? idx.Files[h.FileId] : "?";
            var sheet = h.SheetId < idx.Sheets.Count ? idx.Sheets[h.SheetId] : "?";
            var val = FindValueForHit(idx, h);
            results.Add(new SearchResult(file, sheet, h.Row, h.Col, val));
            if (results.Count >= 300)
                break;
        }
        return results;
    }

    private static string FindValueForHit(ExcelSearchIndex idx, CellHit h)
    {
        int scanned = 0;
        foreach (var (k, hits) in idx.Exact)
        {
            foreach (var hit in hits)
            {
                if (hit.FileId == h.FileId && hit.Row == h.Row && hit.Col == h.Col)
                    return k;
            }
            if (++scanned > 1000)
                break;
        }
        return "";
    }

    // ── 渲染 ──────────────────────────────────────────────────────────────────

    /// <summary>
    /// 构建渲染文本（纯函数，不依赖控制台，可被测试调用）。
    /// 返回的字符串含 Spectre markup，供测试断言。
    /// </summary>
    internal static string BuildRenderText(
        string query,
        List<SearchResult> results,
        int selectedIdx,
        bool usePrefix,
        string? statusMsg,
        int pageSize = 0
    )
    {
        if (pageSize <= 0)
            pageSize = PageSize;

        var sb = new StringBuilder(2048);

        var modeTag = usePrefix ? "[dim][前缀][/]" : "[dim][包含][/]";
        sb.AppendLine($"搜索 {modeTag} > [bold]{Markup.Escape(query)}[/]▌");

        if (!string.IsNullOrEmpty(statusMsg))
            sb.AppendLine(statusMsg);
        else if (string.IsNullOrEmpty(query))
            sb.AppendLine("[dim]等待输入...[/]");
        else
            sb.AppendLine($"[dim]命中 {results.Count} 条[/]");

        sb.AppendLine();

        if (results.Count == 0)
        {
            if (!string.IsNullOrEmpty(query))
                sb.AppendLine("[dim]  无结果[/]");
            sb.AppendLine();
            sb.AppendLine("[dim]↑↓/PgUp/PgDn 翻页  Enter 打开  Tab 切换模式  Esc 退出[/]");
            return sb.ToString();
        }

        pageSize = Math.Max(5, pageSize);
        int pageStart = (selectedIdx / pageSize) * pageSize;
        int pageEnd = Math.Min(pageStart + pageSize, results.Count);
        int totalPages = (results.Count + pageSize - 1) / pageSize;
        int curPage = selectedIdx / pageSize + 1;

        for (int i = pageStart; i < pageEnd; i++)
        {
            var r = results[i];
            var isSel = i == selectedIdx;

            var parts = r.RelPath.Replace('\\', '/').Split('/');
            var shortPath = parts.Length > 2 ? string.Join("/", parts[^2], parts[^1]) : r.RelPath;

            if (isSel)
            {
                sb.AppendLine(
                    $"[green]▶[/] [bold]{Markup.Escape(shortPath)}[/]  "
                        + $"[cyan bold]{Markup.Escape(r.Sheet)}[/]  "
                        + $"[yellow bold]{Markup.Escape(r.Value)}[/]  "
                        + $"[bold]{ColToLetter(r.Col)}{r.Row}[/]"
                );
            }
            else
            {
                sb.AppendLine(
                    $"[dim]  {Markup.Escape(shortPath)}  "
                        + $"{Markup.Escape(r.Sheet)}  "
                        + $"{Markup.Escape(r.Value)}  "
                        + $"{ColToLetter(r.Col)}{r.Row}[/]"
                );
            }
        }

        if (totalPages > 1)
            sb.AppendLine(
                $"[dim]  第 {curPage}/{totalPages} 页（共 {results.Count} 条）  PgUp/PgDn 翻页[/]"
            );

        sb.AppendLine();
        sb.AppendLine("[dim]↑↓/PgUp/PgDn 翻页  Enter 打开  Tab 切换模式  Esc 退出[/]");

        return sb.ToString();
    }

    private static void Render(
        string query,
        List<SearchResult> results,
        int selectedIdx,
        bool usePrefix,
        string? statusMsg
    )
    {
        AnsiConsole.Clear();
        AnsiConsole.Markup(BuildRenderText(query, results, selectedIdx, usePrefix, statusMsg));
    }

    internal readonly record struct SearchResult(
        string RelPath,
        string Sheet,
        int Row,
        int Col,
        string Value
    );

    /// <summary>
    /// 从 startDir 向上查找含 Excels 子目录的根，返回 Excels 目录绝对路径；找不到返回 null。
    /// </summary>
    private static string? FindExcelsRoot(string startDir)
    {
        var dir = new DirectoryInfo(startDir);
        while (dir != null)
        {
            var excels = Path.Combine(dir.FullName, "Excels");
            if (Directory.Exists(excels))
                return excels;
            dir = dir.Parent;
        }
        return null;
    }
}
