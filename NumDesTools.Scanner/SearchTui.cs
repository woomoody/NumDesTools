using NumDesTools.ExcelIndex;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// 终端交互式 Excel 索引搜索（fzf 风格）。
/// 用法：NumDesTools.Scanner --search [--index &lt;path.json.gz&gt;]
///
/// 工作流：
///   1. 加载 excel_index_*.json（自动发现 Documents\NumDesTools 下的索引文件）
///   2. 用户实时输入关键词 → 动态过滤命中项
///   3. 上下键选择结果，Enter 复制文件路径到剪贴板并打印，Esc 退出
/// </summary>
internal static class SearchTui
{
    private const int MaxDisplayRows = 20;

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
                    $"[red]未找到索引文件[/]（在 {dir} 下搜索 excel_index_*.json）"
                );
                AnsiConsole.MarkupLine("[dim]请先在 Excel 插件里执行「构建搜索索引」[/]");
                return 1;
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
                            // 多索引合并：简单追加 Exact 表
                            foreach (var (k, hits) in loaded.Exact)
                            {
                                if (!idx.Exact.TryGetValue(k, out var list))
                                {
                                    idx.Exact[k] = hits;
                                }
                                else
                                {
                                    list.AddRange(hits);
                                }
                            }
                            // 追加文件/Sheet 池（偏移 ID）
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

        var totalKeys = idx.Exact.Count;
        AnsiConsole.MarkupLine(
            $"[green]索引就绪[/]  {totalKeys:N0} 个值 · {idx.Files.Count} 个文件  "
                + $"[dim](构建于 {idx.BuiltAt:yyyy-MM-dd HH:mm})[/]"
        );
        AnsiConsole.MarkupLine(
            "[dim]输入关键词实时搜索 | ↑↓ 选择 | Enter 复制路径 | Tab 切换模式 | Esc 退出[/]"
        );
        AnsiConsole.WriteLine();

        // ── 主循环 ────────────────────────────────────────────────────────────
        var query = string.Empty;
        var selectedIdx = 0;
        var usePrefix = false; // Tab 切换前缀/包含
        List<SearchResult> results = [];

        while (true)
        {
            // 渲染
            Render(query, results, selectedIdx, usePrefix);

            // 读键
            var key = Console.ReadKey(intercept: true);

            if (key.Key == ConsoleKey.Escape)
            {
                AnsiConsole.Clear();
                AnsiConsole.MarkupLine("[dim]已退出搜索。[/]");
                return 0;
            }

            if (key.Key == ConsoleKey.Enter)
            {
                if (results.Count > 0 && selectedIdx < results.Count)
                {
                    var r = results[selectedIdx];
                    var path = r.RelPath;
                    try
                    {
                        TextCopy.ClipboardService.SetText(path);
                        AnsiConsole.Clear();
                        AnsiConsole.MarkupLine($"[green]✓ 已复制：[/]{Markup.Escape(path)}");
                        AnsiConsole.MarkupLine(
                            $"  Sheet: [cyan]{Markup.Escape(r.Sheet)}[/]  行 {r.Row}  列 {r.Col}  值: [yellow]{Markup.Escape(r.Value)}[/]"
                        );
                    }
                    catch
                    {
                        // TextCopy 不可用时退化到打印
                        AnsiConsole.Clear();
                        AnsiConsole.MarkupLine($"[green]结果：[/]{Markup.Escape(path)}");
                        AnsiConsole.MarkupLine(
                            $"  Sheet: [cyan]{Markup.Escape(r.Sheet)}[/]  行 {r.Row}  列 {r.Col}"
                        );
                    }
                }
                return 0;
            }

            if (key.Key == ConsoleKey.Tab)
            {
                usePrefix = !usePrefix;
                selectedIdx = 0;
                results = DoSearch(idx, query, usePrefix);
                continue;
            }

            if (key.Key == ConsoleKey.UpArrow)
            {
                if (selectedIdx > 0)
                    selectedIdx--;
                continue;
            }

            if (key.Key == ConsoleKey.DownArrow)
            {
                if (selectedIdx < results.Count - 1)
                    selectedIdx++;
                continue;
            }

            if (key.Key == ConsoleKey.Backspace)
            {
                if (query.Length > 0)
                    query = query[..^1];
            }
            else if (!char.IsControl(key.KeyChar))
            {
                query += key.KeyChar;
            }

            selectedIdx = 0;
            results = DoSearch(idx, query, usePrefix);
        }
    }

    // ── 搜索 ──────────────────────────────────────────────────────────────────

    private static List<SearchResult> DoSearch(ExcelSearchIndex idx, string query, bool usePrefix)
    {
        if (string.IsNullOrEmpty(query))
            return [];

        List<CellHit> hits;
        List<string> matchedValues;

        if (usePrefix)
        {
            // 前缀模式：收集所有以 query 开头的 key 及命中
            var sorted = idx.SortedKeys;
            hits = [];
            matchedValues = [];
            if (sorted != null)
            {
                // 二分找起始位
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
                    {
                        matchedValues.Add(sorted[i]);
                        hits.AddRange(h);
                    }
                }
            }
        }
        else
        {
            // 包含模式
            hits = idx.SearchByContains(query, StringComparison.OrdinalIgnoreCase, maxCap: 200);
            matchedValues = [];
        }

        // 将 CellHit 转换为 SearchResult，去重（同 file+sheet+row 只取一条）
        var seen = new HashSet<(int, int, int)>();
        var results = new List<SearchResult>(Math.Min(hits.Count, MaxDisplayRows * 2));
        foreach (var h in hits)
        {
            if (!seen.Add((h.FileId, h.SheetId, h.Row)))
                continue;
            var file = h.FileId < idx.Files.Count ? idx.Files[h.FileId] : "?";
            var sheet = h.SheetId < idx.Sheets.Count ? idx.Sheets[h.SheetId] : "?";

            // 反查 value（包含模式下 matchedValues 没填，需要从 Exact 中找）
            var val =
                usePrefix && matchedValues.Count > 0
                    ? FindValueForHit(idx, h)
                    : FindValueForHit(idx, h);

            results.Add(new SearchResult(file, sheet, h.Row, h.Col, val));
            if (results.Count >= MaxDisplayRows * 3)
                break;
        }
        return results;
    }

    private static string FindValueForHit(ExcelSearchIndex idx, CellHit h)
    {
        // 扫描 Exact 找 fileId+row+col 匹配的值（最多扫 1000 条）
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

    private static void Render(
        string query,
        List<SearchResult> results,
        int selectedIdx,
        bool usePrefix
    )
    {
        AnsiConsole.Clear();

        var modeTag = usePrefix ? "[dim][前缀][/]" : "[dim][包含][/]";
        AnsiConsole.Markup($"搜索 {modeTag} > [bold]{Markup.Escape(query)}[/]▌\n");
        AnsiConsole.MarkupLine(
            $"[dim]{(string.IsNullOrEmpty(query) ? "等待输入..." : $"命中 {results.Count} 条")}[/]"
        );
        AnsiConsole.WriteLine();

        if (results.Count == 0)
        {
            if (!string.IsNullOrEmpty(query))
                AnsiConsole.MarkupLine("[dim]  无结果[/]");
            return;
        }

        var table = new Table()
            .NoBorder()
            .HideHeaders()
            .AddColumn(new TableColumn("").Width(3))
            .AddColumn(new TableColumn("").Width(40))
            .AddColumn(new TableColumn("").Width(15))
            .AddColumn(new TableColumn(""))
            .AddColumn(new TableColumn("").Width(6));

        int show = Math.Min(results.Count, MaxDisplayRows);
        for (int i = 0; i < show; i++)
        {
            var r = results[i];
            var cursor = i == selectedIdx ? "[green]▶[/]" : " ";
            var style = i == selectedIdx ? "bold" : "dim";

            // 文件名（只显示最后两段）
            var parts = r.RelPath.Replace('\\', '/').Split('/');
            var shortPath = parts.Length > 2 ? string.Join("/", parts[^2], parts[^1]) : r.RelPath;

            table.AddRow(
                cursor,
                $"[{style}]{Markup.Escape(shortPath)}[/]",
                $"[{style}][cyan]{Markup.Escape(r.Sheet)}[/][/]",
                $"[{style}][yellow]{Markup.Escape(r.Value)}[/][/]",
                $"[{style}]r{r.Row}c{r.Col}[/]"
            );
        }

        AnsiConsole.Write(table);

        if (results.Count > MaxDisplayRows)
            AnsiConsole.MarkupLine(
                $"[dim]  ... 还有 {results.Count - MaxDisplayRows} 条结果（继续输入缩小范围）[/]"
            );

        AnsiConsole.WriteLine();
        AnsiConsole.MarkupLine("[dim]↑↓选择  Enter复制路径  Tab切换模式  Esc退出[/]");
    }

    private readonly record struct SearchResult(
        string RelPath,
        string Sheet,
        int Row,
        int Col,
        string Value
    );
}
