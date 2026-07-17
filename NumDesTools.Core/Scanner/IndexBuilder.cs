using NumDesTools.ExcelIndex;
using Spectre.Console;

namespace NumDesTools.Scanner;

/// <summary>
/// 命令行独立构建 Excel 搜索索引，不依赖 Excel 插件进程。
/// 用法：NumDesTools.Scanner.exe --build-index &lt;excels-root-dir&gt;
/// </summary>
internal static class IndexBuilder
{
    public static int Run(string[] args)
    {
        int idx = Array.IndexOf(args, "--build-index");
        if (idx < 0 || idx + 1 >= args.Length)
        {
            AnsiConsole.MarkupLine("[red]用法:[/] --build-index <excels-root-dir>");
            return 1;
        }

        var dir = args[idx + 1];
        if (!Directory.Exists(dir))
        {
            AnsiConsole.MarkupLine($"[red]目录不存在:[/] {Markup.Escape(dir)}");
            return 1;
        }

        var indexPath = ExcelSearchIndex.GetIndexPath(dir);
        var existing = ExcelSearchIndex.LoadFromDisk(indexPath);
        var isIncremental = existing != null;

        AnsiConsole.MarkupLine(
            isIncremental
                ? $"[dim]增量更新索引（上次构建：{existing!.BuiltAt:yyyy-MM-dd HH:mm}）[/]"
                : "[dim]全量构建索引...[/]"
        );

        ExcelSearchIndex? newIndex = null;
        AnsiConsole
            .Progress()
            .AutoClear(false)
            .HideCompleted(false)
            .Start(ctx =>
            {
                var task = ctx.AddTask("[green]扫描文件[/]", maxValue: 100);
                var builder = new ExcelIndexBuilder(dir);
                newIndex = builder.Build(
                    existing,
                    new Progress<(int done, int total)>(p =>
                    {
                        if (p.total > 0)
                            task.Value = (double)p.done / p.total * 100;
                    })
                );
                task.Value = 100;
            });

        newIndex!.BuildSortedKeys();
        newIndex.SaveToDisk(indexPath);

        AnsiConsole.MarkupLine(
            $"[green]✓ 索引已保存[/]  {newIndex.Exact.Count} 个唯一值  {newIndex.Files.Count} 个文件"
        );
        AnsiConsole.MarkupLine($"  [dim]{indexPath}[/]");
        return 0;
    }
}
