using System.Diagnostics;
using NumDesTools.ExcelIndex;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// 用真实索引文件测量 SearchByContains 性能。
/// 结果写到系统临时目录下的 search_benchmark.txt
/// 运行：dotnet test --filter "SearchByContainsBenchmark" -c Release
/// </summary>
public class SearchByContainsBenchmark
{
    private static readonly string IndexPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesTools",
        "excel_index_M1Work_Public.json"
    );

    private static readonly string OutputPath = Path.Combine(
        Path.GetTempPath(),
        "NumDesTools.Tests",
        "tmp",
        "search_benchmark.txt"
    );

    [Fact]
    public void Benchmark_RealIndex_SearchByContains()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(OutputPath)!);
        var lines = new List<string>();
        void Log(string s)
        {
            lines.Add(s);
            Console.WriteLine(s);
        }

        if (!File.Exists(IndexPath))
        {
            Log($"SKIP: 索引文件不存在 {IndexPath}");
            File.WriteAllLines(OutputPath, lines);
            return;
        }

        Log($"索引路径: {IndexPath}");
        var sw = Stopwatch.StartNew();
        var idx = ExcelSearchIndex.LoadFromDisk(IndexPath);
        sw.Stop();
        Assert.NotNull(idx);
        Log($"加载耗时: {sw.ElapsedMilliseconds} ms  keys={idx.Exact.Count:N0}");

        idx.BuildSortedKeys();
        Log($"SortedKeys: {idx.SortedKeys!.Length:N0} 条");
        Log("");

        var keywords = new[]
        {
            ("item",   StringComparison.OrdinalIgnoreCase),
            ("1001",   StringComparison.Ordinal),
            ("76320",  StringComparison.Ordinal),
            ("reward", StringComparison.OrdinalIgnoreCase),
            ("z",      StringComparison.OrdinalIgnoreCase),
            ("a",      StringComparison.OrdinalIgnoreCase),
        };

        const int rounds = 5;
        Log($"{"关键词",-12} {"比较",-20} {"平均ms",8} {"命中数",8}");
        Log(new string('-', 56));

        foreach (var (kw, cmp) in keywords)
        {
            long total = 0;
            int hitCount = 0;
            for (int i = 0; i < rounds; i++)
            {
                sw.Restart();
                var hits = idx.SearchByContains(kw, cmp, maxCap: 500);
                sw.Stop();
                total += sw.ElapsedMilliseconds;
                if (i == 0)
                    hitCount = hits.Count;
            }
            Log($"{kw,-12} {cmp,-20} {total / rounds,8} {hitCount,8}");
        }

        Log("");
        Log($"结果已写到: {OutputPath}");
        File.WriteAllLines(OutputPath, lines);
    }
}
