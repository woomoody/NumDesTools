using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using NumDesTools.XlsxEditor;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 方案B（C#读xlsx→序列化传Rust→排序/整列复制→序列化差量传回）性能实测。
/// 复用方案C（<see cref="TuiMigrationBenchmarkC"/>）已验证过的 ColumnStore 读入，
/// 只加序列化/进程往返这一段——这正是方案B相对A/C多出来的、需要单独量化的开销。
/// </summary>
public sealed class TuiMigrationBenchmarkB(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";
    private const string BenchBExePath =
        @"C:\Pro\ExcelToolsAlbum\ExcelDna-Pro\NumDesTools\tools\xlsx-bench\target\release\bench_b.exe";
    private const int SortCol = 1; // 对齐方案A/C的 SortBy(1, ...)

    private sealed record TableDto(
        [property: JsonPropertyName("sort_col")] int SortCol,
        [property: JsonPropertyName("data")] string?[][] Data
    );

    private sealed record ResultDto(
        [property: JsonPropertyName("row_order")] int[] RowOrder,
        [property: JsonPropertyName("copy_len")] long CopyLen
    );

    [Fact]
    public async Task ReportTimings_SerializeAndRustRoundTrip()
    {
        Assert.True(File.Exists(BenchBExePath), $"先编译 xlsx-bench: cargo build --release --bin bench_b（{BenchBExePath} 不存在）");

        var store = ColumnStoreExcelLoader.Load(ItemPath);
        var rowCount = store.RowCount;
        var colCount = store.ColumnCount;

        // 序列化耗时：把 ColumnStore 的列式数据整表转成 JSON 字符串（含从 GetCell 抽取的准备成本）。
        var serializeMs = Median(
            Times(
                3,
                () =>
                {
                    var data = new string?[colCount][];
                    for (var c = 0; c < colCount; c++)
                    {
                        data[c] = new string?[rowCount];
                        for (var r = 0; r < rowCount; r++)
                            data[c][r] = store.GetCell(r, c) ?? "";
                    }
                    var json = JsonSerializer.Serialize(new TableDto(SortCol, data));
                    return json.Length;
                }
            )
        );

        // 用同一份 JSON 做一次真的往返：C#序列化→写给Rust进程→Rust反序列化+排序+整列复制+序列化回传→C#读回。
        var dataOnce = new string?[colCount][];
        for (var c = 0; c < colCount; c++)
        {
            dataOnce[c] = new string?[rowCount];
            for (var r = 0; r < rowCount; r++)
                dataOnce[c][r] = store.GetCell(r, c) ?? "";
        }
        var jsonPayload = JsonSerializer.Serialize(new TableDto(SortCol, dataOnce));

        var roundTripMs = Median(await TimesAsync(3, () => RunRoundTripOnce(jsonPayload)));

        output.WriteLine(
            $"[方案B C#+Rust] rows={rowCount} cols={colCount} | "
                + $"序列化(median of 3)={serializeMs:F0}ms 整表JSON字节数={jsonPayload.Length:N0} | "
                + $"往返(含反序列化+排序+整列复制+序列化回传+进程IO,median of 3)={roundTripMs:F0}ms"
        );
    }

    private static async Task<double> RunRoundTripOnce(string json)
    {
        var psi = new ProcessStartInfo
        {
            FileName = BenchBExePath,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        var sw = Stopwatch.StartNew();
        using var proc = Process.Start(psi)!;
        var writeTask = proc.StandardInput.WriteAsync(json).ContinueWith(_ => proc.StandardInput.Close());
        var readTask = proc.StandardOutput.ReadToEndAsync();
        await Task.WhenAll(writeTask, readTask);
        await proc.WaitForExitAsync();
        sw.Stop();

        var result = JsonSerializer.Deserialize<ResultDto>(await readTask);
        Assert.NotNull(result);
        return sw.Elapsed.TotalMilliseconds;
    }

    private static IEnumerable<double> Times(int runs, Func<int> action)
    {
        for (var i = 0; i < runs; i++)
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();
            yield return sw.Elapsed.TotalMilliseconds;
        }
    }

    private static async Task<List<double>> TimesAsync(int runs, Func<Task<double>> action)
    {
        var times = new List<double>(runs);
        for (var i = 0; i < runs; i++)
            times.Add(await action());
        return times;
    }

    private static double Median(IEnumerable<double> values)
    {
        var xs = values.ToList();
        xs.Sort();
        return xs[xs.Count / 2];
    }
}
