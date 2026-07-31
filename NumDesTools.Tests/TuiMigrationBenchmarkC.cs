using System.Diagnostics;
using System.Text;
using NumDesTools.XlsxEditor;
using Xunit.Abstractions;

namespace NumDesTools.Tests;

/// <summary>
/// 方案C（纯C# TUI，零序列化直接复用现有 ColumnStore）性能实测。
/// 用真实 Item.xlsx（~6.5万行×85列）测「读入 / 排序 / 整列复制」三段耗时，供 A/B/C 选型对比。
/// 只读，不修改原文件。跑 3 次取中位数，避免单次抖动误导结论。
/// </summary>
public sealed class TuiMigrationBenchmarkC(ITestOutputHelper output)
{
    private const string ItemPath = @"C:\M1Work\public\Excels\Tables\Item.xlsx";

    [Fact]
    public void ReportTimings_Load_Sort_EntireColumnCopy()
    {
        var store = ColumnStoreExcelLoader.Load(ItemPath); // warm-up, also gives dims for the runs below
        var rowCount = store.RowCount;
        var colCount = store.ColumnCount;

        var loadMs = Median(Times(3, () => ColumnStoreExcelLoader.Load(ItemPath)));

        var view = new VirtualizingSortableView(store);
        var sortMs = Median(Times(3, () => view.SortBy(1, ascending: true)));

        var copyMs = Median(Times(3, () => CopyEntireColumnsLikeMainWindow(view, store, rowCount, colCount)));

        output.WriteLine(
            $"[方案C 纯C#] rows={rowCount} cols={colCount} | "
                + $"读入(median of 3)={loadMs:F0}ms 排序={sortMs:F0}ms 整列复制(全{colCount}列)={copyMs:F0}ms"
        );
    }

    /// <summary>复刻 MainWindow.CopySelectionToClipboard 的 EntireColumn 分支：不物化 RowView，
    /// 直接用 view.GetStoreRowIndex(r) 反查 store 行号后 GetCell(storeRow, c)。</summary>
    private static void CopyEntireColumnsLikeMainWindow(
        VirtualizingSortableView view,
        ColumnStore store,
        int rowCount,
        int colCount
    )
    {
        var rows = new SortedDictionary<int, SortedDictionary<int, string?>>();
        for (var r = 0; r < rowCount; r++)
        {
            var storeRow = view.GetStoreRowIndex(r);
            var cols = new SortedDictionary<int, string?>();
            for (var c = 0; c < colCount; c++)
                cols[c] = store.GetCell(storeRow, c);
            rows[r] = cols;
        }
        var sb = new StringBuilder();
        foreach (var row in rows.Values)
        {
            sb.Append(string.Join('\t', row.Values));
            sb.Append('\n');
        }
    }

    private static IEnumerable<double> Times(int runs, Action action)
    {
        for (var i = 0; i < runs; i++)
        {
            var sw = Stopwatch.StartNew();
            action();
            sw.Stop();
            yield return sw.Elapsed.TotalMilliseconds;
        }
    }

    private static double Median(IEnumerable<double> values)
    {
        var xs = values.ToList();
        xs.Sort();
        return xs[xs.Count / 2];
    }
}
