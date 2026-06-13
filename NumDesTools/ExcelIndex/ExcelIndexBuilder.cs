using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MiniExcelLibs;

namespace NumDesTools.ExcelIndex;

/// <summary>并行构建/增量更新搜索索引</summary>
internal class ExcelIndexBuilder
{
    private readonly string _excelsRoot;

    public ExcelIndexBuilder(string excelsRoot) => _excelsRoot = excelsRoot;

    /// <summary>
    /// 构建或增量更新索引。
    /// existing != null 时只重扫 MD5 变化的文件，其余从 existing 迁移。
    /// </summary>
    public ExcelSearchIndex Build(
        ExcelSearchIndex? existing = null,
        IProgress<(int done, int total)>? progress = null,
        CancellationToken ct = default)
    {
        var files = new SelfExcelFileCollector(_excelsRoot).GetAllExcelFilesPath();
        var total = files.Length;
        var done = 0;

        // 确定需要重扫的文件（MD5 变化或新增）
        var toRebuild = files
            .Where(f => NeedsRebuild(f, existing))
            .ToArray();

        var newIndex = new ExcelSearchIndex { BuiltAt = DateTime.UtcNow };

        // 把未变化文件的命中从 existing 迁移进 newIndex
        if (existing != null)
            MergeUnchanged(existing, files.Except(toRebuild, StringComparer.OrdinalIgnoreCase), newIndex);

        // 并行扫需要重建的文件
        var bag = new ConcurrentBag<(string relPath, string sheet, string val, int row, int col, string md5)>();
        Parallel.ForEach(
            toRebuild,
            new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount, CancellationToken = ct },
            file =>
            {
                ScanFile(file, bag);
                var cur = Interlocked.Increment(ref done);
                progress?.Report((cur, total));
            });

        // 单线程合并 bag → newIndex（避免 Dictionary 并发写）
        foreach (var (relPath, sheet, val, row, col, md5) in bag)
        {
            AddHit(newIndex, relPath, sheet, val, row, col);
            newIndex.FileMd5[relPath] = md5;
        }

        return newIndex;
    }

    // ── 私有方法 ─────────────────────────────────────────────────────────────

    private bool NeedsRebuild(string absPath, ExcelSearchIndex? existing)
    {
        if (existing == null) return true;
        var rel = ToRelative(absPath);
        if (!existing.FileMd5.TryGetValue(rel, out var oldMd5)) return true;
        try { return ComputeMd5(absPath) != oldMd5; }
        catch { return true; }
    }

    private void ScanFile(
        string absPath,
        ConcurrentBag<(string, string, string, int, int, string)> bag)
    {
        try
        {
            var relPath = ToRelative(absPath);
            var md5 = ComputeMd5(absPath);
            var sheetNames = MiniExcel.GetSheetNames(absPath);

            foreach (var sheetName in sheetNames)
            {
                if (sheetName.Contains('#') ||
                    (sheetName.Contains("Sheet") && sheetName != "Sheet1"))
                    continue;

                var rows = MiniExcel.Query(absPath, sheetName: sheetName,
                    configuration: NumDesAddIn.OnOffMiniExcelCatches);

                int rowIdx = 1;
                foreach (IDictionary<string, object> row in rows)
                {
                    if (rowIdx >= 4) // 跳过前3行标题
                    {
                        int colIdx = 1;
                        foreach (var cell in row)
                        {
                            if (colIdx >= 2) // 跳过第1列注释列
                            {
                                var val = cell.Value?.ToString();
                                if (!string.IsNullOrEmpty(val))
                                    bag.Add((relPath, sheetName, val, rowIdx, colIdx, md5));
                            }
                            colIdx++;
                        }
                    }
                    rowIdx++;
                }
            }
        }
        catch { /* 单文件失败不中断整体 */ }
    }

    private void MergeUnchanged(
        ExcelSearchIndex existing,
        IEnumerable<string> unchangedAbs,
        ExcelSearchIndex target)
    {
        var unchangedRels = unchangedAbs
            .Select(ToRelative)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        // 把 existing 中属于未变化文件的命中迁移到 target
        foreach (var (val, hits) in existing.Exact)
        {
            foreach (var hit in hits)
            {
                if (hit.FileId >= existing.Files.Count) continue;
                var rel = existing.Files[hit.FileId];
                if (!unchangedRels.Contains(rel)) continue;

                var sheet = hit.SheetId < existing.Sheets.Count
                    ? existing.Sheets[hit.SheetId] : "";
                AddHit(target, rel, sheet, val, hit.Row, hit.Col);
            }
        }

        // 迁移 MD5 快照
        foreach (var rel in unchangedRels)
            if (existing.FileMd5.TryGetValue(rel, out var md5))
                target.FileMd5[rel] = md5;
    }

    private static void AddHit(
        ExcelSearchIndex idx, string relPath, string sheet,
        string val, int row, int col)
    {
        if (!idx.FileIds.TryGetValue(relPath, out var fid))
        {
            fid = idx.Files.Count;
            idx.Files.Add(relPath);
            idx.FileIds[relPath] = fid;
        }
        if (!idx.SheetIds.TryGetValue(sheet, out var sid))
        {
            sid = idx.Sheets.Count;
            idx.Sheets.Add(sheet);
            idx.SheetIds[sheet] = sid;
        }
        var hit = new CellHit(fid, sid, row, col);
        if (!idx.Exact.TryGetValue(val, out var list))
        {
            list = new List<CellHit>(2);
            idx.Exact[val] = list;
        }
        list.Add(hit);
    }

    private string ToRelative(string absPath)
    {
        var root = _excelsRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        return absPath.StartsWith(root, StringComparison.OrdinalIgnoreCase)
            ? absPath[root.Length..].Replace('\\', '/')
            : absPath;
    }

    private static string ComputeMd5(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var md5 = System.Security.Cryptography.MD5.Create();
        return BitConverter.ToString(md5.ComputeHash(fs)).Replace("-", "");
    }
}
