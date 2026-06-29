using System.Collections.Concurrent;
using NumDesTools.ExcelIndex;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// TDD：验证 ExcelIndexBuilder 增量更新逻辑。
/// 使用 ScanOverride 注入假数据，完全不依赖真实 xlsx 文件。
/// </summary>
public class ExcelIndexBuilderIncrementalTests : IDisposable
{
    // ── 临时目录（模拟 Excels 根）──────────────────────────────────────────

    private readonly string _root;

    public ExcelIndexBuilderIncrementalTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"idx_test_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    // ── 辅助 ────────────────────────────────────────────────────────────────

    /// <summary>在 _root 下创建空文件，返回绝对路径。</summary>
    private string MakeFile(string relPath)
    {
        var abs = Path.Combine(_root, relPath);
        Directory.CreateDirectory(Path.GetDirectoryName(abs)!);
        File.WriteAllText(abs, "placeholder");
        return abs;
    }

    /// <summary>
    /// 构造一个已有索引，包含指定文件的命中记录和 MD5。
    /// md5 传 null 时自动计算真实文件 MD5。
    /// </summary>
    private static ExcelSearchIndex MakeExisting(
        string relPath,
        string sheetName,
        string cellValue,
        int row,
        int col,
        string md5
    )
    {
        var idx = new ExcelSearchIndex();
        idx.Files.Add(relPath);
        idx.Sheets.Add(sheetName);
        idx.AllSheets.Add((0, 0));
        idx.FileIds[relPath] = 0;
        idx.SheetIds[sheetName] = 0;
        idx.FileMd5[relPath] = md5;
        idx.Exact[cellValue] = [new CellHit(0, 0, row, col)];
        return idx;
    }

    /// <summary>
    /// 构建 builder，注入 scanOverride 和文件列表。
    /// files 传 null 时自动枚举 _root 下所有 .xlsx 文件。
    /// </summary>
    private ExcelIndexBuilder MakeBuilder(
        Action<
            string,
            ConcurrentBag<(string relPath, string sheet, string val, int row, int col, string md5)>
        > scanOverride,
        string[]? files = null
    )
    {
        var builder = new ExcelIndexBuilder(_root)
        {
            ScanOverride = scanOverride,
            FileListOverride =
                files != null
                    ? () => files
                    : () => Directory.GetFiles(_root, "*.xlsx", SearchOption.AllDirectories),
        };
        return builder;
    }

    private static string FakeMd5(string seed) =>
        BitConverter
            .ToString(
                System.Security.Cryptography.MD5.HashData(System.Text.Encoding.UTF8.GetBytes(seed))
            )
            .Replace("-", "");

    // ── 测试 1：未变化文件的命中被迁移 ─────────────────────────────────────

    [Fact]
    public void Incremental_UnchangedFile_HitsPreserved()
    {
        var absFile = MakeFile("Tables/Item.xlsx");
        var relPath = "Tables/Item.xlsx";
        var md5 = FakeMd5("same");

        // existing 里有旧命中
        var existing = MakeExisting(relPath, "Sheet1", "item_1001", 5, 2, md5);

        // ScanOverride：模拟文件未变化 → 使用相同 md5，但 scan 不应被调用
        var scanned = new List<string>();
        var builder = MakeBuilder(
            (absPath, bag) =>
            {
                scanned.Add(absPath);
                // 写入相同内容（NeedsRebuild 走 MD5 比较，若 md5 一致不会调用此函数）
                bag.Add((relPath, "Sheet1", "item_1001", 5, 2, md5));
            }
        );

        // 文件 md5 要和 existing 一致 → 写入固定内容使 ComputeMd5 匹配
        // 由于是真实文件，直接用 ScanOverride 绕过 NeedsRebuild 的 ComputeMd5
        // 改用：existing 里 md5 设为文件真实 md5
        var realMd5 = ComputeRealMd5(absFile);
        existing.FileMd5[relPath] = realMd5;

        var result = builder.Build(existing);

        // 未变化文件的命中必须在新索引里
        Assert.True(result.Exact.ContainsKey("item_1001"), "未变化文件的命中应被迁移");
        Assert.Empty(scanned); // 文件未变化，不应触发 scanOverride
    }

    // ── 测试 2：变化文件的旧命中丢弃，新命中进入 ──────────────────────────

    [Fact]
    public void Incremental_ChangedFile_HitsReplaced()
    {
        var absFile = MakeFile("Tables/Item.xlsx");
        var relPath = "Tables/Item.xlsx";

        // existing：旧 md5 与文件不一致（文件已被修改）
        var existing = MakeExisting(relPath, "Sheet1", "old_value", 5, 2, FakeMd5("stale"));

        // scanOverride 写入新命中
        var builder = MakeBuilder(
            (absPath, bag) =>
            {
                var rel = Path.GetRelativePath(_root, absPath).Replace('\\', '/');
                bag.Add((rel, "Sheet1", "new_value", 5, 2, ComputeRealMd5(absPath)));
            }
        );

        var result = builder.Build(existing);

        Assert.False(result.Exact.ContainsKey("old_value"), "旧命中应被丢弃");
        Assert.True(result.Exact.ContainsKey("new_value"), "新命中应进入索引");
    }

    // ── 测试 3：删除的文件命中不出现在新索引 ────────────────────────────

    [Fact]
    public void Incremental_DeletedFile_HitsRemoved()
    {
        var relPath = "Tables/Deleted.xlsx";
        // existing 里有旧命中，但对应文件已不存在（不在磁盘上）
        var existing = MakeExisting(relPath, "Sheet1", "ghost_value", 5, 2, FakeMd5("old"));

        // 没有创建文件，文件收集器找不到它
        var builder = MakeBuilder((_, _) => { });

        var result = builder.Build(existing);

        Assert.False(result.Exact.ContainsKey("ghost_value"), "已删除文件的命中不应出现");
    }

    // ── 测试 4：新文件命中进入新索引 ────────────────────────────────────

    [Fact]
    public void Incremental_NewFile_HitsAdded()
    {
        var absFile = MakeFile("Tables/New.xlsx");

        // existing 为空（没有此文件的记录）
        var existing = new ExcelSearchIndex();

        var builder = MakeBuilder(
            (absPath, bag) =>
            {
                var rel = Path.GetRelativePath(_root, absPath).Replace('\\', '/');
                bag.Add((rel, "Sheet1", "brand_new_value", 5, 2, ComputeRealMd5(absPath)));
            }
        );

        var result = builder.Build(existing);

        Assert.True(result.Exact.ContainsKey("brand_new_value"), "新文件的命中应进入索引");
    }

    // ── 测试 5：只有变化文件触发 scanOverride ───────────────────────────

    [Fact]
    public void Incremental_OnlyChangedFilesRescanned()
    {
        // 创建两个文件
        var absA = MakeFile("Tables/A.xlsx");
        var relA = "Tables/A.xlsx";
        var absB = MakeFile("Tables/B.xlsx");
        var relB = "Tables/B.xlsx";

        // existing：A 的 md5 与真实文件一致（未变化），B 的 md5 过时（已变化）
        var existing = new ExcelSearchIndex();
        existing.Files.AddRange([relA, relB]);
        existing.Sheets.Add("Sheet1");
        existing.AllSheets.AddRange([(0, 0), (1, 0)]);
        existing.FileIds[relA] = 0;
        existing.FileIds[relB] = 1;
        existing.SheetIds["Sheet1"] = 0;
        existing.FileMd5[relA] = ComputeRealMd5(absA); // A 未变化
        existing.FileMd5[relB] = FakeMd5("stale"); // B 已变化
        existing.Exact["val_a"] = [new CellHit(0, 0, 5, 2)];
        existing.Exact["val_b_old"] = [new CellHit(1, 0, 5, 2)];

        var scannedFiles = new List<string>();
        var builder = MakeBuilder(
            (absPath, bag) =>
            {
                scannedFiles.Add(Path.GetFileName(absPath));
                var rel = Path.GetRelativePath(_root, absPath).Replace('\\', '/');
                bag.Add((rel, "Sheet1", "val_b_new", 5, 2, ComputeRealMd5(absPath)));
            }
        );

        var result = builder.Build(existing);

        // 只有 B 被重扫
        Assert.Single(scannedFiles);
        Assert.Equal("B.xlsx", scannedFiles[0]);

        // A 的旧命中保留，B 的旧命中丢弃，B 的新命中进入
        Assert.True(result.Exact.ContainsKey("val_a"), "A 未变化，旧命中应保留");
        Assert.False(result.Exact.ContainsKey("val_b_old"), "B 已变化，旧命中应丢弃");
        Assert.True(result.Exact.ContainsKey("val_b_new"), "B 重扫后新命中应进入");
    }

    // ── 测试 6：AllSheets 未变化文件条目被正确迁移 ────────────────────────

    [Fact]
    public void Incremental_AllSheets_UnchangedFileMigrated()
    {
        var absFile = MakeFile("Tables/Item.xlsx");
        var relPath = "Tables/Item.xlsx";
        var realMd5 = ComputeRealMd5(absFile);

        var existing = MakeExisting(relPath, "c_item", "item_1001", 5, 2, realMd5);

        var builder = MakeBuilder((_, _) => { }); // 不触发扫描

        var result = builder.Build(existing);

        // AllSheets 里应该有 (Item.xlsx, c_item) 这个组合
        var hasEntry = result.AllSheets.Any(pair =>
        {
            var file = pair.FileId < result.Files.Count ? result.Files[pair.FileId] : "";
            var sheet = pair.SheetId < result.Sheets.Count ? result.Sheets[pair.SheetId] : "";
            return file == relPath && sheet == "c_item";
        });
        Assert.True(hasEntry, "AllSheets 应包含未变化文件的 (file, sheet) 条目");
    }

    // ── 工具 ────────────────────────────────────────────────────────────────

    private static string ComputeRealMd5(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return BitConverter
            .ToString(System.Security.Cryptography.MD5.HashData(fs))
            .Replace("-", "");
    }
}
