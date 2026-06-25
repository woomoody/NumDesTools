using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NumDesTools.ExcelIndex;

/// <summary>命中记录，用 int ID 代替字符串，压缩内存占用</summary>
public readonly record struct CellHit(int FileId, int SheetId, int Row, int Col);

/// <summary>搜索索引：cell值 → 命中列表，含字符串池和文件MD5快照</summary>
public class ExcelSearchIndex
{
    /// <summary>正向索引：精确值 → 命中列表</summary>
    public Dictionary<string, List<CellHit>> Exact { get; set; } = new(StringComparer.Ordinal);

    /// <summary>文件路径字符串池（相对路径，相对于 Excels 根目录）</summary>
    public List<string> Files { get; set; } = new();

    /// <summary>Sheet 名字符串池</summary>
    public List<string> Sheets { get; set; } = new();

    /// <summary>所有 (fileId, sheetId) 组合，用于 sheet 名搜索</summary>
    public List<(int FileId, int SheetId)> AllSheets { get; set; } = new();

    /// <summary>文件相对路径 → MD5，用于增量判断</summary>
    public Dictionary<string, string> FileMd5 { get; set; } = new();

    /// <summary>索引构建时间</summary>
    public DateTime BuiltAt { get; set; }

    // ── 仅运行时使用，不序列化 ────────────────────────────────────────────────

    [JsonIgnore]
    public Dictionary<string, int> FileIds { get; } = new(StringComparer.OrdinalIgnoreCase);

    [JsonIgnore]
    public Dictionary<string, int> SheetIds { get; } = new(StringComparer.Ordinal);

    /// <summary>前缀查询用有序 key 数组（RebuildLookups 或 BuildSortedKeys 后可用）</summary>
    [JsonIgnore]
    public string[]? SortedKeys { get; private set; }

    /// <summary>构建前缀查询用的有序数组（索引 ready 后调用一次即可）</summary>
    public void BuildSortedKeys()
    {
        SortedKeys = Exact.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray();
    }

    /// <summary>
    /// 前缀搜索：二分定位第一个 ≥ prefix 的 key，线性扫描直到不再以 prefix 开头。
    /// 必须先调用 BuildSortedKeys()，否则返回空。
    /// </summary>
    public List<CellHit> SearchByPrefix(string prefix)
    {
        var sorted = SortedKeys;
        if (sorted == null || sorted.Length == 0)
            return new List<CellHit>();

        // 二分查找第一个 >= prefix 的位置
        int lo = 0,
            hi = sorted.Length;
        while (lo < hi)
        {
            int mid = (lo + hi) / 2;
            if (string.Compare(sorted[mid], prefix, StringComparison.Ordinal) < 0)
                lo = mid + 1;
            else
                hi = mid;
        }

        var result = new List<CellHit>();
        for (int i = lo; i < sorted.Length; i++)
        {
            if (!sorted[i].StartsWith(prefix, StringComparison.Ordinal))
                break;
            if (Exact.TryGetValue(sorted[i], out var hits))
                result.AddRange(hits);
        }
        return result;
    }

    /// <summary>
    /// 包含搜索：遍历所有 key 做 Contains，最多返回 maxCap 条命中。
    /// 比全文件扫描快，但有线性扫描开销，用 cap 防止海量结果。
    /// </summary>
    public List<CellHit> SearchByContains(
        string keyword,
        StringComparison comparison,
        int maxCap = 500
    )
    {
        var result = new List<CellHit>();
        foreach (var (key, hits) in Exact)
        {
            if (!key.Contains(keyword, comparison))
                continue;
            result.AddRange(hits);
            if (result.Count >= maxCap)
                break;
        }
        return result;
    }

    /// <summary>构建完成后同步反向池（从磁盘加载时调用）</summary>
    public void RebuildLookups()
    {
        FileIds.Clear();
        SheetIds.Clear();
        for (int i = 0; i < Files.Count; i++)
            FileIds[Files[i]] = i;
        for (int i = 0; i < Sheets.Count; i++)
            SheetIds[Sheets[i]] = i;
    }

    // ── 序列化 / 反序列化 ────────────────────────────────────────────────────

    private static readonly JsonSerializerOptions _opts = new()
    {
        WriteIndented = false,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>
    /// 先写临时文件再替换正式文件，防止写入中断导致 JSON 损坏。
    /// </summary>
    public void SaveToDisk(string jsonPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(jsonPath)!);
        var tmpPath = jsonPath + ".tmp";
        try
        {
            using var fs = new FileStream(
                tmpPath,
                FileMode.Create,
                FileAccess.Write,
                FileShare.None,
                65536
            );
            using var gz = new GZipStream(fs, CompressionLevel.Fastest);
            JsonSerializer.Serialize(gz, this, _opts);
            // GZipStream 必须在 fs 关闭前 Flush，否则尾部不完整
            gz.Flush();
            fs.Flush();
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ExcelIndex] SaveToDisk write failed: {ex.Message}");
            try { File.Delete(tmpPath); } catch { }
            throw;
        }
        // 写成功后原子替换
        File.Move(tmpPath, jsonPath, overwrite: true);
    }

    public static ExcelSearchIndex? LoadFromDisk(string jsonPath)
    {
        if (!File.Exists(jsonPath))
            return null;
        try
        {
            using var fs = new FileStream(
                jsonPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                65536
            );
            using var gz = new GZipStream(fs, CompressionMode.Decompress);
            var idx = JsonSerializer.Deserialize<ExcelSearchIndex>(gz, _opts);
            idx?.RebuildLookups();
            return idx;
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[ExcelIndex] LoadFromDisk failed: {ex.GetType().Name}: {ex.Message}");
            // 损坏文件直接删除，下次构建会重建
            try { File.Delete(jsonPath); } catch { }
            return null;
        }
    }

    /// <summary>根据 Excels 根目录路径生成 JSON 文件名（可读且唯一）</summary>
    public static string GetIndexPath(string excelsRoot)
    {
        // 取 Excels 父路径各段拼成文件名，如 C:\M1Work\public\Excels → M1Work_public
        var parent = Path.GetDirectoryName(excelsRoot.TrimEnd(Path.DirectorySeparatorChar));
        var segments = new List<string>();
        var dir = new DirectoryInfo(parent ?? excelsRoot);
        // 取最多2级父目录名，避免文件名过长
        for (int i = 0; i < 2 && dir != null && dir.Parent != null; i++, dir = dir.Parent)
            segments.Insert(0, dir.Name);
        var name = string.Join("_", segments).Replace(' ', '_');
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesTools",
            $"excel_index_{name}.json"
        );
    }
}
