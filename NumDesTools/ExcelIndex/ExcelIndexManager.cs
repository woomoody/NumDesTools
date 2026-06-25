using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace NumDesTools.ExcelIndex;

/// <summary>
/// 搜索索引单例管理器：后台构建、FileSystemWatcher 增量更新、O(1) 精确查询。
/// </summary>
internal sealed class ExcelIndexManager
{
    public static readonly ExcelIndexManager Instance = new();
    private ExcelIndexManager() { }

    private volatile ExcelSearchIndex? _index;
    private volatile string? _excelsRoot;
    private FileSystemWatcher? _watcher;
    private readonly ConcurrentDictionary<string, byte> _pendingRebuild = new(StringComparer.OrdinalIgnoreCase);
    private System.Threading.Timer? _debounceTimer;
    private CancellationTokenSource _cts = new();

    public bool IsReady => _index != null;
    public ExcelSearchIndex? Index => _index;
    public string? ExcelsRoot => _excelsRoot;
    /// <summary>FileSystemWatcher 检测到变化但新索引尚未建完，搜索结果可能来自旧索引。</summary>
    public bool IsOutdated => _pendingRebuild.Count > 0;

    /// <summary>按 sheet 名搜索，返回 (absPath, sheetName, row=1, col=1) 列表。</summary>
    public static List<(string file, string sheet, int row, int col)> SearchSheetNameFromIndex(
        string findValue, bool isContains,
        ExcelSearchIndex index, string excelsRoot)
    {
        var root = excelsRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var result = new List<(string, string, int, int)>();

        foreach (var (fid, sid) in index.AllSheets)
        {
            var relPath = fid < index.Files.Count ? index.Files[fid] : "";
            var sheet   = sid < index.Sheets.Count ? index.Sheets[sid] : "";
            var absPath = root + relPath.Replace('/', Path.DirectorySeparatorChar);
            var fileName = Path.GetFileNameWithoutExtension(absPath);

            bool fileMatch  = isContains ? fileName.Contains(findValue) : fileName == findValue;
            bool sheetMatch = isContains ? sheet.Contains(findValue)    : sheet == findValue;

            if (fileMatch || sheetMatch)
                result.Add((absPath, sheet, 1, 1));
        }
        return result;
    }

    // ── 启动 ────────────────────────────────────────────────────────────────

    /// <summary>插件加载时调用，传入当前工作簿路径，后台异步构建索引。</summary>
    public void StartForPath(string workbookPath)
    {
        var root = FindExcelsRoot(workbookPath);
        if (root == null) return;

        // 同一个项目无需重复启动
        if (string.Equals(root, _excelsRoot, StringComparison.OrdinalIgnoreCase)) return;

        // 切换项目：取消旧任务，重置状态
        _cts.Cancel();
        _cts = new CancellationTokenSource();
        _index = null;
        _excelsRoot = root;

        SetupWatcher(root);

        var ct = _cts.Token;
        Task.Run(() => BuildIndex(ct), ct);
    }

    // ── 搜索 ────────────────────────────────────────────────────────────────

    /// <summary>
    /// 精确搜索。返回 null 表示索引未就绪（调用方 fallback 全扫）。
    /// 返回空列表表示确实没有命中。
    /// </summary>
    public List<(string file, string sheet, int row, int col)>? TrySearch(string value, int colFilter = 0)
    {
        var idx = _index;
        if (idx == null) return null;
        if (!idx.Exact.TryGetValue(value, out var hits))
            return new List<(string, string, int, int)>();

        var root = _excelsRoot ?? "";
        var query = hits.AsEnumerable();
        if (colFilter > 0)
            query = query.Where(h => h.Col == colFilter);

        return query
            .Select(h => (
                Path.Combine(root, (idx.Files.Count > h.FileId ? idx.Files[h.FileId] : "").Replace('/', Path.DirectorySeparatorChar)),
                idx.Sheets.Count > h.SheetId ? idx.Sheets[h.SheetId] : "",
                h.Row,
                h.Col
            ))
            .ToList();
    }

    // ── 构建 ────────────────────────────────────────────────────────────────

    private void BuildIndex(CancellationToken ct)
    {
        try
        {
            var root = _excelsRoot!;
            var jsonPath = ExcelSearchIndex.GetIndexPath(root);
            var existing = ExcelSearchIndex.LoadFromDisk(jsonPath);

            // 构建期间用旧索引提供服务，不置 null，避免 fallback 全扫
            if (existing != null && _index == null)
                _index = existing;

            PluginLog.Write($"[ExcelIndex] building index for {root}  existing={existing != null}");

            var built = new ExcelIndexBuilder(root).Build(existing, ct: ct);
            if (ct.IsCancellationRequested) return;

            built.BuildSortedKeys();
            built.SaveToDisk(jsonPath);
            _index = built;
            PluginLog.Write($"[ExcelIndex] ready  keys={built.Exact.Count}  files={built.Files.Count}");
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            PluginLog.Write($"[ExcelIndex] build error: {ex.Message}");
        }
    }

    // ── FileSystemWatcher ────────────────────────────────────────────────────

    private void SetupWatcher(string root)
    {
        _watcher?.Dispose();
        if (!Directory.Exists(root)) return;

        _watcher = new FileSystemWatcher(root, "*.xlsx")
        {
            IncludeSubdirectories = true,
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
            EnableRaisingEvents = true,
        };
        _watcher.Changed += OnFileChanged;
        _watcher.Created += OnFileChanged;
        _watcher.Deleted += OnFileChanged;
        _watcher.Renamed += (_, e) => { OnFileChanged(null, new FileSystemEventArgs(WatcherChangeTypes.Deleted, Path.GetDirectoryName(e.OldFullPath)!, e.OldName!)); OnFileChanged(null, e); };
    }

    private void OnFileChanged(object? _, FileSystemEventArgs e)
    {
        _pendingRebuild[e.FullPath] = 0;
        // Debounce 5秒：Excel 保存时会连续触发多次事件
        _debounceTimer?.Change(5000, Timeout.Infinite);
        _debounceTimer ??= new System.Threading.Timer(_ => IncrementalRebuild(), null, 5000, Timeout.Infinite);
    }

    private void IncrementalRebuild()
    {
        var files = _pendingRebuild.Keys.ToArray();
        _pendingRebuild.Clear();
        if (files.Length == 0 || _excelsRoot == null) return;

        Task.Run(() =>
        {
            try
            {
                PluginLog.Write($"[ExcelIndex] incremental rebuild  changed={files.Length}");
                var root = _excelsRoot!;
                var jsonPath = ExcelSearchIndex.GetIndexPath(root);
                var existing = _index ?? ExcelSearchIndex.LoadFromDisk(jsonPath);
                var built = new ExcelIndexBuilder(root).Build(existing);
                built.SaveToDisk(jsonPath);
                _index = built;
                PluginLog.Verbose($"[ExcelIndex] incremental done");
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ExcelIndex] incremental error: {ex.Message}");
            }
        });
    }

    /// <summary>前缀搜索：返回所有以 prefix 开头的 cell value 的命中位置。</summary>
    public static List<(string file, string sheet, int row, int col)> SearchByPrefix(
        string prefix, ExcelSearchIndex index, string excelsRoot)
    {
        var sorted = index.SortedKeys;
        if (sorted == null || sorted.Length == 0)
            return new List<(string, string, int, int)>();

        var root = excelsRoot.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var result = new List<(string, string, int, int)>();

        // 二分查找第一个 >= prefix 的位置
        int lo = 0, hi = sorted.Length;
        while (lo < hi)
        {
            int mid = (lo + hi) / 2;
            result.Clear(); // just reuse variable for comparison
            if (string.Compare(sorted[mid], prefix, StringComparison.Ordinal) < 0) lo = mid + 1;
            else hi = mid;
        }
        result.Clear();

        for (int i = lo; i < sorted.Length; i++)
        {
            if (!sorted[i].StartsWith(prefix, StringComparison.Ordinal)) break;
            if (!index.Exact.TryGetValue(sorted[i], out var hits)) continue;
            foreach (var hit in hits)
            {
                var relPath = hit.FileId < index.Files.Count ? index.Files[hit.FileId] : "";
                var absPath = root + relPath.Replace('/', Path.DirectorySeparatorChar);
                var sheet   = hit.SheetId < index.Sheets.Count ? index.Sheets[hit.SheetId] : "";
                result.Add((absPath, sheet, hit.Row, hit.Col));
            }
        }
        return result;
    }

    // ── 工具 ────────────────────────────────────────────────────────────────

    private static string? FindExcelsRoot(string path)
    {
        var dir = Directory.Exists(path) ? new DirectoryInfo(path) : new DirectoryInfo(path).Parent;
        while (dir != null)
        {
            if (dir.Name.Equals("Excels", StringComparison.OrdinalIgnoreCase))
                return dir.FullName;
            dir = dir.Parent;
        }
        return null;
    }
}
