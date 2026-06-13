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
                PluginLog.Write($"[ExcelIndex] incremental done");
            }
            catch (Exception ex)
            {
                PluginLog.Write($"[ExcelIndex] incremental error: {ex.Message}");
            }
        });
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
