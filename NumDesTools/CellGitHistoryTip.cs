using System.Collections.Concurrent;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using LibGit2Sharp;
using NumDesTools.UI;
using Timer = System.Windows.Forms.Timer;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// "谁的锅"气泡的一条历史变更记录（供 FillRichText 着色渲染）。
/// sha 可为空（lua 反推路径无 sha）。
/// NewBranch/OldBranch：diff 两端所在分支（init-only，不改主构造参数以兼容 lua 路径调用）。
/// </summary>
internal sealed record CellHistoryEntry(
    string Sha,
    string Date,
    string Author,
    string Msg,
    string OldVal,
    string NewVal
)
{
    public string NewBranch { get; init; } = "";
    public string OldBranch { get; init; } = "";
};

/// <summary>
/// "谁的锅"气泡薄适配类：持有 WPF 气泡窗口 + Excel 滚动检测 timer。
/// 对外 API 不变（Instance/ShowBubble/ClearBubble/DisposeInstance），
/// 实际渲染交给 CellHistoryBubbleWindow（WPF + WPF-UI，零 Excel 依赖，可独立编译）。
/// </summary>
public sealed class CellGitHistoryTip
{
    private readonly CellHistoryBubbleWindow _window;
    private readonly Timer _scrollTimer;
    private int _lastScrollRow;
    private int _lastScrollCol;
    private bool _hasAnchor;

    private static CellGitHistoryTip? _instance;
    public static CellGitHistoryTip Instance => _instance ??= new CellGitHistoryTip();

    private CellGitHistoryTip()
    {
        _window = new CellHistoryBubbleWindow();
        _scrollTimer = new Timer { Interval = 150 };
        _scrollTimer.Tick += OnScrollCheck;
    }

    /// <summary>显示带历史记录的富文本气泡（着色渲染）。锁定时跳过，保持钉住的内容。</summary>
    internal void ShowBubble(List<CellHistoryEntry> results)
    {
        CellHistoryBubbleWindow.EnsureWpfInitialized();
        _scrollTimer.Stop();
        _window.SetEntries(results);
        if (!_window.IsVisible)
            _window.Show();
        if (!_hasAnchor)
        {
            _window.PlaceAtCursor();
            _hasAnchor = true;
        }
        try
        {
            var win = AppServices.App.ActiveWindow;
            _lastScrollRow = win.ScrollRow;
            _lastScrollCol = win.ScrollColumn;
        }
        catch { }
        _scrollTimer.Start();
    }

    /// <summary>显示纯文本提示气泡（如"搜索中"）。锁定时跳过。</summary>
    internal void ShowBubble(string text)
    {
        CellHistoryBubbleWindow.EnsureWpfInitialized();
        _scrollTimer.Stop();
        _window.SetMessage(text);
        if (!_window.IsVisible)
            _window.Show();
        if (!_hasAnchor)
        {
            _window.PlaceAtCursor();
            _hasAnchor = true;
        }
        _scrollTimer.Start();
    }

    internal void ResetAnchor()
    {
        _hasAnchor = false;
    }

    public void ClearBubble()
    {
        _scrollTimer.Stop();
        _window.Hide();
        _hasAnchor = false;
    }

    /// <summary>强制关闭气泡（无视锁定，停 timer + 隐藏并清空）。不受 IsLocked 拦截。</summary>
    internal void ForceCloseBubble()
    {
        _scrollTimer.Stop();
        _window.ForceClose();
        _hasAnchor = false;
    }

    /// <summary>气泡窗口是否激活（控制器 deactivate 据此跳过 Clear，保留气泡供选文本复制）。</summary>
    public bool IsBubbleActive => _window.IsActiveBubble;

    private void OnScrollCheck(object? sender, EventArgs e)
    {
        // 历史气泡由“选中新单元格”或显式关闭控制，不因滚动/焦点切换自动消失。
        // 保留 timer 仅兼容既有生命周期；不再主动 ClearBubble。
    }

    public static void DisposeInstance()
    {
        if (_instance is null)
            return;
        _instance._scrollTimer.Dispose();
        _instance._window.Close();
        _instance = null;
    }
}

// ── 后台查询服务 ─────────────────────────────────────────────────────────────

internal static class CellGitHistoryService
{
    private static CancellationTokenSource? _cts;

    // LRU 缓存：key = "absFile|sheet|rowKey|colName" → 历史记录列表（着色渲染用）
    private static readonly Dictionary<string, List<CellHistoryEntry>> _cache = new(
        StringComparer.Ordinal
    );
    private static readonly object _cacheLock = new();
    private const int CacheCapacity = 100;

    // 文件级 commit 列表缓存：key = absFilePath → (list, fileLastWriteStamp)
    private static readonly Dictionary<
        string,
        (List<(string sha, string date, string author, string msg)> commits, long stamp)
    > _commitListCache = new(StringComparer.OrdinalIgnoreCase);

    // 单格值缓存：sha8|relPath|sheetName|rowKey|colName → value（条目小，内存可忽略）
    private static readonly ConcurrentDictionary<string, string> _cellValCache = new(
        StringComparer.Ordinal
    );

    public static void Query(
        string absFilePath,
        string gitRoot,
        string sheetName,
        string rowKey,
        string colName,
        Action<List<CellHistoryEntry>> onResult
    )
    {
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;

        var cacheKey = $"{absFilePath}|{sheetName}|{rowKey}|{colName}";
        lock (_cacheLock)
        {
            if (_cache.TryGetValue(cacheKey, out var cached))
            {
                onResult(cached);
                return;
            }
        }

        _ = Task.Run(
            async () =>
            {
                try
                {
                    await Task.Delay(400, ct);
                    if (ct.IsCancellationRequested)
                        return;

                    // 先立刻显示"搜索中"气泡，让用户知道已在查询
                    ExcelAsyncUtil.QueueAsMacro(() =>
                        CellGitHistoryTip.Instance.ShowBubble("🔍 搜索提交历史中…")
                    );

                    // 流式：每找到一条变更就立刻更新气泡，不等全部扫完
                    QueryHistoryStreaming(
                        absFilePath,
                        gitRoot,
                        sheetName,
                        rowKey,
                        colName,
                        ct,
                        partialResults =>
                        {
                            if (ct.IsCancellationRequested)
                                return;
                            onResult(partialResults); // 每找到一条就刷新气泡
                        },
                        finalResults =>
                        {
                            if (!ct.IsCancellationRequested && finalResults != null)
                                PutCache(cacheKey, finalResults); // 全部找完后缓存最终结果
                        }
                    );
                }
                catch (OperationCanceledException) { }
                catch { }
            },
            ct
        );
    }

    public static void CancelPending() => _cts?.Cancel();

    private static void PutCache(string key, List<CellHistoryEntry> value)
    {
        lock (_cacheLock)
        {
            if (_cache.ContainsKey(key))
                return;
            if (_cache.Count >= CacheCapacity)
            {
                var old = _cache.Keys.First();
                _cache.Remove(old);
            }
            _cache[key] = value;
        }
    }

    /// <summary>
    /// 混合流式查询：阶段1 顺序流式处理前 N 个 commit，立即出结果；
    /// 阶段2 并行处理剩余 commit，merge 后补充 diff。兼顾首条速度与全量覆盖。
    /// </summary>
    private static void QueryHistoryStreaming(
        string absFilePath,
        string gitRoot,
        string sheetName,
        string rowKey,
        string colName,
        CancellationToken ct,
        System.Action<List<CellHistoryEntry>> onPartial,
        System.Action<List<CellHistoryEntry>?> onFinal
    )
    {
        var relativePath = Path.GetRelativePath(gitRoot, absFilePath).Replace('\\', '/');
        var commits = GetRecentCommits(absFilePath, gitRoot, relativePath);
        PluginLog.Write($"[谁的锅] commits={commits.Count} for {relativePath}");
        if (commits.Count == 0)
        {
            onFinal(null);
            return;
        }

        var tmpDir = Path.Combine(Path.GetTempPath(), "NumDesCellHistory");
        Directory.CreateDirectory(tmpDir);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        int parsedCount = 0,
            skippedCount = 0,
            totalCommits = 0;

        const int MaxChanges = 50;
        const int MaxCommits = 500;
        // 前 3 个 commit 用 xlsx 兜住"未导表的最近改动"：lua 是导表快照，滞后于 xlsx 的最近改动，
        // 只扫最近几个提交就能覆盖到"lua 还没导"的部分；更早的历史交给 lua 反推（10s 覆盖全量）。
        const int StreamingPhaseCount = 3;
        const int ParallelChunks = 6; // XmlReader 流式取格内存极小，6 线程安全

        var takeCount = Math.Min(commits.Count, MaxCommits);
        var limitedCommits = commits.GetRange(0, takeCount);

        // 共享的 accumulated 列表，两阶段共用（含 sha 用于头部显示短 sha8）
        var accumulated = new List<CellHistoryEntry>();
        bool usedLua = false; // 阶段2 是否走了 lua 反推快路径

        // ── 阶段1：顺序流式（前 StreamingPhaseCount 个 commit）─────────────
        // 实时出结果，用户立刻看到
        int streamingEnd = Math.Min(StreamingPhaseCount, takeCount);
        string? prevVal = null;
        (string sha, string date, string author, string msg)? prevMeta = null;
        bool hadNonNull = false;

        // 供阶段2用的边界值（阶段1最后的 prevVal / prevMeta）
        string? phase1LastVal = null;
        (string sha, string date, string author, string msg)? phase1LastMeta = null;

        using (var streamRepo = new Repository(gitRoot))
        {
            string? prevBlobOid = null;

            for (int i = 0; i < streamingEnd; i++)
            {
                if (ct.IsCancellationRequested)
                {
                    onFinal(null);
                    return;
                }
                if (accumulated.Count >= MaxChanges)
                    break;

                var (sha, date, author, msg) = limitedCommits[i];

                // blob OID 预过滤
                var commit = streamRepo.Lookup<Commit>(sha);
                var blobEntry = commit?.Tree[relativePath];
                var blobOid = blobEntry?.Target.Sha;
                if (blobOid != null && blobOid == prevBlobOid)
                {
                    PluginLog.Verbose($"[谁的锅] commit {sha[..8]} blob unchanged, skip");
                    continue;
                }
                prevBlobOid = blobOid;

                // 提取值（MiniExcel 流式取格，内存极低）
                string? val = GetCellValueAtCommit(
                    streamRepo,
                    sha,
                    relativePath,
                    sheetName,
                    rowKey,
                    colName,
                    tmpDir
                );

                PluginLog.Verbose($"[谁的锅] phase1 commit {sha[..8]} val={val ?? "null"}");

                if (val == null)
                {
                    if (hadNonNull)
                        break;
                    continue;
                }

                hadNonNull = true;

                if (prevVal != null && prevMeta.HasValue && val != prevVal)
                {
                    // diff：较新=prevMeta.Value.sha，较旧=当前循环变量 sha
                    accumulated.Add(
                        new CellHistoryEntry(
                            prevMeta.Value.sha,
                            prevMeta.Value.date,
                            prevMeta.Value.author,
                            prevMeta.Value.msg,
                            val,
                            prevVal
                        )
                        {
                            NewBranch = GetBranch(gitRoot, prevMeta.Value.sha),
                            OldBranch = GetBranch(gitRoot, sha),
                        }
                    );
                    onPartial(accumulated);
                }

                prevVal = val;
                prevMeta = (sha, date, author, msg);
            }

            phase1LastVal = prevVal;
            phase1LastMeta = prevMeta;

            totalCommits = streamingEnd;
            sw.Stop();
            PluginLog.Write(
                $"[谁的锅] Phase1 done: {streamingEnd} commits, {accumulated.Count} changes, {sw.ElapsedMilliseconds}ms"
            );
            sw.Start();
        }

        // ── 阶段2：处理剩余 commit ───────────────────────────────────────
        if (
            streamingEnd < takeCount
            && accumulated.Count < MaxChanges
            && !ct.IsCancellationRequested
        )
        {
            // ── 阶段2a：lua 反推快路径 ─────────────────────────────────────
            // xlsx 是二进制，剩余几百个提交逐个解包扫描要几分钟；lua.txt 是导出产物（纯文本），
            // git pickaxe 原生行级追踪，~10s 覆盖全部历史。阶段1（xlsx）已兜住"未导表的最近改动"，
            // 这里只补更早的历史；按（旧值,新值）对去重，避免与阶段1已找到的变更重复计数。
            var luaEvents = CellGitHistoryLuaIndex.TryQueryHistory(
                absFilePath,
                sheetName,
                rowKey,
                colName,
                out var luaReason
            );
            PluginLog.Write($"[谁的锅] lua 反推: {luaReason}");

            if (luaEvents != null)
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var e in accumulated)
                    seen.Add(EventKey(e.Sha, e.Date, e.Author, e.Msg, e.OldVal, e.NewVal));

                int added = 0;
                foreach (var e in luaEvents)
                {
                    if (accumulated.Count >= MaxChanges || ct.IsCancellationRequested)
                        break;
                    var key = EventKey(e.Sha, e.Date, e.Author, e.Msg, e.OldVal, e.NewVal);
                    if (!seen.Add(key))
                        continue;
                    // lua 反推路径无 sha（pickaxe 行级追踪，sha 在索引内部未暴露），留空显示
                    accumulated.Add(
                        new CellHistoryEntry(
                            e.Sha,
                            e.Date,
                            e.Author,
                            e.Msg,
                            e.OldVal ?? "（空）",
                            e.NewVal ?? "（行被删除）"
                        )
                    );
                    onPartial(accumulated);
                    added++;
                }
                sw.Stop();
                PluginLog.Write(
                    $"[谁的锅] Phase2 lua 反推完成: 命中 {luaEvents.Count} 条，新增 {added} 条更早变更，{sw.ElapsedMilliseconds}ms"
                );
                sw.Start();
                usedLua = true;
            }
            else
            {
                // ── 阶段2b 回退：并行 xlsx 扫描（原有逻辑）──────────────────
                int remaining = takeCount - streamingEnd;
                int chunkSize = Math.Max(1, (remaining + ParallelChunks - 1) / ParallelChunks);
                int chunkCount = (remaining + chunkSize - 1) / chunkSize;
                var chunkResults = new (
                    string? val,
                    string sha,
                    string date,
                    string author,
                    string msg
                )[chunkCount][];

                Parallel.For(
                    0,
                    chunkCount,
                    new ParallelOptions
                    {
                        MaxDegreeOfParallelism = ParallelChunks,
                        CancellationToken = ct,
                    },
                    chunkIdx =>
                    {
                        int start = streamingEnd + chunkIdx * chunkSize;
                        int end = Math.Min(start + chunkSize, takeCount);

                        using var threadRepo = new Repository(gitRoot);
                        string? prevBlobOid = null;
                        var local = new (
                            string? val,
                            string sha,
                            string date,
                            string author,
                            string msg
                        )[end - start];

                        for (int i = start; i < end; i++)
                        {
                            if (ct.IsCancellationRequested)
                                return;

                            var (sha, date, author, msg) = limitedCommits[i];

                            var commit = threadRepo.Lookup<Commit>(sha);
                            var blobEntry = commit?.Tree[relativePath];
                            var blobOid = blobEntry?.Target.Sha;
                            if (blobOid != null && blobOid == prevBlobOid)
                            {
                                PluginLog.Verbose(
                                    $"[谁的锅] commit {sha[..8]} blob unchanged, skip"
                                );
                                continue;
                            }
                            prevBlobOid = blobOid;

                            string? val = GetCellValueAtCommit(
                                threadRepo,
                                sha,
                                relativePath,
                                sheetName,
                                rowKey,
                                colName,
                                tmpDir
                            );

                            local[i - start] = (val, sha, date, author, msg);
                        }
                        chunkResults[chunkIdx] = local;
                    }
                );

                // 合并阶段2结果（保持 commit 顺序）
                // 先用阶段1的边界值作为 prevVal/prevMeta
                prevVal = phase1LastVal;
                prevMeta = phase1LastMeta;

                for (int chunkIdx = 0; chunkIdx < chunkResults.Length; chunkIdx++)
                {
                    if (ct.IsCancellationRequested)
                    {
                        onFinal(null);
                        return;
                    }
                    var chunk = chunkResults[chunkIdx];
                    if (chunk == null)
                        continue;

                    foreach (var (val, sha, date, author, msg) in chunk)
                    {
                        if (accumulated.Count >= MaxChanges)
                            break;

                        if (val == null)
                        {
                            if (hadNonNull)
                                goto Phase2Done; // 创建边界，跳出双层循环
                            continue;
                        }

                        hadNonNull = true;

                        if (prevVal != null && prevMeta.HasValue && val != prevVal)
                        {
                            // diff：较新=prevMeta.Value.sha，较旧=当前循环变量 sha
                            accumulated.Add(
                                new CellHistoryEntry(
                                    prevMeta.Value.sha,
                                    prevMeta.Value.date,
                                    prevMeta.Value.author,
                                    prevMeta.Value.msg,
                                    val,
                                    prevVal
                                )
                                {
                                    NewBranch = GetBranch(gitRoot, prevMeta.Value.sha),
                                    OldBranch = GetBranch(gitRoot, sha),
                                }
                            );
                            onPartial(accumulated);
                        }

                        prevVal = val;
                        prevMeta = (sha, date, author, msg);
                    }
                }
                Phase2Done:
                ;
                sw.Stop();
                long phase2Ms = sw.ElapsedMilliseconds;
                PluginLog.Write(
                    $"[谁的锅] Phase2 done: {takeCount - streamingEnd} commits ({parsedCount} parsed, {skippedCount} blob-skipped), {phase2Ms}ms"
                );
                sw.Start();
            }
        }

        // 补"从无到有"创建条目：prevMeta 是该行首次出现（或窗口内最早可查）的 commit
        // lua 反推路径自带创建事件（首条 diff 即行新增），避免重复
        if (!usedLua && accumulated.Count > 0 && prevMeta.HasValue && prevVal != null)
        {
            accumulated.Add(
                new CellHistoryEntry(
                    prevMeta.Value.sha,
                    prevMeta.Value.date,
                    prevMeta.Value.author,
                    prevMeta.Value.msg + "（行首次出现）",
                    "（空）",
                    prevVal
                )
            );
            onPartial(accumulated);
        }

        // 若找到值但无任何变更，展示最老一条
        if (accumulated.Count == 0 && prevMeta.HasValue && prevVal != null)
        {
            accumulated.Add(
                new CellHistoryEntry(
                    prevMeta.Value.sha,
                    prevMeta.Value.date,
                    prevMeta.Value.author,
                    prevMeta.Value.msg + "（最早可查，值未改变）",
                    prevVal,
                    prevVal
                )
            );
            onPartial(accumulated);
        }

        sw.Stop();
        PluginLog.Write(
            $"[谁的锅] TOTAL: {takeCount} commits scanned, {accumulated.Count} changes found, {sw.ElapsedMilliseconds}ms"
        );
        var finalResults = accumulated.Count > 0 ? accumulated : null;
        onFinal(finalResults);
    }

    /// <summary>
    /// 归一化列值用于跨路径去重（xlsx 读出的值 vs lua 反推的值）：
    /// 空/「（空）」→ 空串；简单字符串去掉首尾引号。
    /// </summary>
    private static string NormVal(string? v)
    {
        if (string.IsNullOrWhiteSpace(v))
            return "";
        v = v.Trim();
        if (v == "（空）")
            return "";
        if (v.Length >= 2 && v[0] == '"' && v[^1] == '"')
            v = v[1..^1];
        return v;
    }

    private static string EventKey(
        string sha,
        string date,
        string author,
        string msg,
        string? oldVal,
        string? newVal
    ) =>
        !string.IsNullOrEmpty(sha)
            ? sha
            : string.Join('\u0001', date, author, msg, NormVal(oldVal), NormVal(newVal));

    private static List<(string sha, string date, string author, string msg)> GetRecentCommits(
        string absFilePath,
        string gitRoot,
        string relativePath
    )
    {
        var stamp = File.GetLastWriteTimeUtc(absFilePath).Ticks;
        if (_commitListCache.TryGetValue(absFilePath, out var cached) && cached.stamp == stamp)
            return cached.commits;

        try
        {
            // --all 搜索所有分支；-n 500 配合 MaxCommits=500 覆盖全部历史
            var args = $"log --all -n 500 --format=\"%H|%ai|%an|%s\" -- \"{relativePath}\"";
            var output = RunGit(gitRoot, args);

            var result = new List<(string, string, string, string)>();
            foreach (var line in output.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = line.Trim('"').Split('|', 4);
                if (parts.Length < 4)
                    continue;
                var sha = parts[0].Trim();
                if (sha.Length < 8)
                    continue;
                var dt = parts[1].Trim();
                var date = dt.Length >= 16 ? dt[..16] : dt; // git %ai = "2026-08-06 11:05:50 +0800"，[..16] 留到分钟
                result.Add((sha, date, parts[2].Trim(), parts[3].Trim()));
            }

            _commitListCache[absFilePath] = (result, stamp);
            return result;
        }
        catch
        {
            return [];
        }
    }

    /// <summary>
    /// <summary>
    /// 从目标 commit 的 xlsx 中提取单格值（裸 OOXML 流式）。
    /// XmlReader 每行只看 key 列/目标列两个 cell 的开始标签，跳过其余上百万 cell：
    /// 5.5 万行表约 300ms/commit（MiniExcel 逐行建行对象约 5s；EPPlus 全表 DOM 内存太大）。
    /// </summary>
    private static string? GetCellValueAtCommit(
        Repository repo,
        string sha,
        string relativePath,
        string sheetName,
        string rowKey,
        string colName,
        string tmpDir
    )
    {
        var cacheKey = $"{sha[..8]}|{relativePath}|{sheetName}|{rowKey}|{colName}";
        if (_cellValCache.TryGetValue(cacheKey, out var cached))
            return cached;

        try
        {
            // 提取 blob 到临时文件（已存在则复用）
            var tmpFile = Path.Combine(tmpDir, $"{sha[..8]}_{Path.GetFileName(relativePath)}");
            if (!File.Exists(tmpFile))
            {
                var commit = repo.Lookup<Commit>(sha);
                if (commit == null)
                    return null;
                var entry = commit[relativePath];
                if (entry == null)
                    return null;
                var blob = (Blob)entry.Target;
                using var src = blob.GetContentStream();
                using var dst = new FileStream(tmpFile, FileMode.Create, FileAccess.Write);
                src.CopyTo(dst);
            }

            // 裸 OOXML 流式：row 2 定列号，row 3+ 只读 key/目标两列，找到即 break
            string? result = null;
            using (var za = System.IO.Compression.ZipFile.OpenRead(tmpFile))
            {
                var sharedStrings = LoadSharedStrings(za);
                var sheetEntry = za.GetEntry(ResolveSheetPath(za, sheetName) ?? "");
                if (sheetEntry == null)
                    return null;

                int keyCol = -1,
                    targetCol = -1;
                bool headerDone = false;
                int curRow = 0;
                string? curKey = null;

                using var sh = sheetEntry.Open();
                using var xr = System.Xml.XmlReader.Create(sh);
                while (xr.Read())
                {
                    if (xr.NodeType != System.Xml.XmlNodeType.Element)
                        continue;
                    if (xr.LocalName == "row")
                    {
                        curRow = int.TryParse(xr.GetAttribute("r"), out var rr) ? rr : curRow + 1;
                        curKey = null;
                        continue;
                    }
                    if (xr.LocalName != "c")
                        continue;

                    var col = ColIndexOf(xr.GetAttribute("r"));

                    if (curRow == 2)
                    {
                        // row 2 = 列名行：找 key 列（第一个非 #）和目标列
                        var h = ReadCellValue(xr, sharedStrings) ?? "";
                        if (keyCol < 0 && !string.IsNullOrEmpty(h) && !h.StartsWith('#'))
                            keyCol = col;
                        if (h == colName)
                            targetCol = col;
                        continue;
                    }
                    if (!headerDone && curRow > 2)
                    {
                        headerDone = true;
                        if (keyCol < 0)
                            break;
                        if (targetCol < 0)
                        {
                            result = "（列当时不存在）";
                            break;
                        }
                    }
                    if (curRow < 3)
                        continue;
                    // 不匹配的 cell 直接放过：子节点 <v>/<t> 会被 LocalName 过滤自然跳过。
                    // 不能用 xr.Skip()——Skip 后停在下一兄弟节点，循环顶部 Read() 会再吞一个 cell。
                    if (col == keyCol || col == targetCol)
                    {
                        var v = ReadCellValue(xr, sharedStrings);
                        if (col == keyCol)
                            curKey = v;
                        else if (curKey == rowKey)
                        {
                            result = string.IsNullOrEmpty(v) ? "（空）" : v;
                            break;
                        }
                    }
                }
            }

            if (result != null)
            {
                if (_cellValCache.Count >= 2000)
                    _cellValCache.Clear();
                _cellValCache[cacheKey] = result;
            }
            return result;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>单元格引用 "AG55883" → 列号 33。</summary>
    private static int ColIndexOf(string? cellRef)
    {
        int c = 0;
        if (cellRef == null)
            return 0;
        foreach (var ch in cellRef)
        {
            if (char.IsLetter(ch))
                c = c * 26 + (char.ToUpper(ch) - 'A' + 1);
            else
                break;
        }
        return c;
    }

    /// <summary>读取当前 &lt;c&gt; 节点的值（处理共享字符串/内联字符串/数字）。</summary>
    private static string? ReadCellValue(System.Xml.XmlReader xr, List<string> sharedStrings)
    {
        var t = xr.GetAttribute("t");
        if (xr.IsEmptyElement)
            return null;
        string? v = null;
        using var sub = xr.ReadSubtree();
        while (sub.Read())
        {
            if (
                sub.NodeType == System.Xml.XmlNodeType.Element
                && (sub.LocalName == "v" || sub.LocalName == "t")
            )
                v = sub.ReadElementContentAsString();
        }
        if (
            t == "s"
            && v != null
            && int.TryParse(v, out var si)
            && si >= 0
            && si < sharedStrings.Count
        )
            return sharedStrings[si];
        return v;
    }

    /// <summary>sharedStrings.xml → 字符串表（无此文件返回空表）。</summary>
    private static List<string> LoadSharedStrings(System.IO.Compression.ZipArchive za)
    {
        var ss = new List<string>();
        var entry = za.GetEntry("xl/sharedStrings.xml");
        if (entry == null)
            return ss;
        using var s = entry.Open();
        using var xr = System.Xml.XmlReader.Create(s);
        while (xr.Read())
        {
            if (xr.NodeType != System.Xml.XmlNodeType.Element || xr.LocalName != "si")
                continue;
            // 整个 <si> 用 subtree 读：ReadElementContentAsString 会吞掉 </si> 事件，不能靠 EndElement 收尾
            var sb = new StringBuilder();
            using var sub = xr.ReadSubtree();
            while (sub.Read())
            {
                if (sub.NodeType == System.Xml.XmlNodeType.Element && sub.LocalName == "t")
                    sb.Append(sub.ReadElementContentAsString());
            }
            ss.Add(sb.ToString());
        }
        return ss;
    }

    /// <summary>sheet 名 → zip 内 worksheet 路径（workbook.xml + rels 解析）。</summary>
    private static string? ResolveSheetPath(System.IO.Compression.ZipArchive za, string sheetName)
    {
        string? rid = null;
        var wbEntry = za.GetEntry("xl/workbook.xml");
        if (wbEntry == null)
            return null;
        using (var wb = wbEntry.Open())
        using (var xr = System.Xml.XmlReader.Create(wb))
        {
            while (xr.Read())
            {
                if (
                    xr.NodeType == System.Xml.XmlNodeType.Element
                    && xr.LocalName == "sheet"
                    && string.Equals(
                        xr.GetAttribute("name"),
                        sheetName,
                        StringComparison.OrdinalIgnoreCase
                    )
                )
                {
                    rid = xr.GetAttribute(
                        "id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                    );
                    break;
                }
            }
        }
        if (rid == null)
            return null;
        var relsEntry = za.GetEntry("xl/_rels/workbook.xml.rels");
        if (relsEntry == null)
            return null;
        using var rels = relsEntry.Open();
        using var xr2 = System.Xml.XmlReader.Create(rels);
        while (xr2.Read())
        {
            if (
                xr2.NodeType == System.Xml.XmlNodeType.Element
                && xr2.LocalName == "Relationship"
                && xr2.GetAttribute("Id") == rid
            )
            {
                var target = xr2.GetAttribute("Target");
                if (string.IsNullOrEmpty(target))
                    return null;
                return target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
            }
        }
        return null;
    }

    private static string RunGit(string gitRoot, string arguments)
    {
        using var proc = new System.Diagnostics.Process
        {
            StartInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = FindGitExe(),
                Arguments = arguments,
                WorkingDirectory = gitRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
            },
        };
        proc.Start();
        var stdout = proc.StandardOutput.ReadToEnd();
        proc.StandardError.ReadToEnd();
        proc.WaitForExit(15_000);
        return stdout;
    }

    // sha → 所在分支（首个非 HEAD 的 refname:short）。失败/无返回空串。
    private static readonly Dictionary<string, string> _branchCache = new(StringComparer.Ordinal);

    /// <summary>
    /// 查询某 commit 所在分支（取首个不含 HEAD 的短引用名）。
    /// 只在产生 diff 条目时调用（次数少），带进程内缓存。
    /// </summary>
    private static string GetBranch(string gitRoot, string sha)
    {
        if (string.IsNullOrEmpty(sha))
            return "";
        if (_branchCache.TryGetValue(sha, out var cached))
            return cached;
        var branch = "";
        try
        {
            var args = $"-C \"{gitRoot}\" branch -a --contains {sha} --format=%(refname:short)";
            var output = RunGit(gitRoot, args);
            foreach (var line in output.Split('\n'))
            {
                var t = line.Trim();
                if (t.Length == 0 || t.Contains("HEAD"))
                    continue;
                branch = t;
                break;
            }
        }
        catch { }
        if (_branchCache.Count >= 500)
            _branchCache.Clear();
        _branchCache[sha] = branch;
        return branch;
    }

    private static string? _gitExe;

    private static string FindGitExe()
    {
        if (_gitExe != null)
            return _gitExe;
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(';'))
        {
            try
            {
                var p = Path.Combine(dir.Trim(), "git.exe");
                if (File.Exists(p))
                    return _gitExe = p;
            }
            catch { }
        }
        var candidates = new[]
        {
            @"C:\Program Files\Git\bin\git.exe",
            @"C:\Program Files (x86)\Git\bin\git.exe",
        };
        foreach (var c in candidates)
            if (File.Exists(c))
                return _gitExe = c;
        return _gitExe = "git";
    }
}

// ── 事件控制器 ─────────────────────────────────────────────────────────────

internal static class CellGitHistoryController
{
    private static Microsoft.Office.Interop.Excel.Application? _app;
    public static bool IsActive { get; private set; }

    public static void Enable(Microsoft.Office.Interop.Excel.Application app)
    {
        if (IsActive)
            return;
        IsActive = true;
        _app = app;
        app.SheetSelectionChange += OnSelectionChange;
        app.WindowDeactivate += OnWindowDeactivate;
        app.WorkbookDeactivate += OnWorkbookDeactivate;
        app.WorkbookBeforeClose += OnWorkbookBeforeClose;
    }

    public static void Disable()
    {
        if (!IsActive || _app == null)
            return;
        IsActive = false;
        CellGitHistoryService.CancelPending();
        _app.SheetSelectionChange -= OnSelectionChange;
        _app.WindowDeactivate -= OnWindowDeactivate;
        _app.WorkbookDeactivate -= OnWorkbookDeactivate;
        _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;
        _app = null;
        CellGitHistoryTip.Instance.ClearBubble();
    }

    private static void OnSelectionChange(object sh, Microsoft.Office.Interop.Excel.Range target)
    {
        CellGitHistoryTip.Instance.ClearBubble();
        CellGitHistoryTip.Instance.ResetAnchor();
        CellGitHistoryService.CancelPending();
        ExcelAsyncUtil.QueueAsMacro(() => TryQuery(sh, target));
    }

    private static void TryQuery(object sh, Microsoft.Office.Interop.Excel.Range target)
    {
        try
        {
            PluginLog.Verbose($"[谁的锅] TryQuery start row={target?.Row} col={target?.Column}");

            // 多选时跳过
            if (target.Cells.Count > 1)
            {
                PluginLog.Verbose("[谁的锅] skip: multi-select");
                return;
            }

            var wb = (Microsoft.Office.Interop.Excel.Workbook)AppServices.App.ActiveWorkbook;
            var ws = (Microsoft.Office.Interop.Excel.Worksheet)sh;
            var absFilePath = wb.FullName;

            if (!absFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                PluginLog.Verbose($"[谁的锅] skip: not xlsx ({absFilePath})");
                return;
            }

            // 从文件路径自动检测 git 仓库根目录（不依赖配置）
            var gitRoot = SvnGitTools.FindGitRoot(absFilePath);
            if (gitRoot == null)
            {
                PluginLog.Verbose($"[谁的锅] skip: no .git found for {absFilePath}");
                return;
            }

            int row = target.Row;
            int col = target.Column;
            if (row < 3)
            {
                PluginLog.Verbose($"[谁的锅] skip: header row {row}");
                return;
            }

            var sheetName = ws.Name;
            var colName = ws.Cells[2, col]?.Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(colName) || colName.StartsWith('#'))
            {
                PluginLog.Verbose($"[谁的锅] skip: colName='{colName}' (empty or #)");
                return;
            }

            // 找 key 列（row 2 中第一个非 # 列）
            int keyColIdx = 1;
            for (int c = 1; c <= 30; c++)
            {
                var h = ws.Cells[2, c]?.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(h) && !h.StartsWith('#'))
                {
                    keyColIdx = c;
                    break;
                }
            }

            var rowKey = ws.Cells[row, keyColIdx]?.Value?.ToString() ?? "";
            if (string.IsNullOrEmpty(rowKey))
            {
                PluginLog.Verbose($"[谁的锅] skip: rowKey empty at row={row} keyCol={keyColIdx}");
                return;
            }

            PluginLog.Write(
                $"[谁的锅] querying: file={System.IO.Path.GetFileName(absFilePath)} sheet={sheetName} row={row} col={colName} key={rowKey} gitRoot={gitRoot}"
            );

            // QueueAsMacro：把 ShowBubble 排入 Excel 主线程执行（与放大镜气泡做法一致）
            Action<List<CellHistoryEntry>> onResult = results =>
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    PluginLog.Verbose($"[谁的锅] ShowBubble count={results.Count}");
                    CellGitHistoryTip.Instance.ShowBubble(results);
                });
            CellGitHistoryService.Query(absFilePath, gitRoot, sheetName, rowKey, colName, onResult);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[谁的锅] TryQuery exception: {ex.Message}");
        }
    }

    private static void OnWindowDeactivate(object wb, object wn)
    {
        // 用户点击气泡后气泡激活→Excel 触发 WindowDeactivate；此时不清，保留气泡供选文本复制。
        if (CellGitHistoryTip.Instance.IsBubbleActive)
            return;
        CellGitHistoryTip.Instance.ClearBubble();
    }

    private static void OnWorkbookDeactivate(object wb)
    {
        if (CellGitHistoryTip.Instance.IsBubbleActive)
            return;
        CellGitHistoryTip.Instance.ClearBubble();
    }

    private static void OnWorkbookBeforeClose(
        Microsoft.Office.Interop.Excel.Workbook wb,
        ref bool cancel
    ) => CellGitHistoryTip.Instance.ClearBubble();
}
