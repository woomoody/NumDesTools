using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using LibGit2Sharp;
using OfficeOpenXml;
using Font = System.Drawing.Font;
using Timer = System.Windows.Forms.Timer;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 悬浮气泡：选中单元格后显示该格最近 2 次 git 提交的历史值。
/// 不抢焦点，鼠标透传；滚动或 Excel 失焦时自动隐藏。
/// </summary>
public sealed class CellGitHistoryTip : Form
{
    private string[]? _lines;

    private static readonly Font _headerFont = new("微软雅黑", 9.5f, FontStyle.Bold);
    private static readonly Font _bodyFont = new("微软雅黑", 9f);
    private const int Pad = 10;
    private const int LineGap = 3;

    private readonly Timer _scrollTimer;
    private readonly Timer _focusTimer;
    private int _lastScrollRow;
    private int _lastScrollCol;

    private static CellGitHistoryTip? _instance;
    public static CellGitHistoryTip Instance => _instance ??= new CellGitHistoryTip();

    private CellGitHistoryTip()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(22, 27, 34);
        ForeColor = Color.FromArgb(220, 220, 220);
        AutoScaleMode = AutoScaleMode.None;
        StartPosition = FormStartPosition.Manual;
        SetStyle(
            ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint,
            true
        );

        var ex = GetWindowLong(Handle, GWL_EXSTYLE);
        SetWindowLong(Handle, GWL_EXSTYLE, ex | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE);

        _scrollTimer = new Timer { Interval = 150 };
        _scrollTimer.Tick += OnScrollCheck;

        _focusTimer = new Timer { Interval = 300 };
        _focusTimer.Tick += OnFocusCheck;
        _focusTimer.Start();
    }

    protected override CreateParams CreateParams
    {
        get
        {
            var cp = base.CreateParams;
            cp.ExStyle |= WS_EX_TRANSPARENT | WS_EX_NOACTIVATE;
            return cp;
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.Clear(BackColor);
        if (_lines == null)
            return;

        // 绘制左侧竖线装饰
        using (var lineBrush = new SolidBrush(Color.FromArgb(80, 130, 200)))
            e.Graphics.FillRectangle(lineBrush, 0, 0, 3, ClientSize.Height);

        float y = Pad;
        foreach (var line in _lines)
        {
            var isHeader = line.StartsWith('[');
            var font = isHeader ? _headerFont : _bodyFont;
            var color = isHeader ? Color.FromArgb(100, 180, 255) : ForeColor;
            using var brush = new SolidBrush(color);
            e.Graphics.DrawString(line, font, brush, new PointF(Pad + 4, y));
            y += font.GetHeight(e.Graphics) + LineGap;
        }
    }

    public void ShowBubble(string text)
    {
        _lines = text.Split('\n');

        // 计算气泡尺寸
        float maxW = 0;
        float totalH = Pad * 2;
        using var g = CreateGraphics();
        foreach (var line in _lines)
        {
            var font = line.StartsWith('[') ? _headerFont : _bodyFont;
            var sz = g.MeasureString(line, font);
            if (sz.Width > maxW)
                maxW = sz.Width;
            totalH += font.GetHeight(g) + LineGap;
        }

        int w = (int)maxW + Pad * 2 + 8;
        int h = (int)totalH;

        var cursor = Cursor.Position;
        int x = cursor.X + 16;
        int y = cursor.Y + 16;
        var wa = Screen.FromPoint(cursor).WorkingArea;
        if (x + w > wa.Right)
            x = cursor.X - w - 4;
        if (y + h > wa.Bottom)
            y = cursor.Y - h - 4;
        if (x < wa.Left)
            x = wa.Left;
        if (y < wa.Top)
            y = wa.Top;

        ClientSize = new Size(w, h);
        Location = new Point(x, y);
        ShowWindow(Handle, SW_SHOWNOACTIVATE);
        Invalidate();

        try
        {
            var win = AppServices.App.ActiveWindow;
            _lastScrollRow = win.ScrollRow;
            _lastScrollCol = win.ScrollColumn;
        }
        catch { }
        _scrollTimer.Start();
    }

    public void ClearBubble()
    {
        _scrollTimer.Stop();
        _lines = null;
        if (!IsHandleCreated || IsDisposed)
            return;
        if (InvokeRequired)
            BeginInvoke((System.Action)Hide);
        else
            Hide();
    }

    private void OnScrollCheck(object? sender, EventArgs e)
    {
        try
        {
            var win = AppServices.App.ActiveWindow;
            if (win.ScrollRow != _lastScrollRow || win.ScrollColumn != _lastScrollCol)
                ClearBubble();
        }
        catch
        {
            ClearBubble();
        }
    }

    private void OnFocusCheck(object? sender, EventArgs e)
    {
        if (!Visible)
            return;
        try
        {
            var fg = GetForegroundWindow();
            if (fg == Handle)
                return;
            GetWindowThreadProcessId(fg, out uint fgPid);
            GetWindowThreadProcessId((IntPtr)AppServices.App.Hwnd, out uint excelPid);
            if (fgPid != excelPid)
                ClearBubble();
        }
        catch { }
    }

    public static void DisposeInstance()
    {
        if (_instance is { IsDisposed: false })
        {
            _instance._scrollTimer.Dispose();
            _instance._focusTimer.Dispose();
            _instance.Close();
            _instance.Dispose();
        }
        _instance = null;
    }

    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_TRANSPARENT = 0x20;
    private const int WS_EX_NOACTIVATE = 0x8000000;
    private const int SW_SHOWNOACTIVATE = 4;

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}

// ── 后台查询服务 ─────────────────────────────────────────────────────────────

internal static class CellGitHistoryService
{
    private static CancellationTokenSource? _cts;

    // LRU 缓存：key = "absFile|sheet|rowKey|colName"
    private static readonly Dictionary<string, string> _cache = new(StringComparer.Ordinal);
    private static readonly Queue<string> _cacheOrder = new();
    private const int CacheCapacity = 100;

    // 文件级 commit 列表缓存：key = absFilePath → (list, fileLastWriteStamp)
    private static readonly Dictionary<
        string,
        (List<(string sha, string date, string author, string msg)> commits, long stamp)
    > _commitListCache = new(StringComparer.OrdinalIgnoreCase);

    // Sheet 级数据缓存：key = "sha8|relPath|sheetName" → rowKey → colName → value
    // 一次 EPPlus 解析覆盖整 sheet，同一 sheet 的多格查询直接命中
    // 使用 ConcurrentDictionary 支持并行分块处理下的线程安全读写
    private static readonly ConcurrentDictionary<
        string,
        Dictionary<string, Dictionary<string, string>>
    > _sheetDataCache = new(StringComparer.Ordinal);
    private const int SheetCacheCapacity = 500;

    public static void Query(
        string absFilePath,
        string gitRoot,
        string sheetName,
        string rowKey,
        string colName,
        Action<string> onResult
    )
    {
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;

        var cacheKey = $"{absFilePath}|{sheetName}|{rowKey}|{colName}";
        if (_cache.TryGetValue(cacheKey, out var cached))
        {
            onResult(cached); // 缓存命中：直接返回，不触发 ribbon 状态变化
            return;
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
                        partialText =>
                        {
                            if (ct.IsCancellationRequested)
                                return;
                            onResult(partialText); // 每找到一条就刷新气泡
                        },
                        finalText =>
                        {
                            if (!ct.IsCancellationRequested && finalText != null)
                                PutCache(cacheKey, finalText); // 全部找完后缓存最终结果
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

    private static void PutCache(string key, string value)
    {
        if (_cache.ContainsKey(key))
            return;
        if (_cache.Count >= CacheCapacity)
        {
            var old = _cacheOrder.Dequeue();
            _cache.Remove(old);
        }
        _cache[key] = value;
        _cacheOrder.Enqueue(key);
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
        System.Action<string> onPartial,
        System.Action<string?> onFinal
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

        const int MaxChanges = 50;
        const int MaxCommits = 500;
        const int StreamingPhaseCount = 25; // 前 25 个 commit 顺序流式，快速出首条
        const int ParallelChunks = 6;

        var takeCount = Math.Min(commits.Count, MaxCommits);
        var limitedCommits = commits.GetRange(0, takeCount);

        // 共享的 accumulated 列表，两阶段共用
        var accumulated = new List<(string date, string author, string msg, string oldVal, string newVal)>();

        // ── 阶段1：顺序流式（前 StreamingPhaseCount 个 commit）─────────────
        // 实时出结果，用户立刻看到
        int streamingEnd = Math.Min(StreamingPhaseCount, takeCount);
        string? prevVal = null;
        (string date, string author, string msg)? prevMeta = null;
        bool hadNonNull = false;

        // 供阶段2用的边界值（阶段1最后的 prevVal / prevMeta）
        string? phase1LastVal = null;
        (string date, string author, string msg)? phase1LastMeta = null;

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

                // 提取值
                string? val = null;
                var sheetData = LoadSheetData(streamRepo, gitRoot, sha, relativePath, sheetName, tmpDir);
                if (sheetData != null && sheetData.TryGetValue(rowKey, out var rowData))
                    rowData.TryGetValue(colName, out val);

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
                    accumulated.Add((prevMeta.Value.date, prevMeta.Value.author, prevMeta.Value.msg, val, prevVal));
                    onPartial(BuildText(accumulated));
                }

                prevVal = val;
                prevMeta = (date, author, msg);
            }

            phase1LastVal = prevVal;
            phase1LastMeta = prevMeta;
        }

        // ── 阶段2：并行处理剩余 commit ─────────────────────────────────────
        if (streamingEnd < takeCount && accumulated.Count < MaxChanges && !ct.IsCancellationRequested)
        {
            int remaining = takeCount - streamingEnd;
            int chunkSize = Math.Max(1, (remaining + ParallelChunks - 1) / ParallelChunks);
            int chunkCount = (remaining + chunkSize - 1) / chunkSize;
            var chunkResults = new (string? val, string sha, string date, string author, string msg)[chunkCount][];

            Parallel.For(0, chunkCount,
                new ParallelOptions { MaxDegreeOfParallelism = ParallelChunks, CancellationToken = ct }, chunkIdx =>
            {
                int start = streamingEnd + chunkIdx * chunkSize;
                int end = Math.Min(start + chunkSize, takeCount);

                using var threadRepo = new Repository(gitRoot);
                string? prevBlobOid = null;
                var local = new (string? val, string sha, string date, string author, string msg)[end - start];

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
                        PluginLog.Verbose($"[谁的锅] commit {sha[..8]} blob unchanged, skip");
                        continue;
                    }
                    prevBlobOid = blobOid;

                    string? val = null;
                    var sheetData = LoadSheetData(threadRepo, gitRoot, sha, relativePath, sheetName, tmpDir);
                    if (sheetData != null && sheetData.TryGetValue(rowKey, out var rowData))
                        rowData.TryGetValue(colName, out val);

                    local[i - start] = (val, sha, date, author, msg);
                }
                chunkResults[chunkIdx] = local;
            });

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
                        accumulated.Add((prevMeta.Value.date, prevMeta.Value.author, prevMeta.Value.msg, val, prevVal));
                        onPartial(BuildText(accumulated));
                    }

                    prevVal = val;
                    prevMeta = (date, author, msg);
                }
            }
            Phase2Done:;
        }

        // 若找到值但无任何变更，展示最老一条
        if (accumulated.Count == 0 && prevMeta.HasValue && prevVal != null)
        {
            accumulated.Add((prevMeta.Value.date, prevMeta.Value.author, prevMeta.Value.msg + "（最早可查，值未改变）", prevVal, prevVal));
            onPartial(BuildText(accumulated));
        }

        PluginLog.Write($"[谁的锅] hybrid done: changes={accumulated.Count}");
        var finalText = accumulated.Count > 0 ? BuildText(accumulated) : null;
        onFinal(finalText);
    }

    private static string BuildText(
        List<(string date, string author, string msg, string oldVal, string newVal)> results
    )
    {
        var sb = new StringBuilder();
        for (int i = 0; i < results.Count; i++)
        {
            var (date, author, msg, oldVal, newVal) = results[i];
            var datePart = date.Length >= 10 ? date[..10] : date;
            sb.AppendLine($"[{i + 1}] {datePart}  {author}");
            var shortMsg = msg.Length > 40 ? msg[..40] + "…" : msg;
            sb.AppendLine($"    {shortMsg}");
            var shortOld = oldVal.Length > 60 ? oldVal[..60] + "…" : oldVal;
            var shortNew = newVal.Length > 60 ? newVal[..60] + "…" : newVal;
            if (i < results.Count - 1)
                sb.AppendLine($"    旧值: {shortOld}  →  新值: {shortNew}");
            else
                sb.Append($"    旧值: {shortOld}  →  新值: {shortNew}");
        }
        return sb.ToString().TrimEnd('\n');
    }

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
                var date = parts[1].Trim().Length >= 10 ? parts[1].Trim()[..10] : parts[1].Trim();
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
    /// 加载并缓存某个 commit 下 xlsx 某 sheet 的全部数据：rowKey → colName → value。
    /// 一次 EPPlus 解析覆盖整个 sheet，同 sheet 后续格查询直接命中内存缓存。
    /// 线程安全：_sheetDataCache 使用 ConcurrentDictionary，支持并行分块处理。
    /// </summary>
    private static Dictionary<string, Dictionary<string, string>>? LoadSheetData(
        Repository repo,
        string gitRoot,
        string sha,
        string relativePath,
        string sheetName,
        string tmpDir
    )
    {
        var cacheKey = $"{sha[..8]}|{relativePath}|{sheetName}";
        if (_sheetDataCache.TryGetValue(cacheKey, out var cached))
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

            ExcelPackage.License.SetNonCommercialPersonal("NumDesTools");
            using var pkg = new ExcelPackage(new FileInfo(tmpFile));
            var ws = pkg.Workbook.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase)
            );
            if (ws?.Dimension == null)
                return null;

            var data = CellHistoryXlsxReader.ParseSheetData(ws);

            // 入缓存（ConcurrentDictionary 线程安全，超过容量时简单清理）
            if (_sheetDataCache.Count >= SheetCacheCapacity)
            {
                // 直接清空——简化 LRU，反正同一 sheet 的 commit 数据是连续访问的
                _sheetDataCache.Clear();
            }
            _sheetDataCache[cacheKey] = data;
            return data;
        }
        catch
        {
            return null;
        }
    }

    // ── 不再使用 MiniExcel 逐行扫描路径（已改用 LoadSheetData 全表解析+字典） ──
    // 旧代码 GetCellValueAtCommit 和 _cellValCache 已清理，如需恢复可 git checkout 此文件历史版本。

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

            PluginLog.Write($"[谁的锅] querying: file={System.IO.Path.GetFileName(absFilePath)} sheet={sheetName} row={row} col={colName} key={rowKey} gitRoot={gitRoot}");

            // QueueAsMacro：把 ShowBubble 排入 Excel 主线程执行（与放大镜气泡做法一致）
            Action<string> onResult = text =>
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    PluginLog.Verbose($"[谁的锅] ShowBubble text.len={text?.Length}");
                    CellGitHistoryTip.Instance.ShowBubble(text!);
                });
            CellGitHistoryService.Query(absFilePath, gitRoot, sheetName, rowKey, colName, onResult);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"[谁的锅] TryQuery exception: {ex.Message}");
        }
    }

    private static void OnWindowDeactivate(object wb, object wn) =>
        CellGitHistoryTip.Instance.ClearBubble();

    private static void OnWorkbookDeactivate(object wb) => CellGitHistoryTip.Instance.ClearBubble();

    private static void OnWorkbookBeforeClose(
        Microsoft.Office.Interop.Excel.Workbook wb,
        ref bool cancel
    ) => CellGitHistoryTip.Instance.ClearBubble();
}
