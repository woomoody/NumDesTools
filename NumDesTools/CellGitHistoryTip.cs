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
    private static readonly Dictionary<
        string,
        Dictionary<string, Dictionary<string, string>>
    > _sheetDataCache = new(StringComparer.Ordinal);
    private static readonly Queue<string> _sheetCacheOrder = new();
    private const int SheetCacheCapacity = 30;

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
    /// 流式查询：每找到一条真实变更立刻调用 onPartial 刷新气泡，全部找完后调用 onFinal 缓存。
    /// 修正对比逻辑：commit[i].val ≠ commit[i+1].val 才说明 commit[i] 真的改了这格。
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

        // 滑动窗口流式：读完 commit[i+1] 就能判断 commit[i] 是否是真实改动者，无需等全部收集
        const int MaxChanges = 5;
        const int MaxLoop = 200; // 兜底上限，防止极大文件历史无限扫描
        var accumulated = new List<(string date, string author, string msg, string val)>();

        string? prevVal = null;
        (string date, string author, string msg)? prevMeta = null;

        // P0 优化：Repository 实例提升到循环外，50 commits 只初始化一次
        using var repo = new Repository(gitRoot);
        string? prevBlobOid = null; // blob OID 预过滤：OID 相同 = 文件内容未变 = 单元格值一定未变
        int loopCount = 0;
        bool hadNonNull = false; // 是否曾经找到过值（用于检测"此行首次出现"边界）

        foreach (var (sha, date, author, msg) in commits)
        {
            if (ct.IsCancellationRequested)
            {
                onFinal(null);
                return;
            }
            if (accumulated.Count >= MaxChanges || loopCount++ >= MaxLoop)
                break;

            // blob OID 预过滤：纳秒级内存操作，文件内容未变直接跳过
            var commit = repo.Lookup<Commit>(sha);
            var blobEntry = commit?.Tree[relativePath];
            var blobOid = blobEntry?.Target.Sha;
            if (blobOid != null && blobOid == prevBlobOid)
            {
                // 文件内容与上一个 commit 完全相同，单元格值一定未变，跳过昂贵的 MiniExcel 解析
                PluginLog.Verbose($"[谁的锅] commit {sha[..8]} blob unchanged, skip");
                continue;
            }
            prevBlobOid = blobOid;

            var val = GetCellValueAtCommit(
                repo,
                sha,
                relativePath,
                sheetName,
                rowKey,
                colName,
                tmpDir
            );
            PluginLog.Verbose($"[谁的锅] commit {sha[..8]} val={val ?? "null"}");

            if (val == null)
            {
                if (hadNonNull)
                {
                    // 找到创建边界：更老的 commit 不存在此行，说明 prevMeta 是"首次添加此行"的 commit
                    // 直接跳出，后续处理会将 prevMeta 作为"最早可查记录"显示
                    break;
                }
                continue; // 此行在这个 commit 里不存在，继续向更老的 commit 找
            }

            hadNonNull = true;

            if (prevVal != null && prevMeta.HasValue && val != prevVal)
            {
                // prevMeta 对应的 commit 是真实改动者（它的值与再早一个 commit 不同）
                accumulated.Add(
                    (prevMeta.Value.date, prevMeta.Value.author, prevMeta.Value.msg, prevVal)
                );
                onPartial(BuildText(accumulated));
            }

            prevVal = val;
            prevMeta = (date, author, msg);
        }

        // 若找到值但无任何变更（值在所有可查提交中从未改变），
        // 则展示最老一条提交作为"最早可查记录"，告诉用户是谁最早放入了这个值
        if (accumulated.Count == 0 && prevMeta.HasValue && prevVal != null)
        {
            var stableEntry = (
                prevMeta.Value.date,
                prevMeta.Value.author,
                prevMeta.Value.msg + "（最早可查，值未改变）",
                prevVal
            );
            accumulated.Add(stableEntry);
            onPartial(BuildText(accumulated));
        }

        PluginLog.Write($"[谁的锅] streaming done: changes={accumulated.Count}");
        var finalText = accumulated.Count > 0 ? BuildText(accumulated) : null;
        onFinal(finalText);
    }

    private static string BuildText(
        List<(string date, string author, string msg, string val)> results
    )
    {
        var sb = new StringBuilder();
        for (int i = 0; i < results.Count; i++)
        {
            var (date, author, msg, val) = results[i];
            var datePart = date.Length >= 10 ? date[..10] : date;
            sb.AppendLine($"[{i + 1}] {datePart}  {author}");
            var shortMsg = msg.Length > 40 ? msg[..40] + "…" : msg;
            sb.AppendLine($"    {shortMsg}");
            var shortVal = val.Length > 60 ? val[..60] + "…" : val;
            if (i < results.Count - 1)
                sb.AppendLine($"    值: {shortVal}");
            else
                sb.Append($"    值: {shortVal}");
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
            // --all 搜索所有分支；-n 200 防止极大仓库无限返回（QueryHistoryStreaming 有 MaxLoop=200 兜底）
            var args = $"log --all -n 200 --format=\"%H|%ai|%an|%s\" -- \"{relativePath}\"";
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
    /// </summary>
    private static Dictionary<string, Dictionary<string, string>>? LoadSheetData(
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
                using var repo = new Repository(gitRoot);
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

            // 入缓存（LRU 淘汰）
            if (_sheetDataCache.Count >= SheetCacheCapacity)
            {
                var old = _sheetCacheOrder.Dequeue();
                _sheetDataCache.Remove(old);
            }
            _sheetDataCache[cacheKey] = data;
            _sheetCacheOrder.Enqueue(cacheKey);
            return data;
        }
        catch
        {
            return null;
        }
    }

    // 单格值缓存：sha8|relPath|sheetName|rowKey|colName → value
    private static readonly Dictionary<string, string> _cellValCache =
        new(StringComparer.Ordinal);

    /// <summary>
    /// 用 MiniExcel 流式读取，找到目标 rowKey 立即 break，无需解析整个 sheet。
    /// 比 EPPlus 全量加载快 10-20x（20-50ms vs 200-500ms per commit）。
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
            var tmpFile = Path.Combine(tmpDir, $"{sha[..8]}_{Path.GetFileName(relativePath)}");
            if (!File.Exists(tmpFile))
            {
                // 使用调用方传入的 repo 实例，不再重复初始化
                var commit = repo.Lookup<Commit>(sha);
                if (commit == null) return null;
                var entry = commit[relativePath];
                if (entry == null) return null;
                var blob = (Blob)entry.Target;
                using var src = blob.GetContentStream();
                using var dst = new FileStream(tmpFile, FileMode.Create, FileAccess.Write);
                src.CopyTo(dst);
            }

            // MiniExcel 流式读，row 2 找列字母，row 3+ 找 rowKey，找到即 break
            string? keyLetter = null, targetLetter = null;
            int rowIdx = 0;
            string? result = null;

            foreach (IDictionary<string, object> row in
                MiniExcelLibs.MiniExcel.Query(tmpFile, sheetName: sheetName, useHeaderRow: false)
                    .Cast<IDictionary<string, object>>())
            {
                rowIdx++;
                if (rowIdx == 2)
                {
                    // 行2 = 列名行，按列字母排序找 key 列和目标列
                    foreach (var kv in row.OrderBy(k => k.Key.Length).ThenBy(k => k.Key))
                    {
                        var h = kv.Value?.ToString() ?? "";
                        if (keyLetter == null && !string.IsNullOrEmpty(h) && !h.StartsWith('#'))
                            keyLetter = kv.Key;
                        if (h == colName)
                            targetLetter = kv.Key;
                    }
                    if (keyLetter == null) break; // 找不到 key 列，放弃
                    if (targetLetter == null)
                    {
                        // 该列不存在于此 commit
                        result = "(列当时不存在)";
                        break;
                    }
                }
                else if (rowIdx >= 3 && keyLetter != null && targetLetter != null)
                {
                    var kv2 = row.TryGetValue(keyLetter, out var kv) ? kv?.ToString() ?? "" : "";
                    if (kv2 == rowKey)
                    {
                        var val = row.TryGetValue(targetLetter, out var tv) ? tv?.ToString() ?? "" : "";
                        result = string.IsNullOrEmpty(val) ? "(空)" : val;
                        break; // 找到目标行，立即停止
                    }
                }
            }

            if (result != null)
                _cellValCache[cacheKey] = result;
            return result;
        }
        catch
        {
            return null;
        }
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

    private static string? FindGitRoot(string filePath)
    {
        var dir = Path.GetDirectoryName(filePath);
        while (!string.IsNullOrEmpty(dir))
        {
            if (Directory.Exists(Path.Combine(dir, ".git")))
                return dir;
            dir = Path.GetDirectoryName(dir);
        }
        return null;
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
            var gitRoot = FindGitRoot(absFilePath);
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
