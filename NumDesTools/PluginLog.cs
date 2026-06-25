using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text;

namespace NumDesTools;

/// <summary>
/// 追加写日志到 我的文档\workspace\plugin.log，超过 2MB 自动保留最新一半。
/// 同时维护内存缓冲 Lines，供 PluginLogWindow 实时绑定展示。
/// </summary>
internal static class PluginLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "workspace",
        "plugin.log"
    );

    private const long MaxBytes = 2 * 1024 * 1024; // 2 MB
    private static readonly object _lock = new();

    // ── 内存日志缓冲（供 UI 绑定）────────────────────────────────────────
    // _pending：任意线程安全地入队；PluginLogWindow 的定时器在 UI 线程定期排空到 _lines
    private static readonly ConcurrentQueue<string> _pending = new();
    private static readonly ObservableCollection<string> _lines = new();
    private const int MaxLines = 2000;

    /// <summary>内存日志行集合，供 PluginLogWindow 绑定。只在 UI 线程访问。</summary>
    public static ObservableCollection<string> Lines => _lines;

    /// <summary>由 PluginLogWindow 的定时器在 UI 线程调用，把积压的日志刷入 Lines。</summary>
    public static void DrainPendingToUi()
    {
        while (_pending.TryDequeue(out var line))
        {
            _lines.Add(line);
            while (_lines.Count > MaxLines)
                _lines.RemoveAt(0);
        }
    }

    // ── 原有方法（不改动）────────────────────────────────────────────────

    public static void Write(string message)
    {
        try
        {
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
            lock (_lock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath, line, Encoding.UTF8);
                TrimIfNeeded();
            }
            Debug.Print(line.TrimEnd());
            AppendToUi(line.TrimEnd());
        }
        catch { }
    }

    /// <summary>
    /// 仅 Debug 构建写入，Release 时编译器连调用点和参数求值一起消除，零性能开销。
    /// </summary>
    [Conditional("DEBUG")]
    public static void Verbose(string message) => Write(message);

    // ── 新增方法 ──────────────────────────────────────────────────────────

    /// <summary>
    /// 记录一行日志并同步到内存 UI 列表。
    /// 调用方通常已自带 [DateTime.Now] 前缀，此方法原样写入，不再重复加时间戳。
    /// </summary>
    public static void RecordLine(string format, params object[] args)
    {
        var line = args.Length > 0 ? string.Format(format, args) : format;
        WriteRaw(line);
        AppendToUi(line);
    }

    // ── 私有工具 ──────────────────────────────────────────────────────────

    private static void WriteRaw(string line)
    {
        try
        {
            lock (_lock)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath, line + Environment.NewLine, Encoding.UTF8);
                TrimIfNeeded();
            }
            Debug.Print(line);
        }
        catch { }
    }

    private static void AppendToUi(string line)
    {
        // 任意线程安全入队，由 PluginLogWindow 的定时器在 UI 线程排空
        _pending.Enqueue(line);
    }

    private static void TrimIfNeeded()
    {
        try
        {
            var info = new FileInfo(LogPath);
            if (info.Length < MaxBytes)
                return;
            var content = File.ReadAllText(LogPath, Encoding.UTF8);
            var half = content.Length / 2;
            var cut = content.IndexOf('\n', half);
            if (cut < 0)
                return;
            File.WriteAllText(LogPath, content[(cut + 1)..], Encoding.UTF8);
        }
        catch { }
    }
}
