using System.Diagnostics;
using System.Text;

namespace NumDesTools;

/// <summary>
/// 追加写日志到 我的文档\workspace\plugin.log，超过 2MB 自动保留最新一半。
/// </summary>
internal static class PluginLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "workspace",
        "plugin.log"
    );

    private const long MaxBytes = 2 * 1024 * 1024; // 2 MB

    public static void Write(string message)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
            File.AppendAllText(LogPath, line, Encoding.UTF8);
            TrimIfNeeded();
            Debug.Print(line.TrimEnd());
        }
        catch { }
    }

    /// <summary>
    /// 仅 Debug 构建写入，Release 时编译器连调用点和参数求值一起消除，零性能开销。
    /// 用于每行级别的高频诊断日志（ApplyIdMap、Detect HIT、Remark 等）。
    /// </summary>
    [System.Diagnostics.Conditional("DEBUG")]
    public static void Verbose(string message) => Write(message);

    private static void TrimIfNeeded()
    {
        try
        {
            var info = new FileInfo(LogPath);
            if (info.Length < MaxBytes)
                return;
            var content = File.ReadAllText(LogPath, Encoding.UTF8);
            // 保留后半段，从中间某个换行处截断
            var half = content.Length / 2;
            var cut = content.IndexOf('\n', half);
            if (cut < 0)
                return;
            File.WriteAllText(LogPath, content[(cut + 1)..], Encoding.UTF8);
        }
        catch { }
    }
}
