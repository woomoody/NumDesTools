using System.IO;

namespace NumDesTools.XlsxEditor;

/// <summary>
/// 原子写结果：成功与否、生成的备份路径（若有）、失败时的异常、残留的 tmp 路径（若有，供排查）。
/// </summary>
public readonly record struct AtomicWriteResult(
    bool Succeeded,
    string? BackupPath,
    string? LeftoverTempPath,
    Exception? Error
);

/// <summary>
/// 纯 BCL 原子文件写：不依赖 WPF/DataGrid/UI，可单测。
/// 语义：先让 <paramref name="writeToTemp"/> 委托把内容写到同目录的 .tmp 文件，
/// 成功后用 <see cref="File.Replace(string, string, string?)"/> 把 tmp 原子替换到目标路径，
/// 并把被替换的旧内容留在 <c>目标路径.bak</c>。
/// 任何环节失败（委托抛异常/替换失败）时，原文件保持不变——绝不会处于半写状态。
/// tmp 文件在失败时保留在磁盘（不删），以便人工排查半成品；成功后 tmp 已被 File.Replace 消费。
/// </summary>
public static class AtomicFileWriter
{
    /// <summary>
    /// 原子写入 <paramref name="finalPath"/>。
    /// </summary>
    /// <param name="finalPath">最终目标文件路径。</param>
    /// <param name="writeToTemp">
    /// 写入委托：接收 tmp 文件路径，负责把完整内容写到该路径。
    /// 抛异常即视为写入失败，原文件不受影响。
    /// </param>
    public static AtomicWriteResult Write(string finalPath, Action<string> writeToTemp)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(finalPath);
        ArgumentNullException.ThrowIfNull(writeToTemp);

        var tempPath = finalPath + ".tmp";
        var backupPath = finalPath + ".bak";

        // 清掉可能残留的上次失败的 tmp，避免委托误以为是已有内容
        TryDelete(tempPath);

        try
        {
            writeToTemp(tempPath);
        }
        catch (Exception ex)
        {
            // 委托失败：原文件从未被触碰，保持完好。保留 tmp 供排查。
            return new AtomicWriteResult(
                Succeeded: false,
                BackupPath: null,
                LeftoverTempPath: File.Exists(tempPath) ? tempPath : null,
                Error: ex
            );
        }

        try
        {
            if (File.Exists(finalPath))
            {
                // 原文件存在：原子替换，旧内容进 .bak
                File.Replace(tempPath, finalPath, backupPath);
                return new AtomicWriteResult(
                    Succeeded: true,
                    BackupPath: backupPath,
                    LeftoverTempPath: null,
                    Error: null
                );
            }

            // 原文件不存在：无可备份，直接原子改名到位（Move 同盘为原子操作）
            File.Move(tempPath, finalPath);
            return new AtomicWriteResult(
                Succeeded: true,
                BackupPath: null,
                LeftoverTempPath: null,
                Error: null
            );
        }
        catch (Exception ex)
        {
            // 替换阶段失败：原文件仍是替换前的旧内容（File.Replace 失败不会破坏 destination）。
            return new AtomicWriteResult(
                Succeeded: false,
                BackupPath: null,
                LeftoverTempPath: File.Exists(tempPath) ? tempPath : null,
                Error: ex
            );
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (IOException)
        {
            // 删不掉旧 tmp 不致命：写入委托会覆写它
        }
        catch (UnauthorizedAccessException)
        {
            // 同上
        }
    }
}
