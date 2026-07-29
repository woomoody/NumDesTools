using System.IO;
using NumDesTools.XlsxEditor;

namespace NumDesTools.Tests;

/// <summary>
/// AtomicFileWriter 原子写单测：验证 tmp + File.Replace 语义。
/// 核心保证：写入委托抛异常时，预先存在的原文件必须完好无损（不被截断/半写）。
/// </summary>
public sealed class AtomicFileWriterTests : IDisposable
{
    private readonly string _dir = Path.Combine(
        Path.GetTempPath(),
        "AtomicFileWriterTests_" + Guid.NewGuid().ToString("N")
    );

    public AtomicFileWriterTests() => Directory.CreateDirectory(_dir);

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch
        {
            // 清理失败不影响测试结论
        }
    }

    [Fact]
    public void Write_WhenDelegateSucceeds_ReplacesOldContent()
    {
        var target = Path.Combine(_dir, "data.xlsx");
        File.WriteAllText(target, "OLD-CONTENT");

        var result = AtomicFileWriter.Write(target, tmp => File.WriteAllText(tmp, "NEW-CONTENT"));

        Assert.True(result.Succeeded);
        Assert.Equal("NEW-CONTENT", File.ReadAllText(target));
    }

    [Fact]
    public void Write_WhenTargetDoesNotExist_CreatesFile()
    {
        var target = Path.Combine(_dir, "fresh.xlsx");
        Assert.False(File.Exists(target));

        var result = AtomicFileWriter.Write(target, tmp => File.WriteAllText(tmp, "BRAND-NEW"));

        Assert.True(result.Succeeded);
        Assert.Equal("BRAND-NEW", File.ReadAllText(target));
    }

    [Fact]
    public void Write_WhenDelegateThrows_LeavesOriginalIntact()
    {
        var target = Path.Combine(_dir, "data.xlsx");
        File.WriteAllText(target, "ORIGINAL-INTACT");

        var result = AtomicFileWriter.Write(
            target,
            tmp =>
            {
                // 模拟保存中途崩溃：写了一半再抛异常
                File.WriteAllText(tmp, "HALF-WRITTEN-GARBAGE");
                throw new InvalidOperationException("simulated save crash");
            }
        );

        Assert.False(result.Succeeded);
        Assert.NotNull(result.Error);
        Assert.IsType<InvalidOperationException>(result.Error);
        // 关键断言：原文件内容完好，未被截断/损坏
        Assert.Equal("ORIGINAL-INTACT", File.ReadAllText(target));
    }

    [Fact]
    public void Write_WhenSucceeds_ReplacesContentWithoutBackup()
    {
        var target = Path.Combine(_dir, "data.xlsx");
        File.WriteAllText(target, "V1-CONTENT");

        var result = AtomicFileWriter.Write(target, tmp => File.WriteAllText(tmp, "V2-CONTENT"));

        Assert.True(result.Succeeded);
        Assert.Null(result.BackupPath);
        // 不生成 .bak（git 管备份）
        Assert.False(File.Exists(target + ".bak"));
        Assert.False(File.Exists(target + ".bak~"));
        Assert.Equal("V2-CONTENT", File.ReadAllText(target));
    }

    [Fact]
    public void Write_WhenDelegateThrows_DoesNotLeaveOriginalTruncated()
    {
        var target = Path.Combine(_dir, "big.xlsx");
        var original = new string('X', 100_000);
        File.WriteAllText(target, original);
        var originalLength = new FileInfo(target).Length;

        var result = AtomicFileWriter.Write(
            target,
            tmp => throw new IOException("delegate failed before any write")
        );

        Assert.False(result.Succeeded);
        Assert.Equal(originalLength, new FileInfo(target).Length);
        Assert.Equal(original, File.ReadAllText(target));
    }
}
