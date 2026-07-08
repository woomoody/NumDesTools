using NumDesTools.Config;

namespace NumDesTools.Tests;

public class GlobalVariableTests
{
    private static string TmpPath(string name) =>
        Path.Combine(
            Path.GetTempPath(),
            "NumDesTools.Tests",
            "tmp",
            name
        );

    // H15: JSON 损坏时 ReadOrCreate 不应抛异常，应降级到默认值
    [Fact]
    public void ReadOrCreate_CorruptJson_DoesNotThrow()
    {
        var path = TmpPath("test_corrupt_config.json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "{ NOT VALID JSON !!! ");

        var gv = new GlobalVariable(path);

        Assert.NotNull(gv.Value);
        Assert.True(gv.Value.ContainsKey("LiteLLMModel"));
    }

    [Fact]
    public void ReadOrCreate_EmptyFile_DoesNotThrow()
    {
        var path = TmpPath("test_empty_config.json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "");

        var gv = new GlobalVariable(path);

        Assert.NotNull(gv.Value);
    }

    [Fact]
    public void ReadOrCreate_MissingFile_CreatesDefaults()
    {
        var path = TmpPath("test_missing_config_abc123.json");
        if (File.Exists(path))
            File.Delete(path);

        var gv = new GlobalVariable(path);

        Assert.NotNull(gv.Value);
        Assert.True(gv.Value.ContainsKey("LiteLLMModel"));
        Assert.True(File.Exists(path));
        File.Delete(path);
    }
}
