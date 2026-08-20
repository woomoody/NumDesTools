using NumDesTools;

namespace NumDesTools.Tests;

public class CellGitHistoryLuaIndexTests
{
    [Fact]
    public void GetExportLuaBaseName_DollarWorkbook_UsesSheetName()
    {
        Assert.Equal("Help", CellGitHistoryLuaIndex.GetExportLuaBaseName("$帮助", "Help"));
        Assert.Equal(
            "NormalTable",
            CellGitHistoryLuaIndex.GetExportLuaBaseName("NormalTable", "OtherSheet")
        );
    }

    [Fact]
    public void ResolveTargetLuaFile_DollarWorkbookSplitRow_UsesSheetNamedShard()
    {
        var root = Path.Combine(Path.GetTempPath(), $"ndt_lua_{Guid.NewGuid():N}");
        var mainLua = Path.Combine(root, "Help.lua.txt");
        var shardLua = Path.Combine(root, "Help_1001.lua.txt");

        try
        {
            Directory.CreateDirectory(root);
            File.WriteAllText(mainLua, "[3] = 1001,\n");
            File.WriteAllText(shardLua, "[3] = { id = 3, },\n");

            var target = CellGitHistoryLuaIndex.ResolveTargetLuaFile(
                mainLua,
                CellGitHistoryLuaIndex.GetExportLuaBaseName("$帮助", "Help")!,
                "3",
                out var reason
            );

            Assert.Equal(shardLua, target);
            Assert.Empty(reason);
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void ResolveTargetLuaFile_MissingShard_ReportsExpectedPath()
    {
        var root = Path.Combine(Path.GetTempPath(), $"ndt_lua_{Guid.NewGuid():N}");
        var mainLua = Path.Combine(root, "Help.lua.txt");

        try
        {
            Directory.CreateDirectory(root);
            File.WriteAllText(mainLua, "[3] = 1001,\n");

            var target = CellGitHistoryLuaIndex.ResolveTargetLuaFile(
                mainLua,
                "Help",
                "3",
                out var reason
            );

            Assert.Null(target);
            Assert.Contains("Help_1001.lua.txt", reason, StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(root))
                Directory.Delete(root, recursive: true);
        }
    }
}
