namespace NumDesTools.Tests;

public class SvnGitToolsTests
{
    // 回归测试：Repository.Discover 固定返回带尾部分隔符的 ".git" 路径，Directory.GetParent
    // 对带尾部分隔符的路径不会真的往上一级——曾经导致 FindGitRoot 返回 "...\.git" 而不是
    // 仓库根目录，Lua 导出时拼路径变成 "...\.git\Excels\Tables\xxx.xlsx"，找不到文件。
    [Fact]
    public void FindGitRoot_ReturnsRepoRoot_NotGitDirItself()
    {
        // 本仓库自己就是 git repo，用测试项目所在目录反查
        var root = SvnGitTools.FindGitRoot(AppContext.BaseDirectory);

        Assert.NotNull(root);
        Assert.False(root!.EndsWith(".git", StringComparison.OrdinalIgnoreCase));
        Assert.True(Directory.Exists(Path.Combine(root, ".git")));
    }
}
