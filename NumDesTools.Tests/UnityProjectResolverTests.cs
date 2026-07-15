using System.IO;
using NumDesTools.ExcelToLua;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// 验证 ExcelToLua 路径锚点核心逻辑：规范化、Unity 项目识别、从 Excel 表目录推断 Unity 项目根。
/// 这套逻辑取代了原硬编码 "Code/Assets/..."，避免别人项目目录不叫 Code 时写飞。
/// 全部用临时目录，不依赖真实 C:\M1Work 结构，CI 也能跑。
/// </summary>
public class UnityProjectResolverTests
{
    // ── Normalize：路径规范化（缓存键）──────────────────────────────────────

    [Fact]
    public void Normalize_TrimsTrailingSlash_And_Lowercases()
    {
        var n = UnityProjectResolver.Normalize(@"C:\M1Work\Public\Excels\Tables\");
        Assert.Equal("c:/m1work/public/excels/tables", n);
    }

    [Fact]
    public void Normalize_UnifiesBackslash()
    {
        var n = UnityProjectResolver.Normalize(@"C:/M1Work\Public/Excels\Tables");
        Assert.Equal("c:/m1work/public/excels/tables", n);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Normalize_EmptyOrWhitespace_ReturnsEmpty(string path)
    {
        Assert.Equal("", UnityProjectResolver.Normalize(path));
    }

    // ── IsUnityProject：Assets + ProjectSettings 双目录校验 ──────────────────

    [Fact]
    public void IsUnityProject_WithAssetsAndProjectSettings_True()
    {
        var root = MkTempRoot("iup_ok_");
        try
        {
            Directory.CreateDirectory(Path.Combine(root, "Assets"));
            Directory.CreateDirectory(Path.Combine(root, "ProjectSettings"));
            Assert.True(UnityProjectResolver.IsUnityProject(root));
        }
        finally
        {
            Cleanup(root);
        }
    }

    [Fact]
    public void IsUnityProject_MissingProjectSettings_False()
    {
        // 只有 Assets 不算 Unity 项目，防误认（比如游戏的 Assets 子目录）
        var root = MkTempRoot("iup_noassets_");
        try
        {
            Directory.CreateDirectory(Path.Combine(root, "Assets"));
            Assert.False(UnityProjectResolver.IsUnityProject(root));
        }
        finally
        {
            Cleanup(root);
        }
    }

    [Fact]
    public void IsUnityProject_Nonexistent_False()
    {
        Assert.False(UnityProjectResolver.IsUnityProject(@"Z:\no\such\dir\at\all"));
    }

    [Fact]
    public void IsUnityProject_Null_False()
    {
        Assert.False(UnityProjectResolver.IsUnityProject(null!));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void IsUnityProject_EmptyOrWhitespace_False(string dir)
    {
        Assert.False(UnityProjectResolver.IsUnityProject(dir));
    }

    // ── FindUnityProjectCandidates：从 Excel 表目录往上找兄弟 Unity 项目 ─────
    // 对应本地 BasePath=C:\M1Work\Public\Excels\Tables → 命中 C:\M1Work\Code 的链路。

    [Fact]
    public void FindUnityProjectCandidates_FindsSiblingUnityProject()
    {
        // root/Excels/Tables（Excel 表目录）+ root/MyCode/{Assets,ProjectSettings}（Unity 项目）+ root/NotUnity（干扰）
        var root = MkTempRoot("upr_find_");
        try
        {
            Directory.CreateDirectory(Path.Combine(root, "Excels", "Tables"));
            Directory.CreateDirectory(Path.Combine(root, "MyCode", "Assets"));
            Directory.CreateDirectory(Path.Combine(root, "MyCode", "ProjectSettings"));
            Directory.CreateDirectory(Path.Combine(root, "NotUnity"));

            var candidates = UnityProjectResolver.FindUnityProjectCandidates(
                Path.Combine(root, "Excels", "Tables")
            );

            Assert.Contains(
                candidates.Select(Path.GetFullPath),
                c => c == Path.GetFullPath(Path.Combine(root, "MyCode"))
            );
            Assert.DoesNotContain(
                candidates.Select(Path.GetFullPath),
                c => c == Path.GetFullPath(Path.Combine(root, "NotUnity"))
            );
        }
        finally
        {
            Cleanup(root);
        }
    }

    static string MkTempRoot(string prefix) =>
        Path.Combine(Path.GetTempPath(), $"ndt_{prefix}{System.Guid.NewGuid():N}");

    static void Cleanup(string root)
    {
        try
        {
            if (Directory.Exists(root))
                Directory.Delete(root, recursive: true);
        }
        catch
        {
            // 测试清理失败不挂测试
        }
    }
}
