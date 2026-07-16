using System.IO;
using System.Linq;
using NumDesTools.ExcelToLua;
using Xunit;

namespace NumDesTools.Tests;

/// <summary>
/// 全表导出功能核心逻辑测试：扫描 3 目录所有 Excel + 清空 Tables 输出目录（保留 NonOutputTable）。
/// 临时目录构造，不依赖真实 M1Work。
/// </summary>
public class FullExportScannerTests
{
    [Fact]
    public void ScanAllExcels_ReturnsAllXlsxFromThreeSubdirs()
    {
        var root = MkTempRoot("fscan_");
        try
        {
            // Excels/Localizations/loc1.xlsx, Excels/Tables/t1.xlsx, Excels/UIs/ui1.xlsx
            // + 一个 .xls 老格式 + 一个非 Excel 文件（应忽略）+ 一个 # 前缀隐藏表（应忽略）
            Directory.CreateDirectory(Path.Combine(root, "Localizations"));
            Directory.CreateDirectory(Path.Combine(root, "Tables"));
            Directory.CreateDirectory(Path.Combine(root, "UIs"));
            File.WriteAllText(Path.Combine(root, "Localizations", "loc1.xlsx"), "x");
            File.WriteAllText(Path.Combine(root, "Tables", "t1.xlsx"), "x");
            File.WriteAllText(Path.Combine(root, "Tables", "t2.xls"), "x");
            File.WriteAllText(Path.Combine(root, "UIs", "ui1.xlsx"), "x");
            File.WriteAllText(Path.Combine(root, "Tables", "readme.txt"), "x"); // 非 Excel，忽略
            Directory.CreateDirectory(Path.Combine(root, "Tables", "#Hidden"));
            File.WriteAllText(Path.Combine(root, "Tables", "#Hidden", "h.xlsx"), "x"); // # 目录，看实现是否扫子目录

            var files = FullExportScanner.ScanAllExcels(root);

            // 至少 4 个 Excel（loc1/t1/t2/ui1）
            Assert.Contains(files, f => Path.GetFileName(f) == "loc1.xlsx");
            Assert.Contains(files, f => Path.GetFileName(f) == "t1.xlsx");
            Assert.Contains(files, f => Path.GetFileName(f) == "t2.xls");
            Assert.Contains(files, f => Path.GetFileName(f) == "ui1.xlsx");
            Assert.DoesNotContain(files, f => Path.GetFileName(f) == "readme.txt");
            // # 前缀隐藏/WIP 表必须被过滤（对齐 GitExportSelectWindow.IsExportable + ExportAll 规则）
            Assert.DoesNotContain(files, f => Path.GetFileName(f) == "h.xlsx");
        }
        finally
        {
            Cleanup(root);
        }
    }

    [Fact]
    public void CleanTablesOutput_DeletesTxt_KeepsMetaAndNonOutputTableSubdir()
    {
        var tables = MkTempRoot("fclean_");
        Directory.CreateDirectory(tables);
        try
        {
            // Tables 下：a.lua.txt, a.lua.txt.meta, b.lua.txt, b.lua.txt.meta + NonOutputTable/c.lua.txt
            File.WriteAllText(Path.Combine(tables, "a.lua.txt"), "x");
            File.WriteAllText(Path.Combine(tables, "a.lua.txt.meta"), "x");
            File.WriteAllText(Path.Combine(tables, "b.lua.txt"), "x");
            File.WriteAllText(Path.Combine(tables, "b.lua.txt.meta"), "x");
            Directory.CreateDirectory(Path.Combine(tables, "NonOutputTable"));
            File.WriteAllText(Path.Combine(tables, "NonOutputTable", "c.lua.txt"), "x");
            File.WriteAllText(Path.Combine(tables, "NonOutputTable", "c.lua.txt.meta"), "x");

            FullExportScanner.CleanTablesOutput(tables);

            // .txt 删了，.meta 保留（防 GUID 重置）
            Assert.False(File.Exists(Path.Combine(tables, "a.lua.txt")));
            Assert.True(File.Exists(Path.Combine(tables, "a.lua.txt.meta")));
            Assert.False(File.Exists(Path.Combine(tables, "b.lua.txt")));
            Assert.True(File.Exists(Path.Combine(tables, "b.lua.txt.meta")));
            // NonOutputTable 子文件夹 + 里面文件保留
            Assert.True(Directory.Exists(Path.Combine(tables, "NonOutputTable")));
            Assert.True(File.Exists(Path.Combine(tables, "NonOutputTable", "c.lua.txt")));
        }
        finally
        {
            Cleanup(root: tables);
        }
    }

    [Fact]
    public void PruneOrphanMetas_DeletesMetaWithoutTxt_KeepsMetaWithTxt()
    {
        var tables = MkTempRoot("fprune_");
        Directory.CreateDirectory(tables);
        try
        {
            // 活跃表：a.lua.txt + a.lua.txt.meta（txt 在，meta 保留）
            File.WriteAllText(Path.Combine(tables, "a.lua.txt"), "x");
            File.WriteAllText(Path.Combine(tables, "a.lua.txt.meta"), "x");
            // 死表：b.lua.txt.meta（txt 不在，meta 孤儿 → 删）
            File.WriteAllText(Path.Combine(tables, "b.lua.txt.meta"), "x");

            FullExportScanner.PruneOrphanMetas(tables);

            // 活跃表 meta 保留
            Assert.True(File.Exists(Path.Combine(tables, "a.lua.txt.meta")));
            // 死表 meta 删了
            Assert.False(File.Exists(Path.Combine(tables, "b.lua.txt.meta")));
        }
        finally
        {
            Cleanup(root: tables);
        }
    }

    static string MkTempRoot(string prefix) =>
        Path.Combine(Path.GetTempPath(), $"ndt_fe_{prefix}{System.Guid.NewGuid():N}");

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
