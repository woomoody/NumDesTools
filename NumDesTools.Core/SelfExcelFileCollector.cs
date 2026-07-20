using System.Security.Cryptography;

namespace NumDesTools;

public class SelfExcelFileCollector(string currentPath)
{
    //获取指定路径Excel文件路径
    public string[] GetAllExcelFilesPath()
    {
        var rootPath = FindRootDirectory(currentPath, "Excels");
        if (rootPath == null)
            return [];

        var files = GetExcelFiles(rootPath)
            .Where(file => !Path.GetFileName(file).Contains("~"))
            .ToArray();

        return files;
    }

    //获取根目录
    private static string? FindRootDirectory(string rootPath, string rootFolderName)
    {
        DirectoryInfo? dirInfo = new DirectoryInfo(rootPath);

        while (dirInfo != null && dirInfo.Name != rootFolderName)
        {
            dirInfo = dirInfo.Parent;
        }

        return dirInfo?.FullName;
    }

    //获取指定路径Excel文件路径MD5
    public enum KeyMode
    {
        FullPath, //完整路径
        FileNameWithExt, //带扩展名
        FileNameWithoutExt, //不带扩展名
    }

    public Dictionary<
        string,
        (string FullPath, string FileNameWithExt, string FileNameWithoutExt, string MD5)
    > GetAllExcelFilesMd5(KeyMode mode)
    {
        var files = GetAllExcelFilesPath();
        var fileMd5Dictionary =
            new Dictionary<
                string,
                (string FullPath, string FileNameWithExt, string FileNameWithoutExt, string MD5)
            >();

        foreach (var file in files)
        {
            string fullPath = file;
            string fileNameWithExt = Path.GetFileName(file);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(file);
            string md5 = CalculateMd5(file);

            string key = mode switch
            {
                KeyMode.FullPath => fullPath,
                KeyMode.FileNameWithExt => fileNameWithExt,
                KeyMode.FileNameWithoutExt => fileNameWithoutExt,
                _ => throw new ArgumentOutOfRangeException(nameof(mode), mode, null),
            };

            fileMd5Dictionary[key] = (fullPath, fileNameWithExt, fileNameWithoutExt, md5);
        }

        return fileMd5Dictionary;
    }

    // 不扫描这些目录（工具/基础设施，非游戏配置表）
    private static readonly HashSet<string> _excludeDirs = new(StringComparer.OrdinalIgnoreCase)
    {
        "TablesTools",
        "Networks",
    };

    private static IEnumerable<string> GetExcelFiles(string path)
    {
        return Directory
            .EnumerateFiles(path, "*.xlsx", SearchOption.AllDirectories)
            .Where(file =>
                !Path.GetFileName(file).Contains("#") &&
                !_excludeDirs.Contains(new DirectoryInfo(Path.GetDirectoryName(file)!).Name) &&
                !file.Split(Path.DirectorySeparatorChar).Any(seg => _excludeDirs.Contains(seg)));
    }

    private static string CalculateMd5(string filePath)
    {
        using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite
        );
        using var md5 = MD5.Create();
        var hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();
    }
}
