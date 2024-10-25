using System.ComponentModel;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Windows.Data;
using GraphX.Common.Models;

namespace NumDesTools;

/// <summary>
/// 公共的Excel自定义类
/// </summary>

//自定义Com表格容器类
public class SelfComSheetCollect : INotifyPropertyChanged
{
    private string _name;
    private bool _isHidden;
    private string _detailInfo;
    private Tuple<int, int> _usedRangeSize;

    public string Name
    {
        get => _name;
        set
        {
            if (_name != value)
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }
    }

    public bool IsHidden
    {
        get => _isHidden;
        set
        {
            if (_isHidden != value)
            {
                _isHidden = value;
                OnPropertyChanged(nameof(IsHidden));
            }
        }
    }
    public string DetailInfo
    {
        get => _detailInfo;
        set
        {
            if (_detailInfo != value)
            {
                _detailInfo = value;
                OnPropertyChanged(nameof(DetailInfo));
            }
        }
    }
    public Tuple<int, int> UsedRangeSize
    {
        get => _usedRangeSize;
        set
        {
            if (!Equals(_usedRangeSize, value))
            {
                _usedRangeSize = value;
                OnPropertyChanged(nameof(UsedRangeSize));
            }
        }
    }
    public event PropertyChangedEventHandler PropertyChanged;

    protected virtual void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

//自定义Com工作簿信息容器类
public class SelfWorkBookSearchCollect((string, string, int, string) tuple)
{
    public string FilePath { get; set; } = tuple.Item1;
    public string SheetName { get; set; } = tuple.Item2;
    public int CellRow { get; set; } = tuple.Item3;
    public string CellCol { get; set; } = tuple.Item4;
}

//自定义Com单元格类
public class SelfCellData((string, int, int) tuple)
{
    public string Value { get; set; } = tuple.Item1;
    public int Row { get; set; } = tuple.Item2;
    public int Column { get; set; } = tuple.Item3;

}

//自定义Com表格-单元格类
public class SelfSheetCellData((string, int, int,string , string) tuple)
{
    public string Value { get; set; } = tuple.Item1;
    public int Row { get; set; } = tuple.Item2;
    public int Column { get; set; } = tuple.Item3;
    public string SheetName { get; set; } = tuple.Item4;
    public string Tips { get; set; } = tuple.Item5;
}
//字符串正则转换
public class SelfStringRegexConverter : IValueConverter
{
    public string RegexPattern { get; set; }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var inputString = value as string;
        if (inputString == null)
            return null;

        var regex = new Regex(RegexPattern);
        var match = regex.Match(inputString);

        // 返回第一个匹配的结果
        if (match.Success)
        {
            return match.Value;
        }

        // 如果没有匹配的结果，返回null或者其他适当的值
        return null;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

//自定义GraphX顶点数据类
public class SelfGraphXVertex : VertexBase
{
    public string Name { get; set; }

    public override string ToString()
    {
        return Name;
    }
}

//自定义GraphX边数据类
public class SelfGraphXEdge(SelfGraphXVertex source, SelfGraphXVertex target)
    : EdgeBase<SelfGraphXVertex>(source, target)
{
    public override string ToString()
    {
        return $"{Source.Name} -> {Target.Name}";
    }
}

//自定义获取指定路径Excel文件
public class SelfExcelFileCollector(string rootPath , int pathLevels)
{
    //获取指定路径Excel文件路径
    public string[] GetAllExcelFilesPath()
    {
        var paths = new List<string>
        {
            Path.Combine(GetParentDirectory(rootPath, pathLevels), "Excels", "Tables"),
            Path.Combine(GetParentDirectory(rootPath, pathLevels), "Excels", "Localizations"),
            Path.Combine(GetParentDirectory(rootPath, pathLevels), "Excels", "UIs"),
            Path.Combine(GetParentDirectory(rootPath, pathLevels), "Excels", "Tables", "克朗代克"),
            Path.Combine(GetParentDirectory(rootPath, pathLevels), "Excels", "Tables", "二合")
        };

        var files = paths
            .SelectMany(path => Directory.Exists(path) ? GetExcelFiles(path) : [])
            .Where(file => !Path.GetFileName(file).Contains("~")) // 过滤掉包含 ~ 的文件
            .ToArray();

        return files;
    }
    //获取指定路径Excel文件路径MD5
    public enum KeyMode
    {
        FullPath,//完整路径
        FileNameWithExt,//带扩展名
        FileNameWithoutExt//不带扩展名
    }
    public Dictionary<string, (string FullPath, string FileNameWithExt, string FileNameWithoutExt, string MD5)> GetAllExcelFilesMd5(KeyMode mode)
    {
        var files = GetAllExcelFilesPath();
        var fileMd5Dictionary = new Dictionary<string, (string FullPath, string FileNameWithExt, string FileNameWithoutExt, string MD5)>();

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
                _ => throw new ArgumentOutOfRangeException(nameof(mode), mode, null)
            };

            fileMd5Dictionary[key] = (fullPath, fileNameWithExt, fileNameWithoutExt, md5);
        }

        return fileMd5Dictionary;
    }

    private string GetParentDirectory(string path, int levels)
    {
        for (int i = 0; i < levels; i++)
        {
            path = Path.GetDirectoryName(path);
        }
        return path;
    }
    private static IEnumerable<string> GetExcelFiles(string path)
    {
        return Directory
            .EnumerateFiles(path, "*.xlsx")
            .Where(file => !Path.GetFileName(file).Contains("#"));
    }
    private static string CalculateMd5(string filePath)
    {
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToUpperInvariant();
            }
        }
    }
}