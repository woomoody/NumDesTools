using System.ComponentModel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Data;
using GraphX.Common.Models;
using Microsoft.Data.Sqlite;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Action = System.Action;

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
public class SelfSheetCellData
{
    public string SheetName { get; }
    public int Row { get; }
    public int Column { get; }
    public string Value { get; }
    public string Tips { get; }
    public string FilePath { get; }

    public SelfSheetCellData((string, int, int, string, string) tuple)
    {
        Value = tuple.Item1;
        Row = tuple.Item2;
        Column = tuple.Item3;
        SheetName = tuple.Item4;
        Tips = tuple.Item5;
    }

    public SelfSheetCellData((string, int, int, string, string, string) tuple)
    {
        Value = tuple.Item1;
        Row = tuple.Item2;
        Column = tuple.Item3;
        SheetName = tuple.Item4;
        Tips = tuple.Item5;
        FilePath = tuple.Item6;
    }
}

//字符串正则转换
public class SelfStringRegexConverter : IValueConverter
{
    private string _cachedPattern;
    private Regex _cachedRegex;

    public string RegexPattern { get; set; }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var inputString = value as string;
        if (inputString == null)
            return null;

        if (_cachedRegex == null || _cachedPattern != RegexPattern)
        {
            _cachedPattern = RegexPattern;
            _cachedRegex = new Regex(RegexPattern);
        }
        var match = _cachedRegex.Match(inputString);

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

// SelfExcelFileCollector 已迁移至 NumDesTools.Core（跨平台纯逻辑，无 WPF/COM 依赖）

//自定义获取单元格像素坐标
public class SelfGetRangePixels
{
    public static void GetRangePixels() { }
}

// ChatApiClient / ChatMessage / ChatHistoryManager 已迁移至 NumDesTools.AI 命名空间，
// 此处保留转发壳供旧调用方零改动过渡。
public static class ChatApiClient
{
    private static readonly NumDesTools.AI.LiteLlmClient _impl = new();

    public static Task<string> CallApiAsync(
        string model,
        string systemContent,
        string userContent,
        string apiKey,
        string apiUrl
    ) => _impl.CallAsync(model, systemContent, userContent, apiKey, apiUrl);

    public static Task CallApiStreamAsync(
        string model,
        IReadOnlyList<object> messages,
        string apiKey,
        string apiUrl,
        Action<string> onChunkReceived,
        Action onCompleted = null,
        CancellationToken ct = default
    ) => _impl.CallStreamAsync(model, messages, apiKey, apiUrl, onChunkReceived, onCompleted, ct);

    public static Task<List<string>> FetchModelsAsync(string apiKey, string apiUrl) =>
        _impl.FetchModelsAsync(apiKey, apiUrl);
}

public class ChatMessage : NumDesTools.AI.ChatMessage { }

public class ChatHistoryManager : NumDesTools.AI.ChatHistoryManager { }

//自动检测运行环境
public class SelfEnvironmentDetector
{
    public static bool IsInstalled(
        string version,
        string versionName,
        string fileName,
        string arguments
    )
    {
        try
        {
            // 调用 dotnet --list-runtimes 命令
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };

            using Process process = Process.Start(psi);
            if (process == null)
                return false;

            using StreamReader reader = process.StandardOutput;
            string output = reader.ReadToEnd();
            process.WaitForExit();

            // 检查输出中是否包含指定版本
            return output.Contains($"{versionName} {version}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"检查 .NET 版本时发生错误: {ex.Message}");
            PluginLog.Write($"检查 .NET 版本时发生错误: {ex.Message}");
            return false;
        }
    }

    public static void Install(string installerPath)
    {
        try
        {
            if (!File.Exists(installerPath))
            {
                PluginLog.Write("安装程序路径无效，请检查路径是否正确。");
                return;
            }

            // 调用安装程序
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = "/quiet /norestart", // 安装程序参数（根据需要修改）
                UseShellExecute = true, // 使用 Shell 执行
            };

            using Process process = Process.Start(psi);
            if (process != null)
            {
                process.WaitForExit();
                MessageBox.Show("安装程序已执行完成。");
                PluginLog.Write("安装程序已执行完成。");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"执行安装程序时发生错误: {ex.Message}");
            PluginLog.Write($"执行安装程序时发生错误: {ex.Message}");
        }
    }
}
