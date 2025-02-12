using System.ComponentModel;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Data;
using GraphX.Common.Models;
using Microsoft.Data.Sqlite;
using Newtonsoft.Json;
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
public class SelfSheetCellData((string, int, int, string, string) tuple)
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
public class SelfExcelFileCollector(string currentPath)
{
    //获取指定路径Excel文件路径
    public string[] GetAllExcelFilesPath()
    {
        var rootPath = FindRootDirectory(currentPath, "Excels");
        var paths = new List<string>
        {
            Path.Combine(GetParentDirectory(rootPath, 0), "Tables"),
            Path.Combine(GetParentDirectory(rootPath, 0), "Localizations"),
            Path.Combine(GetParentDirectory(rootPath, 0), "UIs"),
            Path.Combine(GetParentDirectory(rootPath, 0), "Tables", "克朗代克"),
            Path.Combine(GetParentDirectory(rootPath, 0), "Tables", "二合")
        };

        var files = paths
            .SelectMany(path => Directory.Exists(path) ? GetExcelFiles(path) : [])
            .Where(file => !Path.GetFileName(file).Contains("~")) // 过滤掉包含 ~ 的文件
            .ToArray();

        return files;
    }

    //获取根目录
    private static string FindRootDirectory(string rootPath, string rootFolderName)
    {
        DirectoryInfo dirInfo = new DirectoryInfo(rootPath);

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
        FileNameWithoutExt //不带扩展名
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

//自定义获取单元格像素坐标
public class SelfGetRangePixels
{
    public static void GetRangePixels() { }
}

//自定义ChatApi
public class ChatApiClient
{
    public static async Task<string> CallApiAsync(object requestBody, string apiKey, string apiUrl)
    {
        if (string.IsNullOrEmpty(apiKey))
        {
            throw new ArgumentException("API 密钥不能为空。");
        }

        using HttpClient client = new HttpClient()
        {
            Timeout = TimeSpan.FromMinutes(5) // 设置超时时间为5分钟
        }; ;

        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

        string jsonBody = JsonConvert.SerializeObject(requestBody);
        var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

        HttpResponseMessage response = await client.PostAsync(apiUrl, content);

        if (response.IsSuccessStatusCode)
        {
            string responseContent = await response.Content.ReadAsStringAsync();
            dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);

            if (
                jsonResponse != null
                && (jsonResponse.choices == null || jsonResponse.choices.Count == 0)
            )
            {
                throw new Exception("API 响应中没有返回有效的 choices 数据。");
            }

            var reponseThink = jsonResponse?.choices[0].message.reasoning_content.ToString();
            var reponseResult = jsonResponse?.choices[0].message.content.ToString();
            string reponseText = "[思考]\n" + reponseThink + "\n[思考]\n" + reponseResult;

            return reponseText;
        }

        string errorContent = await response.Content.ReadAsStringAsync();
        throw new Exception($"API 调用失败，状态码：{response.StatusCode}，错误信息：{errorContent}");
    }
    /// <summary>
    /// 流式调用 API，逐块读取返回的数据，并通过 onChunkReceived 回调实时返回解析后的文本
    /// </summary>
    public static async Task CallApiStreamAsync(object requestBody, string apiKey, string apiUrl,
        Action<string> onChunkReceived, string allText,
        Action onStreamCompleted = null )
    {
        if (string.IsNullOrEmpty(apiKey))
        {
            throw new ArgumentException("API 密钥不能为空。");
        }

        using HttpClient client = new HttpClient();
        client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

        string jsonBody = JsonConvert.SerializeObject(requestBody);
        var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

        // 设置 HttpCompletionOption.ResponseHeadersRead 开启流式读取
        using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, apiUrl)
        {
            Content = content
        };
        HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        if (!response.IsSuccessStatusCode)
        {
            string errorContent = await response.Content.ReadAsStringAsync();
            throw new Exception($"API 调用失败，状态码：{response.StatusCode}，错误信息：{errorContent}");
        }

        using var stream = await response.Content.ReadAsStreamAsync();
        using var reader = new StreamReader(stream, Encoding.UTF8);

        while (!reader.EndOfStream)
        {
            string line = await reader.ReadLineAsync();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // 流式 API 通常每行以 "data: " 开头
            if (line.StartsWith("data: "))
            {
                string jsonPart = line.Substring("data: ".Length).Trim();
                // 遇到结束标记则退出循环
                if (jsonPart == "[DONE]")
                {
                    break;
                }
                try
                {
                    // 此处假设返回的 JSON 格式为：{ "choices": [ { "delta": { "content": "..." } } ] }
                    dynamic jsonChunk = JsonConvert.DeserializeObject(jsonPart);

                    string reponseThink = jsonChunk?.choices[0].delta.reasoning_content.ToString();
   
                    if (reponseThink == "")
                    {
                        reponseThink = jsonChunk?.choices[0].delta.content.ToString();
                    }

                    onChunkReceived(reponseThink);
                    allText += reponseThink;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"onChunkReceived [解析出错：{ex.Message}]");
                }
            }
        }
        onStreamCompleted?.Invoke();
    }
}

//Chat聊天记录存取
public class ChatMessage
{
    public string Role { get; set; } // 用户或系统
    public string Message { get; set; } // 消息内容
    public bool IsUser { get; set; } // 是否是用户消息
    public DateTime Timestamp { get; set; } // 消息时间戳
    public bool IsStreaming { get; set; }    // 新增字段标识流式消息
}

public class ChatHistoryManager
{
    private readonly string _connectionString =
        $"Data Source={Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ChatHistory.db")}";

    public ChatHistoryManager()
    {
        // 初始化数据库和表
        using (var connection = new SqliteConnection(_connectionString))
        {
            connection.Open();
            var command = connection.CreateCommand();
            command.CommandText =
                @"
                CREATE TABLE IF NOT EXISTS ChatHistory (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Role TEXT NOT NULL,
                    Message TEXT NOT NULL,
                    IsUser INTEGER NOT NULL,
                    Timestamp DATETIME NOT NULL
                )";
            command.ExecuteNonQuery();
        }
    }

    // 保存聊天记录
    public void SaveChatMessage(ChatMessage message)
    {
        using (var connection = new SqliteConnection(_connectionString))
        {
            connection.Open();
            var command = connection.CreateCommand();
            command.CommandText =
                @"
                INSERT INTO ChatHistory (Role, Message, IsUser, Timestamp)
                VALUES (@Role, @Message, @IsUser, @Timestamp)";
            command.Parameters.AddWithValue("@Role", message.Role);
            command.Parameters.AddWithValue("@Message", message.Message);
            command.Parameters.AddWithValue("@IsUser", message.IsUser ? 1 : 0);
            command.Parameters.AddWithValue("@Timestamp", message.Timestamp);
            command.ExecuteNonQuery();
        }
    }

    // 读取聊天记录
    public List<ChatMessage> LoadChatHistory()
    {
        var chatHistory = new List<ChatMessage>();
        using (var connection = new SqliteConnection(_connectionString))
        {
            connection.Open();
            var command = connection.CreateCommand();
            command.CommandText =
                "SELECT Role, Message, IsUser, Timestamp FROM ChatHistory ORDER BY Timestamp ASC";

            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    chatHistory.Add(
                        new ChatMessage
                        {
                            Role = reader.GetString(0),
                            Message = reader.GetString(1),
                            IsUser = reader.GetInt32(2) == 1,
                            Timestamp = reader.GetDateTime(3)
                        }
                    );
                }
            }
        }
        return chatHistory;
    }
}

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
                CreateNoWindow = true
            };

            using (Process process = Process.Start(psi))
            {
                if (process == null)
                    return false;

                using (StreamReader reader = process.StandardOutput)
                {
                    string output = reader.ReadToEnd();
                    process.WaitForExit();

                    // 检查输出中是否包含指定版本
                    return output.Contains($"{versionName} {version}");
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"检查 .NET 版本时发生错误: {ex.Message}");
            Debug.Print($"检查 .NET 版本时发生错误: {ex.Message}");
            return false;
        }
    }

    public static void Install(string installerPath)
    {
        try
        {
            if (!File.Exists(installerPath))
            {
                Debug.Print("安装程序路径无效，请检查路径是否正确。");
                return;
            }

            // 调用安装程序
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = "/quiet /norestart", // 安装程序参数（根据需要修改）
                UseShellExecute = true // 使用 Shell 执行
            };

            using (Process process = Process.Start(psi))
            {
                if (process != null)
                {
                    process.WaitForExit();
                    MessageBox.Show("安装程序已执行完成。");
                    Debug.Print("安装程序已执行完成。");
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"执行安装程序时发生错误: {ex.Message}");
            Debug.Print($"执行安装程序时发生错误: {ex.Message}");
        }
    }
}
