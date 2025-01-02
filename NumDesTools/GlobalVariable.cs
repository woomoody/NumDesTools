namespace NumDesTools;

/// <summary>
/// 把一些全局变量生成本地配置，可以自定义修改
/// </summary>
public class GlobalVariable
{
    private readonly Dictionary<string, string> _defaultValue =
        new()
        {
            { "LabelText", "放大镜：关闭" },
            { "FocusLabelText", "聚光灯：关闭" },
            { "LabelTextRoleDataPreview", "角色数据预览：关闭" },
            { "SheetMenuText", "表格目录：关闭" },
            { "TempPath", @"\Client\Assets\Resources\Table" },
            { "CellHiLightText", "高亮单元格：关闭" },
            { "CheckSheetValueText", "数据自检：开启" },
            { "ShowDnaLogText", "插件日志：关闭" },
            { "ShowAIText", "AI对话：关闭" },
            { "ApiKey" , ""},
            { "ApiUrl" , ""},
            { "ApiModel" , ""},
            { "ChatGptApiKey", "***" },
            { "ChatGptApiUrl", "https://api.openai.com/v1/chat/completions"},
            { "ChatGptApiModel", "gpt-4o"},
            { "DeepSeektApiKey" , "***"},
            { "DeepSeektApiUrl" , "https://api.deepseek.com/v1"},
            { "DeepSeektApiModel", "deepseek-chat"},
            { "ChatGptSysContentExcelAss", "你是一个助手，特别擅长回答Excel的各项功能" },
            { "ChatGptSysContentTransferAss", "你是一个助手，特别擅长多种语言的翻译工作，你的回答中只会输出指定的翻译后的内容，不掺杂别的解释，" +
                                              "输入文本以【#cent#】为标识符区分文本的键值" +
                                              "输出文本需要根据所需翻译的语言种类作为键值的不同值" +
                                              "输入的内容需要遵循json格式"
            }
        };

    private readonly string _filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesGlobalKey.txt"
    );

    public Dictionary<string, string> Value { get; set; } = new();

    public GlobalVariable()
    {
        bool defaultValueUpdate = false;
        //文件存在，则读取文件内容
        if (File.Exists(_filePath))
        {
            var lines = File.ReadAllLines(_filePath);

            foreach (var line in lines)
            {
                var parts = line.Split('=');

                if (parts.Length == 2)
                {
                    var key = parts[0].Trim();
                    var value = parts[1].Trim();
                    //验证_defaultValue和文件内容的key是否一致
                    if (_defaultValue.ContainsKey(key))
                    {
                        //共有的Key以文件内容为准
                        _defaultValue[key] = value;
                    }
                    else
                    {
                        // 文件内容中有新的 key，忽略
                        continue;
                    }
                }
            }
            // 检查 _defaultValue 中是否有新的 key
            foreach (var kvp in _defaultValue)
            {
                if (!Value.ContainsKey(kvp.Key))
                {
                    Value[kvp.Key] = kvp.Value;
                    defaultValueUpdate = true;
                }
            }

            // 如果有新的默认值，更新文件内容
            if (defaultValueUpdate)
            {
                var linesToWrite = new List<string>();
                foreach (var kvp in Value)
                {
                    linesToWrite.Add($"{kvp.Key} = {kvp.Value}");
                }
                File.WriteAllLines(_filePath, linesToWrite);
            }
        }
        //不存在则创建文件配置，写入默认值
        else
        {
            Value = new Dictionary<string, string>(_defaultValue);
            var lines = new List<string>();
            foreach (var kvp in _defaultValue)
                lines.Add($"{kvp.Key} = {kvp.Value}");
            File.WriteAllLines(_filePath, lines);
        }
    }

    public void SaveValue(string key, string value)
    {
        // 读取文件中的数据并更新 _defaultValue
        if (File.Exists(_filePath))
        {
            var lines = File.ReadAllLines(_filePath);
            foreach (var line in lines)
            {
                var parts = line.Split('=');
                if (parts.Length == 2)
                {
                    var fileKey = parts[0].Trim();
                    var fileValue = parts[1].Trim();
                    if (_defaultValue.ContainsKey(fileKey))
                        _defaultValue[fileKey] = fileValue;
                }
            }
        }

        // 更新指定的键值对
        if (_defaultValue.ContainsKey(key))
            _defaultValue[key] = value;

        // 将更新后的 _defaultValue 写回文件
        var updatedLines = new List<string>();
        foreach (var kvp in _defaultValue)
            updatedLines.Add($"{kvp.Key} = {kvp.Value}");
        File.WriteAllLines(_filePath, updatedLines);
    }
}
