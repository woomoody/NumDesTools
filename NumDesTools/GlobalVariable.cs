namespace NumDesTools;

/// <summary>
/// 把一些全局变量生成本地配置，可以自定义修改
/// </summary>
public class GlobalVariable
{
    public  readonly Dictionary<string, string> _defaultValue =
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
            { "DeepSeektApiUrl" , "https://api.deepseek.com/v1/chat/completions"},
            { "DeepSeektApiModel", "deepseek-coder"},
            { "ChatGptSysContentExcelAss", "你是一个代码和办公助手，特别擅长回答Excel的公式以及代码编写，特别擅长C#，打印输出不要使用控制台，使用：Debug.Print，判断需要记录日志，使用：LogDisplay.RecordLine(\r\n \"[{0}] , {1}\",\r\n DateTime.Now.ToString(CultureInfo.InvariantCulture),\r\n$\"{selectedRange.Count}\"\r\n);" },
            { "ChatGptSysContentTransferAss", "你是一个助手，特别擅长多种语言的翻译工作,你的回答中只会输出指定的翻译后的内容，不掺杂其他解释， 根据输入内容中的换行符，作为行的分界线，所需要翻译语言的种类为列的分界线，输出的翻译结果格式为Json的嵌套数组，格式如下：[[\"A语言译文1\",\"A语言译文2\"],[\"B语言译文1\",\"B语言译文2\"]]" }
        };

    public  readonly  string _filePath = Path.Combine(
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
                        Value[key] = value;
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

        ReadValue();

        // 更新指定的键值对
        if (_defaultValue.ContainsKey(key))
            Value[key] = value;

        // 将更新后的 _defaultValue 写回文件
        var updatedLines = new List<string>();
        foreach (var kvp in Value)
            updatedLines.Add($"{kvp.Key} = {kvp.Value}");
        File.WriteAllLines(_filePath, updatedLines);
    }

    public void ReadValue()
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
                        Value[fileKey] = fileValue;
                }
            }
        }
    }
}
