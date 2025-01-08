using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;

namespace NumDesTools
{
    public class GlobalVariable
    {
        public readonly Dictionary<string, string> _defaultValue = new()
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
            { "ApiKey", "" },
            { "ApiUrl", "" },
            { "ApiModel", "" },
            { "ChatGptApiKey", "***" },
            { "ChatGptApiUrl", "https://api.openai.com/v1/chat/completions" },
            { "ChatGptApiModel", "gpt-4o" },
            { "DeepSeektApiKey", "***" },
            { "DeepSeektApiUrl", "https://api.deepseek.com/v1/chat/completions" },
            { "DeepSeektApiModel", "deepseek-coder" },
            { "ChatGptSysContentExcelAss", "你是一个代码和办公助手，特别擅长回答Excel的公式以及代码编写，特别擅长C#，打印输出不要使用控制台，使用：Debug.Print，判断需要记录日志，使用：LogDisplay.RecordLine(\"[{0}] , {1}\", DateTime.Now.ToString(CultureInfo.InvariantCulture),$\"{selectedRange.Count}\");" },
            { "ChatGptSysContentTransferAss", "你是一个助手，特别擅长多种语言的翻译工作,你的回答中只会输出指定的翻译后的内容，不掺杂其他解释， 根据输入内容中的换行符，作为行的分界线，所需要翻译语言的种类为列的分界线，输出的翻译结果格式为Json的嵌套数组，格式如下：[[\"A语言译文1\",\"A语言译文2\"],[\"B语言译文1\",\"B语言译文2\"]]" }
        };

        public readonly string _filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesGlobalKey.json"
        );

        public Dictionary<string, string> Value { get; set; } = new();

        public GlobalVariable()
        {
            bool defaultValueUpdate = false;

            // 如果文件存在，则读取文件内容
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                var fileValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

                foreach (var kvp in fileValues)
                {
                    if (_defaultValue.ContainsKey(kvp.Key))
                    {
                        // 如果值是数组，则拼接为字符串
                        if (kvp.Value is List<object> listValue)
                        {
                            Value[kvp.Key] = string.Join("\n", listValue);
                        }
                        else
                        {
                            Value[kvp.Key] = kvp.Value?.ToString() ?? string.Empty;
                        }
                    }
                }

                // 检查是否有新的默认值
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
                    SaveToFile();
            }
            else
            {
                // 如果文件不存在，则创建文件并写入默认值
                Value = new Dictionary<string, string>(_defaultValue);
                SaveToFile();
            }
        }

        public void SaveValue(string key, string value)
        {
            if (_defaultValue.ContainsKey(key))
            {
                Value[key] = value;
                SaveToFile();
            }
        }

        public void ReadValue()
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                var fileValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

                foreach (var kvp in fileValues)
                {
                    if (_defaultValue.ContainsKey(kvp.Key))
                    {
                        // 如果值是数组，则拼接为字符串
                        if (kvp.Value is List<object> listValue)
                        {
                            Value[kvp.Key] = string.Join("\n", listValue);
                        }
                        else
                        {
                            Value[kvp.Key] = kvp.Value?.ToString() ?? string.Empty;
                        }
                    }
                }
            }
        }

        private void SaveToFile()
        {
            // 创建一个新的字典，用于存储格式化后的值
            var formattedValue = new Dictionary<string, object>();

            foreach (var kvp in Value)
            {
                // 判断是否为长文本
                if (IsLongText(kvp.Value))
                {
        
                    if (kvp.Value.Contains("\n"))
                    {
                        // 如果是长文本，按行拆分为数组
                        formattedValue[kvp.Key] = kvp.Value.Split('\n');
                    }
                    else
                    {
                        var lines = new List<string>();
                        for (int i = 0; i < kvp.Value.Length; i += NumDesAddIn.MaxLineLength)
                        {
                            lines.Add(kvp.Value.Substring(i, Math.Min(NumDesAddIn.MaxLineLength, kvp.Value.Length - i)));
                        }
                        formattedValue[kvp.Key] = lines;
                    }
                }
                else
                {
                    // 如果不是长文本，直接存储
                    formattedValue[kvp.Key] = kvp.Value;
                }
            }

            // 序列化为 JSON
            var json = JsonConvert.SerializeObject(formattedValue, Formatting.Indented);

            // 写入文件
            File.WriteAllText(_filePath, json, Encoding.UTF8);
        }

        // 判断是否为长文本
        private bool IsLongText(string text)
        {
            return text?.Length > NumDesAddIn.LongTextThreshold;
        }
    }
}