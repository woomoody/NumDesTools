using System.Collections;
using System.Collections.Specialized;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NumDesTools.Config
{
    public class GlobalVariable
    {
        #region 默认值

        // 默认键值对配置
        public readonly Dictionary<string, string> _defaultValue =
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
                { "ApiKey", "" },
                { "ApiUrl", "" },
                { "ApiModel", "" },
                { "ChatGptApiKey", "***" },
                { "ChatGptApiUrl", "https://api.openai.com/v1/chat/completions" },
                { "ChatGptApiModel", "gpt-4o" },
                { "DeepSeektApiKey", "***" },
                { "DeepSeektApiUrl", "https://api.deepseek.com/v1/chat/completions" },
                { "DeepSeektApiModel", "deepseek-coder" },
                {
                    "ChatGptSysContentExcelAss",
                    "你是一个代码和办公助手，特别擅长回答Excel的公式以及代码编写，特别擅长C#，打印输出不要使用控制台，使用：Debug.Print，判断需要记录日志，使用：LogDisplay.RecordLine(\"[{0}] , {1}\", DateTime.Now.ToString(CultureInfo.InvariantCulture),$\"{selectedRange.Count}\");"
                },
                {
                    "ChatGptSysContentTransferAss",
                    "你是一个助手，特别擅长多种语言的翻译工作,你的回答中只会输出指定的翻译后的内容，不掺杂其他解释， 根据输入内容中的换行符，作为行的分界线，所需要翻译语言的种类为列的分界线，输出的翻译结果格式为Json的嵌套数组，格式如下：[[\"A语言译文1\",\"A语言译文2\"],[\"B语言译文1\",\"B语言译文2\"]]"
                }
            };

        // 默认列表配置
        public readonly List<string> _defaultNormaKeyList =
            new() { ",,", "[,", ",]", "{,", ",}", "，，", "[，", "，]", "{，", "，}" };

        public readonly List<string> _defaultSpecialKeyList = new() { "][", "}{" };

        public readonly List<CoupleKey> _defaultCoupleKeyList =
            new() { new CoupleKey("[", "]"), new CoupleKey("{", "}") };

        #endregion

        // 配置文件路径
        public readonly string _filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesGlobalKey.json"
        );

        // 配置数据
        private ConfigData _configData;

        public GlobalVariable()
        {
            ReadOrCreate();
        }

        #region 公共属性

        public Dictionary<string, string> Value => _configData.Value;
        public List<string> NormaKeyList => _configData.NormaKeyList;
        public List<string> SpecialKeyList => _configData.SpecialKeyList;
        public List<CoupleKey> CoupleKeyList => _configData.CoupleKeyList;

        #endregion

        #region 配置加载与保存

        public void ReadOrCreate()
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                var fileValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

                _configData = new ConfigData
                {
                    Value = new Dictionary<string, string>(),
                    NormaKeyList = new List<string>(),
                    SpecialKeyList = new List<string>(),
                    CoupleKeyList = new List<CoupleKey>()
                };

                // 处理键值对配置
                foreach (var kvp in fileValues)
                {
                    if (kvp.Value is List<object> listValue)
                    {
                        _configData.Value[kvp.Key] = string.Join("\n", listValue);
                    }
                    else if (kvp.Value is string stringValue)
                    {
                        _configData.Value[kvp.Key] = stringValue;
                    }
                }

                // 处理列表配置
                if (
                    fileValues.ContainsKey("NormaKeyList")
                    && fileValues["NormaKeyList"] is JToken normaKeyListToken
                )
                {
                    if (normaKeyListToken.Type == JTokenType.Array)
                    {
                        _configData.NormaKeyList = normaKeyListToken.ToObject<List<string>>();
                    }
                }

                if (
                    fileValues.ContainsKey("SpecialKeyList")
                    && fileValues["SpecialKeyList"] is JToken specialKeyListToken
                )
                {
                    if (specialKeyListToken.Type == JTokenType.Array)
                    {
                        _configData.SpecialKeyList = specialKeyListToken.ToObject<List<string>>();
                    }
                }

                if (
                    fileValues.ContainsKey("CoupleKeyList")
                    && fileValues["CoupleKeyList"] is JToken coupleKeyListToken
                )
                {
                    if (coupleKeyListToken.Type == JTokenType.Array)
                    {
                        _configData.CoupleKeyList = coupleKeyListToken
                            .ToObject<List<Dictionary<string, string>>>()
                            .ConvertAll(dict => new CoupleKey(dict["Left"], dict["Right"]));
                    }
                }

                // 合并默认值
                MergeWithDefaults();
            }
            else
            {
                _configData = CreateDefaultConfig();
                SaveConfig();
            }
        }

        private List<T> MergeLists<T>(List<T> original, List<T> defaults)
        {
            var result = new HashSet<T>(original); // 去重
            result.UnionWith(defaults); // 合并
            return result.ToList(); // 转回 List
        }

        private void MergeWithDefaults()
        {
            // 合并键值对配置
            foreach (var kvp in _defaultValue)
            {
                if (!_configData.Value.ContainsKey(kvp.Key))
                {
                    _configData.Value[kvp.Key] = kvp.Value;
                }
            }

            // 合并列表配置
            _configData.NormaKeyList = MergeLists(_configData.NormaKeyList, _defaultNormaKeyList);
            _configData.SpecialKeyList = MergeLists(
                _configData.SpecialKeyList,
                _defaultSpecialKeyList
            );
            _configData.CoupleKeyList = MergeLists(
                _configData.CoupleKeyList,
                _defaultCoupleKeyList
            );
        }

        private ConfigData CreateDefaultConfig()
        {
            return new ConfigData
            {
                Value = new Dictionary<string, string>(_defaultValue),
                NormaKeyList = [.. _defaultNormaKeyList],
                SpecialKeyList = [.. _defaultSpecialKeyList],
                CoupleKeyList = [.. _defaultCoupleKeyList]
            };
        }

        public void SaveValue(string key, string value)
        {
            // 如果文件存在，先读取现有的配置内容
            OrderedDictionary existingConfig = new OrderedDictionary();
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath, Encoding.UTF8);
                var tempDict =
                    JsonConvert.DeserializeObject<Dictionary<string, object>>(json)
                    ?? new Dictionary<string, object>();

                // 将 Dictionary 转换为 OrderedDictionary
                foreach (var kvp in tempDict)
                {
                    existingConfig[kvp.Key] = kvp.Value;
                }
            }

            // 更新或添加新的键值对
            if (existingConfig.Contains(key))
            {
                existingConfig[key] = value;
            }
            else
            {
                existingConfig.Add(key, value);
            }

            // 将 OrderedDictionary 转换为普通 Dictionary 以便序列化
            var orderedDictAsDict = existingConfig
                .Cast<DictionaryEntry>()
                .ToDictionary(entry => (string)entry.Key, entry => entry.Value);

            // 序列化回文件
            var updatedJson = JsonConvert.SerializeObject(orderedDictAsDict, Formatting.Indented);
            File.WriteAllText(_filePath, updatedJson, Encoding.UTF8);
        }

        public void SaveConfig()
        {
            // 创建一个新的有序字典，用于存储格式化后的值
            var formattedValue = new OrderedDictionary();

            foreach (var kvp in _configData.Value)
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
                            lines.Add(
                                kvp.Value.Substring(
                                    i,
                                    Math.Min(NumDesAddIn.MaxLineLength, kvp.Value.Length - i)
                                )
                            );
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

            // 添加列表配置
            formattedValue["NormaKeyList"] = _configData.NormaKeyList;
            formattedValue["SpecialKeyList"] = _configData.SpecialKeyList;
            formattedValue["CoupleKeyList"] = _configData.CoupleKeyList;

            // 将 OrderedDictionary 转换为普通 Dictionary 以便序列化
            var orderedDictAsDict = formattedValue
                .Cast<DictionaryEntry>()
                .ToDictionary(entry => (string)entry.Key, entry => entry.Value);

            // 序列化为 JSON
            var json = JsonConvert.SerializeObject(orderedDictAsDict, Formatting.Indented);

            // 写入文件
            File.WriteAllText(_filePath, json, Encoding.UTF8);
        }

        public void ResetToDefault()
        {
            try
            {
                // 重置配置数据为默认值
                _configData = CreateDefaultConfig();

                // 保存默认配置到文件
                SaveConfig();

                // 提示用户操作成功
                MessageBox.Show(
                    "全局变量已重置为默认值！",
                    "操作成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                // 捕获异常并提示用户
                MessageBox.Show(
                    $"重置失败：{ex.Message}",
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        #endregion

        #region 长文本处理


        // 判断是否为长文本
        private bool IsLongText(string text)
        {
            return text?.Length > NumDesAddIn.LongTextThreshold;
        }

        #endregion

        #region 配置数据结构

        private struct ConfigData
        {
            public Dictionary<string, string> Value { get; set; }
            public List<string> NormaKeyList { get; set; }
            public List<string> SpecialKeyList { get; set; }
            public List<CoupleKey> CoupleKeyList { get; set; }
        }

        #endregion

        #region CoupleKey 结构体

        public struct CoupleKey(string left, string right)
        {
            public string Left { get; set; } = left;
            public string Right { get; set; } = right;

            public void Deconstruct(out string left, out string right)
            {
                left = Left;
                right = Right;
            }
        }

        #endregion
    }
}
