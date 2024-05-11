using System;
using System.Collections.Generic;
using System.IO;

namespace NumDesTools
{
    /// <summary>
    /// 把一些全局变量生成本地配置，可以自定义修改
    /// </summary>
    public class GlobalVariable
    {
        private readonly Dictionary<string, string> _defaultValue = new()
        {
            { "LabelText", "放大镜：关闭" },
            { "FocusLabelText", "聚光灯：关闭" },
            { "LabelTextRoleDataPreview", "角色数据预览：关闭" },
            { "SheetMenuText", "表格目录：开启" },
            { "TempPath", @"\Client\Assets\Resources\Table" }
        };

        private readonly string _filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "NumDesGlobalKey.txt");

        public Dictionary<string, string> Value  { get; set; } = new();

        public  GlobalVariable()
        {

            if (File.Exists(_filePath))
            {
                // 如果文件存在，读取文件中的数据
                string[] lines = File.ReadAllLines(_filePath);

                foreach (string line in lines)
                {
                    string[] parts = line.Split('=');

                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        string value = parts[1].Trim();
                        //Txt不能随意添加变量
                        if (_defaultValue.ContainsKey(key))
                        {
                            Value[key] = value;
                        }
                    }
                }
            }
            else
            {
                // 如果文件不存在，设置默认值
                Value = _defaultValue;
                // 将默认值写入文件
                List<string> lines = new List<string>();
                foreach (KeyValuePair<string, string> kvp in _defaultValue)
                {
                    lines.Add($"{kvp.Key} = {kvp.Value}");
                }
                File.WriteAllLines(_filePath, lines);
            }
        }
    }
}
