using System;
using System.Collections.Generic;
using System.IO;

namespace NumDesTools;

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

    private readonly string _filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesGlobalKey.txt");

    public Dictionary<string, string> Value { get; set; } = new();

    public GlobalVariable()
    {
        if (File.Exists(_filePath))
        {
            // 如果文件存在，读取文件中的数据
            var lines = File.ReadAllLines(_filePath);

            foreach (var line in lines)
            {
                var parts = line.Split('=');

                if (parts.Length == 2)
                {
                    var key = parts[0].Trim();
                    var value = parts[1].Trim();
                    //Txt不能随意添加变量
                    if (_defaultValue.ContainsKey(key)) Value[key] = value;
                }
            }
        }
        else
        {
            // 如果文件不存在，设置默认值
            Value = _defaultValue;
            // 将默认值写入文件
            var lines = new List<string>();
            foreach (var kvp in _defaultValue) lines.Add($"{kvp.Key} = {kvp.Value}");
            File.WriteAllLines(_filePath, lines);
        }
    }

    public void SaveValue(string key, string value)
    {
        if (File.Exists(_filePath))
        {
            // 检查字典中是否存在要更新的键
            if (_defaultValue.ContainsKey(key))
                // 如果存在，更新它的值
                _defaultValue[key] = value;

            var lines = new List<string>();

            foreach (var kvp in _defaultValue) lines.Add($"{kvp.Key} = {kvp.Value}");
            File.WriteAllLines(_filePath, lines);
        }
    }
}