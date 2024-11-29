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
            { "CheckSheetValueText", "数据自检：开启" }
        };

    private readonly string _filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesGlobalKey.txt"
    );

    public Dictionary<string, string> Value { get; set; } = new();

    public GlobalVariable()
    {
        bool fileUpdated = false;

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
                    if (_defaultValue.ContainsKey(key))
                    {
                        Value[key] = value;
                    }
                    else
                    {
                        // 如果文件中有不在 _defaultValue 中的键值对，添加到 _defaultValue
                        _defaultValue[key] = value;
                        fileUpdated = true;
                    }
                }
            }

            // 检查 _defaultValue 中是否有新的键值对
            foreach (var kvp in _defaultValue)
            {
                if (!Value.ContainsKey(kvp.Key))
                {
                    Value[kvp.Key] = kvp.Value;
                    fileUpdated = true;
                }
            }

            // 如果文件内容与 _defaultValue 不一致，更新文件
            if (fileUpdated)
            {
                var updatedLines = new List<string>();
                foreach (var kvp in _defaultValue)
                    updatedLines.Add($"{kvp.Key} = {kvp.Value}");
                File.WriteAllLines(_filePath, updatedLines);
            }
        }
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
