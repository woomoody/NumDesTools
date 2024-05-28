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
            { "SheetMenuText", "表格目录：开启" },
            { "TempPath", @"\Client\Assets\Resources\Table" }
        };

    private readonly string _filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesGlobalKey.txt"
    );

    public Dictionary<string, string> Value { get; set; } = new();

    public GlobalVariable()
    {
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
                        Value[key] = value;
                }
            }
        }
        else
        {
            Value = _defaultValue;
            var lines = new List<string>();
            foreach (var kvp in _defaultValue)
                lines.Add($"{kvp.Key} = {kvp.Value}");
            File.WriteAllLines(_filePath, lines);
        }
    }

    public void SaveValue(string key, string value)
    {
        if (File.Exists(_filePath))
        {
            if (_defaultValue.ContainsKey(key))
                _defaultValue[key] = value;

            var lines = new List<string>();

            foreach (var kvp in _defaultValue)
                lines.Add($"{kvp.Key} = {kvp.Value}");
            File.WriteAllLines(_filePath, lines);
        }
    }
}
