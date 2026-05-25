using Newtonsoft.Json;

namespace NumDesTools;

/// <summary>
/// 独立工具版的 NumDesAddIn / GlobalVariable 占位，仅实现冲突工具依赖的极少接口。
/// </summary>
internal static class NumDesAddIn
{
    public static GlobalVariable GlobalValue { get; } = new();
}

internal sealed class GlobalVariable
{
    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "NumDesTools", "conflict_tool_settings.json"
    );

    private Dictionary<string, string>? _data;

    public Dictionary<string, string> Value
    {
        get
        {
            if (_data != null) return _data;
            try
            {
                if (File.Exists(SettingsPath))
                    _data = JsonConvert.DeserializeObject<Dictionary<string, string>>(
                        File.ReadAllText(SettingsPath)
                    );
            }
            catch { }
            return _data ??= [];
        }
    }

    public void SaveValue(string key, string value)
    {
        Value[key] = value;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonConvert.SerializeObject(Value, Formatting.Indented));
        }
        catch { }
    }
}
