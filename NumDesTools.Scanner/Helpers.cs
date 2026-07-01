using Newtonsoft.Json;

namespace NumDesTools.Scanner;

internal static class Helpers
{
    internal static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "NumDesTools",
        "Config"
    );
    internal const string ActivityXlsx =
        @"C:\M1Work\public\Excels\Tables\ActivityClientData.xlsx";
    internal static readonly string RulesPath = Path.Combine(ConfigDir, "ActivityTableRules.json");
    internal static readonly string FeishuConfigPath = Path.Combine(
        ConfigDir,
        "feishu_config.json"
    );
    internal static readonly string WrittenItemsPath = Path.Combine(
        ConfigDir,
        "written_items.json"
    );
    internal static readonly string NoPermItemsPath = Path.Combine(
        ConfigDir,
        "no_permission_items.json"
    );
    internal const string TablesDir = @"C:\M1Work\Public\Excels\Tables";
    internal static readonly int ConfirmTimeoutSec = 30 * 60;

    internal static HashSet<string> LoadSet(string path)
    {
        if (!File.Exists(path))
            return [];
        try
        {
            var dict =
                JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(path))
                ?? [];
            return dict.Keys.ToHashSet();
        }
        catch
        {
            return [];
        }
    }

    internal static void SaveSet(string path, string itemId, bool overwriteExisting = true)
    {
        Dictionary<string, string> dict = [];
        if (File.Exists(path))
            try
            {
                dict =
                    JsonConvert.DeserializeObject<Dictionary<string, string>>(
                        File.ReadAllText(path)
                    ) ?? [];
            }
            catch { }
        if (!overwriteExisting && dict.ContainsKey(itemId))
            return;
        dict[itemId] = DateTime.Today.ToString("yyyy-MM-dd");
        File.WriteAllText(path, JsonConvert.SerializeObject(dict, Formatting.Indented));
    }

    internal static (string mcpToken, string projectKey) LoadFeishuConfig()
    {
        var token = Environment.GetEnvironmentVariable("FEISHU_MCP_TOKEN");
        var key = Environment.GetEnvironmentVariable("FEISHU_PROJECT_KEY");

        if (string.IsNullOrEmpty(token) || string.IsNullOrEmpty(key))
        {
            if (File.Exists(FeishuConfigPath))
            {
                var cfg =
                    JsonConvert.DeserializeObject<Dictionary<string, string>>(
                        File.ReadAllText(FeishuConfigPath)
                    ) ?? [];
                token ??= cfg.GetValueOrDefault("McpToken");
                key ??= cfg.GetValueOrDefault("ProjectKey");
            }
        }

        if (string.IsNullOrEmpty(token) || string.IsNullOrEmpty(key))
        {
            Console.Error.WriteLine(
                $"[ERROR] 飞书配置缺失。请设置环境变量 FEISHU_MCP_TOKEN / FEISHU_PROJECT_KEY，"
                    + $"或在 {FeishuConfigPath} 中提供 {{\"McpToken\":\"...\",\"ProjectKey\":\"...\"}}。"
            );
            Environment.Exit(1);
        }

        return (token!, key!);
    }

    internal static bool IsPermissionError(Exception ex) =>
        ex.Message.Contains("1000052092") || ex.Message.Contains("无权编辑");

    internal static ActivityTableRules LoadRules()
    {
        if (!File.Exists(RulesPath))
        {
            Console.WriteLine($"[WARN] 规则文件不存在：{RulesPath}");
            return new ActivityTableRules();
        }
        return JsonConvert.DeserializeObject<ActivityTableRules>(File.ReadAllText(RulesPath))
            ?? new ActivityTableRules();
    }

    internal static string Truncate(string s, int max) => s.Length <= max ? s : s[..max] + "…";

    internal static string Now() => DateTime.Now.ToString("HH:mm:ss");
}
