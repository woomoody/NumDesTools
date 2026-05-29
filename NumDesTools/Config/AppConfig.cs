using NumDesTools.Config;

namespace NumDesTools;

/// <summary>
/// 强类型配置入口。内部读写同一份 GlobalVariable 字典，JSON 格式保持平铺兼容。
/// 双轨并行期间，旧代码仍可通过 NumDesAddIn 静态字段访问，新代码走此类。
/// </summary>
public class AppConfig(GlobalVariable store)
{
    public LlmConfig Llm { get; } = new(store);
    public UiConfig Ui { get; } = new(store);
    public GitConfig Git { get; } = new(store);
    public PathConfig Paths { get; } = new(store);
    public AiPromptConfig AiPrompts { get; } = new(store);

    /// <summary>同步保存当前所有字段到 JSON 文件（替代 SaveValue 逐字段写）。</summary>
    public void Save() => store.SaveConfig();

    /// <summary>保存单个键（兼容旧调用点，逐步迁移用）。</summary>
    public void Save(string key, string value) => store.SaveValue(key, value);
}

public class LlmConfig(GlobalVariable store)
{
    public string ApiKey
    {
        get => store.Value.GetValueOrDefault("LiteLLMApiKey", "");
        set => store.Value["LiteLLMApiKey"] = value;
    }

    public string ApiUrl
    {
        get =>
            store.Value.GetValueOrDefault(
                "LiteLLMApiUrl",
                "https://litellm.solotopia.net/v1/chat/completions"
            );
        set => store.Value["LiteLLMApiUrl"] = value;
    }

    public string Model
    {
        get => store.Value.GetValueOrDefault("LiteLLMModel", "global.anthropic.claude-opus-4-7");
        set => store.Value["LiteLLMModel"] = value;
    }

    public List<string> ModelList
    {
        get =>
            store
                .Value.GetValueOrDefault("LiteLLMModelList", "")
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .ToList();
        set => store.Value["LiteLLMModelList"] = string.Join(",", value);
    }

    public string ChatCompletionsUrl =>
        ApiUrl.EndsWith("/chat/completions") ? ApiUrl : ApiUrl.TrimEnd('/') + "/chat/completions";
}

public class UiConfig(GlobalVariable store)
{
    public string LabelText
    {
        get => store.Value.GetValueOrDefault("LabelText", "放大镜：关闭");
        set => store.Value["LabelText"] = value;
    }

    public string FocusLabelText
    {
        get => store.Value.GetValueOrDefault("FocusLabelText", "聚光灯：关闭");
        set => store.Value["FocusLabelText"] = value;
    }

    public string LabelTextRoleDataPreview
    {
        get => store.Value.GetValueOrDefault("LabelTextRoleDataPreview", "角色数据预览：关闭");
        set => store.Value["LabelTextRoleDataPreview"] = value;
    }

    public string SheetMenuText
    {
        get => store.Value.GetValueOrDefault("SheetMenuText", "表格目录：关闭");
        set => store.Value["SheetMenuText"] = value;
    }

    public string CellHighlightText
    {
        get => store.Value.GetValueOrDefault("CellHiLightText", "高亮单元格：关闭");
        set => store.Value["CellHiLightText"] = value;
    }

    public string CheckSheetValueText
    {
        get => store.Value.GetValueOrDefault("CheckSheetValueText", "数据自检：开启");
        set => store.Value["CheckSheetValueText"] = value;
    }

    public string ShowDnaLogText
    {
        get => store.Value.GetValueOrDefault("ShowDnaLogText", "插件日志：关闭");
        set => store.Value["ShowDnaLogText"] = value;
    }

    public string ShowAiText
    {
        get => store.Value.GetValueOrDefault("ShowAIText", "AI对话：关闭");
        set => store.Value["ShowAIText"] = value;
    }

    public string SpotlightMode
    {
        get => store.Value.GetValueOrDefault("SpotlightMode", "overlay");
        set => store.Value["SpotlightMode"] = value;
    }

    public string AgentCustomInstruction
    {
        get => store.Value.GetValueOrDefault("AgentCustomInstruction", "");
        set => store.Value["AgentCustomInstruction"] = value;
    }
}

public class GitConfig(GlobalVariable store)
{
    public string RootPath
    {
        get => store.Value.GetValueOrDefault("GitRootPath", "");
        set => store.Value["GitRootPath"] = value;
    }

    public bool SkipHashFiles
    {
        get => store.Value.GetValueOrDefault("ConflictSkipHashFiles", "false") == "true";
        set => store.Value["ConflictSkipHashFiles"] = value ? "true" : "false";
    }
}

public class PathConfig(GlobalVariable store)
{
    public string BasePath
    {
        get => store.Value.GetValueOrDefault("BasePath", @"C:\M1Work\Public\Excels\Tables\");
        set => store.Value["BasePath"] = value;
    }

    public string TargetPath
    {
        get => store.Value.GetValueOrDefault("TargetPath", @"C:\M2Work\Public\Excels\Tables\");
        set => store.Value["TargetPath"] = value;
    }

    public string TempPath
    {
        get => store.Value.GetValueOrDefault("TempPath", @"\Client\Assets\Resources\Table");
        set => store.Value["TempPath"] = value;
    }

    public string OutputRootPath
    {
        get =>
            store.Value.GetValueOrDefault(
                "OutputRootPath",
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "NumDesOutput"
                )
            );
        set => store.Value["OutputRootPath"] = value;
    }
}

public class AiPromptConfig(GlobalVariable store)
{
    public string ExcelAssistant
    {
        get => store.Value.GetValueOrDefault("ChatSysContentExcelAss", "");
        set => store.Value["ChatSysContentExcelAss"] = value;
    }

    public string TransferAssistant
    {
        get => store.Value.GetValueOrDefault("ChatSysContentTransferAss", "");
        set => store.Value["ChatSysContentTransferAss"] = value;
    }
}
