using NumDesTools.Config;

namespace NumDesTools;

/// <summary>
/// 强类型配置入口。内部读写同一份 GlobalVariable 字典，JSON 格式保持平铺兼容。
/// </summary>
public class AppConfig(GlobalVariable store)
{
    public LlmConfig Llm { get; } = new(store);
    public UiConfig Ui { get; } = new(store);
    public GitConfig Git { get; } = new(store);
    public PathConfig Paths { get; } = new(store);
    public AiPromptConfig AiPrompts { get; } = new(store);
    public AgentConfig Agent { get; } = new(store);

    public void Save() => store.SaveConfig();

    public void Save(string key, string value) => store.SaveValue(key, value);
}

public class LlmConfig(GlobalVariable store)
{
    public string ApiKey
    {
        get =>
            store.Value.GetValueOrDefault(
                "LiteLLMApiKey",
                store.DefaultValue.GetValueOrDefault("LiteLLMApiKey", "")
            );
        set => store.Value["LiteLLMApiKey"] = value;
    }

    public string ApiUrl
    {
        get =>
            store.Value.GetValueOrDefault(
                "LiteLLMApiUrl",
                store.DefaultValue.GetValueOrDefault("LiteLLMApiUrl", "")
            );
        set => store.Value["LiteLLMApiUrl"] = value;
    }

    public string Model
    {
        get =>
            store.Value.GetValueOrDefault(
                "LiteLLMModel",
                store.DefaultValue.GetValueOrDefault("LiteLLMModel", "")
            );
        set => store.Value["LiteLLMModel"] = value;
    }

    public List<string> ModelList
    {
        get =>
            store
                .Value.GetValueOrDefault(
                    "LiteLLMModelList",
                    store.DefaultValue.GetValueOrDefault("LiteLLMModelList", "")
                )
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
}

public class GitConfig(GlobalVariable store)
{
    public string RootPath
    {
        get =>
            store.Value.GetValueOrDefault(
                "GitRootPath",
                store.DefaultValue.GetValueOrDefault("GitRootPath", "")
            );
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
        get =>
            store.Value.GetValueOrDefault(
                "BasePath",
                store.DefaultValue.GetValueOrDefault("BasePath", @"C:\M1Work\Public\Excels\Tables\")
            );
        set => store.Value["BasePath"] = value;
    }

    public string TargetPath
    {
        get =>
            store.Value.GetValueOrDefault(
                "TargetPath",
                store.DefaultValue.GetValueOrDefault(
                    "TargetPath",
                    @"C:\M2Work\Public\Excels\Tables\"
                )
            );
        set => store.Value["TargetPath"] = value;
    }

    public string TempPath
    {
        get =>
            store.Value.GetValueOrDefault(
                "TempPath",
                store.DefaultValue.GetValueOrDefault("TempPath", @"\Client\Assets\Resources\Table")
            );
        set => store.Value["TempPath"] = value;
    }

    public string OutputRootPath
    {
        get =>
            store.Value.GetValueOrDefault(
                "OutputRootPath",
                store.DefaultValue.GetValueOrDefault(
                    "OutputRootPath",
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        "NumDesOutput"
                    )
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

public class AgentConfig(GlobalVariable store)
{
    public string CustomInstruction
    {
        get => store.Value.GetValueOrDefault("AgentCustomInstruction", "");
        set => store.Value["AgentCustomInstruction"] = value;
    }
}
