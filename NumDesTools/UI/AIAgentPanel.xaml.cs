using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows;
using System.Windows.Input;
using Markdig;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

[SupportedOSPlatform("windows")]
public partial class AIAgentPanel
{
    private CancellationTokenSource _cts;
    private readonly List<object> _history = [];

    private static readonly string HtmlTemplate =
        @"<html><head><meta charset='utf-8'><style>
body{background:#1c1c1c;color:#e0e0e0;font-family:微软雅黑,monospace;line-height:1.6;margin:0;padding:8px;overflow-y:auto}
.msg{margin:6px 0;max-width:96%}
.msg.user{margin-left:auto;text-align:right}
.msg.assistant{margin-left:0}
.role{font-size:.78em;color:#666;margin-bottom:2px}
.role .ts{color:#444}
.content{display:inline-block;padding:7px 10px;border-radius:8px;word-wrap:break-word;text-align:left;max-width:100%}
.user .content{background:#1e3a5f;color:#e0e0e0}
.assistant .content{background:#3e3e42;color:#e0e0e0}
pre{background:#2d2d30;color:#dcdcdc;padding:8px;border-radius:6px;overflow-x:auto;font-size:.85em}
code{font-family:Consolas,monospace;background:#2d2d30;padding:1px 3px;border-radius:3px;font-size:.85em}
a[href^='excel://']{color:#4ec9b0;text-decoration:none;border-bottom:1px dashed #4ec9b0}
a[href^='excel://']:hover{background:#1a3a35;border-radius:2px}
</style></head><body></body></html>";

    private static readonly object[] ToolDefinitions =
    [
        new
        {
            type = "function",
            function = new
            {
                name = "read_selection",
                description = "读取 Excel 当前选中区域的单元格值，返回二维数组字符串",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>()
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "write_range",
                description = "向指定单元格地址写入值",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        address = new { type = "string", description = "单元格地址，如 A1 或 B2:B10" },
                        value = new { type = "string", description = "要写入的值，多行用\\n分隔，多列用\\t分隔" },
                    },
                    required = new[] { "address", "value" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "run_formula",
                description = "在指定单元格写入公式并返回计算结果",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        address = new { type = "string", description = "目标单元格地址" },
                        formula = new { type = "string", description = "Excel 公式，以 = 开头" },
                    },
                    required = new[] { "address", "formula" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "list_udfs",
                description = "列出插件所有可用的 UDF 自定义函数名称",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>()
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "read_sheet",
                description = "读取指定 Sheet 的数据（前 N 行）",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new
                        {
                            type = "string",
                            description = "Sheet 名称，留空则读取当前活动 Sheet"
                        },
                        max_rows = new { type = "integer", description = "最多读取行数，默认 50" },
                    },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "list_sheets",
                description = "列出当前工作簿所有 Sheet 名称",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "get_workbook_structure",
                description = "获取工作簿结构：每个 Sheet 的行列数和前两行内容，用于快速了解数据布局",
                parameters = new
                {
                    type = "object",
                    properties = new { },
                    required = Array.Empty<string>(),
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "batch_write",
                description = "批量向多个单元格写入数据，一次调用写多个地址",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        sheet_name = new { type = "string", description = "Sheet 名称，留空则用当前活动 Sheet" },
                        writes = new
                        {
                            type = "array",
                            description = "写入列表，每项含 address 和 value",
                            items = new
                            {
                                type = "object",
                                properties = new
                                {
                                    address = new { type = "string" },
                                    value = new { type = "string" },
                                },
                            },
                        },
                    },
                    required = new[] { "writes" },
                },
            },
        },
        new
        {
            type = "function",
            function = new
            {
                name = "run_vba_macro",
                description = "执行一段 VBA 代码完成复杂操作（格式、筛选、复制、跨表等），代码在 Excel 中直接运行",
                parameters = new
                {
                    type = "object",
                    properties = new
                    {
                        code = new { type = "string", description = "完整的 VBA Sub 代码，包含 Sub...End Sub" },
                    },
                    required = new[] { "code" },
                },
            },
        },
    ];

    // 匹配 Sheet1!A1:B5 / A1:B5 / A1 形式的单元格地址
    private static readonly System.Text.RegularExpressions.Regex CellAddressRegex =
        new(
            @"(?<![""'#>\/])(?:([A-Za-z0-9_一-龥]+)!)?([A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?)\b",
            System.Text.RegularExpressions.RegexOptions.None
        );

    public AIAgentPanel()
    {
        InitializeComponent();
        PopulateModelList();
        ChatOutput.NavigateToString(HtmlTemplate);
        ChatOutput.Navigating += ChatOutput_Navigating;
        var saved = NumDesAddIn.GlobalValue.Value.ContainsKey("AgentCustomInstruction")
            ? NumDesAddIn.GlobalValue.Value["AgentCustomInstruction"]
            : "";
        CustomInstructionInput.Text = saved;
        CustomInstructionInput.LostFocus += (_, _) =>
            NumDesAddIn.GlobalValue.SaveValue("AgentCustomInstruction", CustomInstructionInput.Text);
    }

    private static void ChatOutput_Navigating(
        object sender,
        System.Windows.Navigation.NavigatingCancelEventArgs e
    )
    {
        var uri = e.Uri?.ToString() ?? "";
        if (!uri.StartsWith("excel://cell/"))
            return;
        e.Cancel = true;
        var address = Uri.UnescapeDataString(uri["excel://cell/".Length..]);
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                dynamic app = NumDesAddIn.App;
                if (address.Contains('!'))
                {
                    var parts = address.Split('!');
                    dynamic ws = app.ActiveWorkbook.Sheets[parts[0]];
                    ws.Activate();
                    ws.Range[parts[1]].Select();
                }
                else
                {
                    app.ActiveSheet.Range[address].Select();
                }
            });
        }
        catch { }
    }

    private void PopulateModelList()
    {
        ModelComboBox.Items.Clear();
        var models = NumDesAddIn.LiteLLMModelList;
        if (models.Count == 0)
            models = [NumDesAddIn.LiteLLMModel];
        foreach (var m in models)
            ModelComboBox.Items.Add(m);
        var current = NumDesAddIn.LiteLLMModel;
        ModelComboBox.SelectedItem = ModelComboBox.Items.Contains(current)
            ? current
            : ModelComboBox.Items[0];
    }

    private void TaskInput_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        if ((e.KeyboardDevice.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            TaskInput.AppendText(Environment.NewLine);
        else
            RunButton_Click(RunButton, new RoutedEventArgs());
    }

    private void StopButton_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        SetStatus("已停止");
    }

    private async void RunButton_Click(object sender, RoutedEventArgs e)
    {
        var task = TaskInput.Text.Trim();
        if (string.IsNullOrEmpty(task))
            return;

        _cts = new CancellationTokenSource();
        RunButton.IsEnabled = false;
        StopButton.IsEnabled = true;
        StepsList.Items.Clear();
        AppendChat("user", task);
        TaskInput.Clear();
        SetStatus("执行中…");

        try
        {
            await RunAgentLoopAsync(task, _cts.Token);
        }
        catch (OperationCanceledException)
        {
            AddStep("⛔ 已取消");
        }
        catch (Exception ex)
        {
            AddStep($"❌ 错误：{ex.Message}");
            SetStatus("出错");
        }
        finally
        {
            RunButton.IsEnabled = true;
            StopButton.IsEnabled = false;
        }
    }

    private async Task RunAgentLoopAsync(string userTask, CancellationToken ct)
    {
        var model = ModelComboBox.SelectedItem as string ?? NumDesAddIn.LiteLLMModel;
        var apiKey = NumDesAddIn.LiteLLMApiKey;
        var apiUrl = NumDesAddIn.LiteLLMApiUrl;
        var maxSteps = (int)(MaxStepsInput.Value ?? 10);

        if (_history.Count == 0)
        {
            var customInstruction = Dispatcher.Invoke(() => CustomInstructionInput.Text.Trim());
            var systemContent = "你是一个专业的 Excel 数据助手，可以调用工具对当前工作簿进行全面操作。\n"
                + "工作流程：1) 先用 get_workbook_structure 或 list_sheets 了解工作簿结构；"
                + "2) 用 read_sheet 读取相关数据；"
                + "3) 用 write_range/batch_write 写入结果，或用 run_vba_macro 执行复杂操作（格式、筛选、跨表复制等）。\n"
                + "每次只调用一个工具，等待结果后再决定下一步。完成后用 Markdown 输出简洁的结果说明。\n"
                + "重要：如果你在上一条消息中给出了编号选项（如 1. 2. 3.），用户回复单个数字时，请将其理解为选择对应选项，直接执行该选项对应的操作，不要再次询问。";
            if (!string.IsNullOrEmpty(customInstruction))
                systemContent += $"\n\n用户自定义指令（始终遵守）：{customInstruction}";
            _history.Add(new { role = "system", content = systemContent });
        }
        _history.Add(new { role = "user", content = userTask });
        var messages = _history;

        AddStep($"📋 {userTask[..Math.Min(40, userTask.Length)]}…");

        for (var step = 1; step <= maxSteps; step++)
        {
            ct.ThrowIfCancellationRequested();
            SetStatus($"步骤 {step}/{maxSteps}…");

            var (content, toolCalls) = await CallWithToolsAsync(
                model,
                messages,
                apiKey,
                apiUrl,
                ct
            );

            if (toolCalls is { Count: > 0 })
            {
                messages.Add(
                    new
                    {
                        role = "assistant",
                        content = content ?? "",
                        tool_calls = toolCalls
                    }
                );

                foreach (var tc in toolCalls)
                {
                    var toolName = tc["function"]?["name"]?.ToString() ?? "";
                    var argsJson = tc["function"]?["arguments"]?.ToString() ?? "{}";
                    var toolCallId = tc["id"]?.ToString() ?? $"tc_{step}";

                    AddStep($"🔧 {toolName}({argsJson[..Math.Min(50, argsJson.Length)]})");
                    var result = ExecuteTool(toolName, argsJson);
                    AddStep($"   ↳ {result[..Math.Min(70, result.Length)]}");

                    messages.Add(
                        new
                        {
                            role = "tool",
                            tool_call_id = toolCallId,
                            content = result
                        }
                    );
                }
            }
            else
            {
                AddStep($"✅ 完成（{step} 步）");
                SetStatus("完成");
                AppendChat("assistant", content ?? "（无输出）");
                return;
            }
        }

        AddStep("⚠️ 已达步骤上限");
        SetStatus("超出步骤上限");
    }

    private static async Task<(string content, List<JObject> toolCalls)> CallWithToolsAsync(
        string model,
        List<object> messages,
        string apiKey,
        string apiUrl,
        CancellationToken ct
    )
    {
        var body = new
        {
            model,
            messages,
            tools = ToolDefinitions,
            tool_choice = "auto",
            max_tokens = 4000
        };
        using var http = new System.Net.Http.HttpClient { Timeout = TimeSpan.FromMinutes(3) };
        using var req = new System.Net.Http.HttpRequestMessage(
            System.Net.Http.HttpMethod.Post,
            apiUrl
        );
        req.Content = new System.Net.Http.StringContent(
            JsonConvert.SerializeObject(body),
            System.Text.Encoding.UTF8,
            "application/json"
        );
        req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue(
            "Bearer",
            apiKey
        );
        using var resp = await http.SendAsync(req, ct);
        var json = JObject.Parse(await resp.Content.ReadAsStringAsync(ct));
        if (!resp.IsSuccessStatusCode)
            throw new Exception(json.ToString());
        var msg = json["choices"]?[0]?["message"];
        return (msg?["content"]?.ToString(), msg?["tool_calls"]?.ToObject<List<JObject>>());
    }

    private static string ExecuteTool(string toolName, string argsJson)
    {
        try
        {
            var args = JObject.Parse(argsJson);
            return toolName switch
            {
                "read_selection" => PubMetToExcel.ArrayToArrayStr(NumDesAddIn.App.Selection.Value2),
                "write_range"
                    => ToolWriteRange(
                        args["address"]?.ToString() ?? "",
                        args["value"]?.ToString() ?? ""
                    ),
                "run_formula"
                    => ToolRunFormula(
                        args["address"]?.ToString() ?? "",
                        args["formula"]?.ToString() ?? ""
                    ),
                "list_udfs" => ToolListUdfs(),
                "read_sheet"
                    => ToolReadSheet(
                        args["sheet_name"]?.ToString() ?? "",
                        (int)(args["max_rows"] ?? 50)
                    ),
                "list_sheets" => ToolListSheets(),
                "get_workbook_structure" => ToolGetWorkbookStructure(),
                "batch_write"
                    => ToolBatchWrite(
                        args["sheet_name"]?.ToString() ?? "",
                        args["writes"] as JArray ?? []
                    ),
                "run_vba_macro" => ToolRunVbaMacro(args["code"]?.ToString() ?? ""),
                _ => $"未知工具: {toolName}",
            };
        }
        catch (Exception ex)
        {
            return $"工具执行失败: {ex.Message}";
        }
    }

    private static string ToolWriteRange(string address, string value)
    {
        dynamic ws = NumDesAddIn.App.ActiveSheet;
        var range = ws.Range[address];
        var lines = value.Split('\n');
        for (var r = 0; r < lines.Length; r++)
        {
            var cols = lines[r].Split('\t');
            for (var c = 0; c < cols.Length; c++)
                range.Cells[r + 1, c + 1] = cols[c];
        }
        return $"已写入 {address}";
    }

    private static string ToolRunFormula(string address, string formula)
    {
        dynamic ws = NumDesAddIn.App.ActiveSheet;
        var cell = ws.Range[address];
        cell.Formula = formula;
        return $"{address} = {cell.Value2}";
    }

    private static string ToolListUdfs()
    {
        var udfs = System
            .Reflection.Assembly.GetExecutingAssembly()
            .GetTypes()
            .SelectMany(t =>
                t.GetMethods(
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                )
            )
            .Where(m =>
                m.GetCustomAttributes(
                    typeof(ExcelDna.Integration.ExcelFunctionAttribute),
                    false
                ).Length > 0
            )
            .Select(m => m.Name)
            .OrderBy(n => n)
            .ToList();
        return udfs.Count > 0 ? string.Join(", ", udfs) : "（未找到 UDF）";
    }

    private static string ToolReadSheet(string sheetName, int maxRows)
    {
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? NumDesAddIn.App.ActiveSheet
            : NumDesAddIn.App.ActiveWorkbook.Sheets[sheetName];
        var used = ws.UsedRange;
        var rows = Math.Min((int)used.Rows.Count, maxRows);
        var cols = (int)used.Columns.Count;
        var sb = new System.Text.StringBuilder();
        for (var r = 1; r <= rows; r++)
        {
            var line = new List<string>();
            for (var c = 1; c <= cols; c++)
                line.Add(used.Cells[r, c].Value2?.ToString() ?? "");
            sb.AppendLine(string.Join("\t", line));
        }
        return sb.ToString();
    }

    private static string ToolListSheets()
    {
        dynamic wb = NumDesAddIn.App.ActiveWorkbook;
        var names = new List<string>();
        foreach (dynamic ws in wb.Sheets)
            names.Add((string)ws.Name);
        return string.Join(", ", names);
    }

    private static string ToolGetWorkbookStructure()
    {
        dynamic wb = NumDesAddIn.App.ActiveWorkbook;
        var sb = new System.Text.StringBuilder();
        foreach (dynamic ws in wb.Sheets)
        {
            try
            {
                var used = ws.UsedRange;
                int rows = used.Rows.Count;
                int cols = used.Columns.Count;
                sb.AppendLine($"[{ws.Name}] {rows}行×{cols}列");
                var previewRows = Math.Min(rows, 2);
                for (var r = 1; r <= previewRows; r++)
                {
                    var line = new List<string>();
                    for (var c = 1; c <= Math.Min(cols, 10); c++)
                        line.Add(used.Cells[r, c].Value2?.ToString() ?? "");
                    sb.AppendLine("  " + string.Join("\t", line));
                }
            }
            catch
            {
                sb.AppendLine($"[{ws.Name}] (无法读取)");
            }
        }
        return sb.ToString();
    }

    private static string ToolBatchWrite(string sheetName, JArray writes)
    {
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? NumDesAddIn.App.ActiveSheet
            : NumDesAddIn.App.ActiveWorkbook.Sheets[sheetName];
        var count = 0;
        foreach (var item in writes)
        {
            var address = item["address"]?.ToString() ?? "";
            var value = item["value"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(address))
                continue;
            var lines = value.Split('\n');
            var range = ws.Range[address];
            for (var r = 0; r < lines.Length; r++)
            {
                var cols = lines[r].Split('\t');
                for (var c = 0; c < cols.Length; c++)
                    range.Cells[r + 1, c + 1] = cols[c];
            }
            count++;
        }
        return $"已写入 {count} 个区域";
    }

    private static string ToolRunVbaMacro(string code)
    {
        try
        {
            dynamic wb = NumDesAddIn.App.ActiveWorkbook;
            var vbProj = wb.VBProject;
            var module = vbProj.VBComponents.Add(1); // vbext_ct_StdModule = 1
            module.CodeModule.AddFromString(code);
            var macroName = wb.Name + "!" + module.Name + ".Main";
            // 尝试找 Sub 名称
            var match = System.Text.RegularExpressions.Regex.Match(code, @"Sub\s+(\w+)\s*\(");
            if (match.Success)
                macroName = wb.Name + "!" + module.Name + "." + match.Groups[1].Value;
            NumDesAddIn.App.Run(macroName);
            vbProj.VBComponents.Remove(module);
            return "VBA 执行完成";
        }
        catch (Exception ex)
        {
            return $"VBA 执行失败: {ex.Message}";
        }
    }

    private void AddStep(string text) =>
        Dispatcher.Invoke(() =>
        {
            StepsList.Items.Add(text);
            StepsScroll.ScrollToBottom();
        });

    private void SetStatus(string text) => Dispatcher.Invoke(() => StatusText.Text = text);

    private static string InjectCellLinks(string html) =>
        CellAddressRegex.Replace(
            html,
            m =>
            {
                var sheet = m.Groups[1].Value;
                var cell = m.Groups[2].Value;
                var address = string.IsNullOrEmpty(sheet) ? cell : $"{sheet}!{cell}";
                var encoded = Uri.EscapeDataString(address);
                return $"<a href='excel://cell/{encoded}' style='color:#4ec9b0;text-decoration:none' title='定位到 {address}'>{m.Value}</a>";
            }
        );

    private void AppendChat(string role, string markdown)
    {
        Dispatcher.Invoke(() =>
        {
            var html = InjectCellLinks(HttpUtility.HtmlDecode(Markdown.ToHtml(markdown)));
            var cls = role == "user" ? "user" : "assistant";
            var label =
                role == "user"
                    ? Environment.UserName
                    : (ModelComboBox.SelectedItem as string ?? "Agent");
            var ts = DateTime.Now.ToString("HH:mm:ss");
            var block =
                $"<div class='msg {cls}'><div class='role'>{label} <span class='ts'>{ts}</span></div><div class='content'>{html}</div></div>";
            ChatOutput.InvokeScript(
                "eval",
                $"document.body.insertAdjacentHTML('beforeend','{HttpUtility.JavaScriptStringEncode(block)}');window.scrollTo(0,document.body.scrollHeight);"
            );
        });
    }
}
