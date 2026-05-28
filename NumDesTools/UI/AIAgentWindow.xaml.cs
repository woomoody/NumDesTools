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
public partial class AIAgentWindow
{
    private CancellationTokenSource _cts;
    private static readonly string HtmlTemplate =
        @"<html><head><meta charset='utf-8'><style>
body{background:#1c1c1c;color:#e0e0e0;font-family:微软雅黑,monospace;line-height:1.6;margin:0;padding:10px;overflow-y:auto}
.msg{margin:8px 0;max-width:92%}
.msg.user{margin-left:auto;text-align:right}
.msg.assistant{margin-left:0}
.role{font-size:.8em;color:#888;margin-bottom:3px}
.role .ts{color:#555}
.content{display:inline-block;padding:8px 12px;border-radius:8px;word-wrap:break-word;text-align:left}
.user .content{background:#2d4a7a;color:#e0e0e0}
.assistant .content{background:#3e3e42;color:#e0e0e0}
pre{background:#2d2d30;color:#dcdcdc;padding:10px;border-radius:6px;overflow-x:auto}
code{font-family:Consolas,monospace;background:#2d2d30;padding:2px 4px;border-radius:3px}
</style></head><body></body></html>";

    // ── 工具定义（OpenAI function calling 格式）────────────────────────────────
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
    ];

    public AIAgentWindow()
    {
        InitializeComponent();
        PopulateModelList();
        ChatOutput.NavigateToString(HtmlTemplate);
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

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
            Close();
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
        ChatOutput.NavigateToString(HtmlTemplate);
        AppendChat("user", task);
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

    // ── Agent 执行循环 ────────────────────────────────────────────────────────

    private async Task RunAgentLoopAsync(string userTask, CancellationToken ct)
    {
        var model = ModelComboBox.SelectedItem as string ?? NumDesAddIn.LiteLLMModel;
        var apiKey = NumDesAddIn.LiteLLMApiKey;
        var apiUrl = NumDesAddIn.LiteLLMApiUrl;
        var maxSteps = (int)(MaxStepsInput.Value ?? 10);

        var messages = new List<object>
        {
            new
            {
                role = "system",
                content = "你是一个 Excel 数据助手，可以调用工具读写 Excel 数据。"
                    + "每次只调用一个工具，等待结果后再决定下一步。"
                    + "完成任务后输出最终结果的 Markdown 说明。",
            },
            new { role = "user", content = userTask },
        };

        AddStep($"📋 任务：{userTask[..Math.Min(40, userTask.Length)]}…");

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

                    AddStep($"🔧 [{step}] {toolName}({argsJson[..Math.Min(60, argsJson.Length)]})");

                    var result = ExecuteTool(toolName, argsJson);
                    AddStep($"   ↳ {result[..Math.Min(80, result.Length)]}");

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
                // AI 给出最终文字回复，结束循环
                AddStep($"✅ 完成（共 {step} 步）");
                SetStatus("完成");
                AppendChat("assistant", content ?? "（无输出）");
                return;
            }
        }

        AddStep("⚠️ 已达最大步骤上限");
        SetStatus("超出步骤上限");
    }

    // ── 带工具调用的 API 请求 ──────────────────────────────────────────────────

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
            max_tokens = 4000,
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

        var choice = json["choices"]?[0];
        var msg = choice?["message"];
        var content = msg?["content"]?.ToString();
        var tcs = msg?["tool_calls"]?.ToObject<List<JObject>>();
        return (content, tcs);
    }

    // ── 工具执行 ──────────────────────────────────────────────────────────────

    private static string ExecuteTool(string toolName, string argsJson)
    {
        try
        {
            var args = JObject.Parse(argsJson);
            return toolName switch
            {
                "read_selection" => ToolReadSelection(),
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
                _ => $"未知工具: {toolName}",
            };
        }
        catch (Exception ex)
        {
            return $"工具执行失败: {ex.Message}";
        }
    }

    private static string ToolReadSelection()
    {
        var sel = NumDesAddIn.App.Selection;
        return PubMetToExcel.ArrayToArrayStr(sel.Value2);
    }

    private static string ToolWriteRange(string address, string value)
    {
        var ws = NumDesAddIn.App.ActiveSheet;
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
        var ws = NumDesAddIn.App.ActiveSheet;
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
        var app = NumDesAddIn.App;
        dynamic ws = string.IsNullOrEmpty(sheetName)
            ? app.ActiveSheet
            : app.ActiveWorkbook.Sheets[sheetName];
        var used = ws.UsedRange;
        var rows = Math.Min(used.Rows.Count, maxRows);
        var cols = used.Columns.Count;
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

    // ── UI 辅助 ───────────────────────────────────────────────────────────────

    private void AddStep(string text)
    {
        Dispatcher.Invoke(() =>
        {
            StepsList.Items.Add(text);
            StepsScroll.ScrollToBottom();
        });
    }

    private void SetStatus(string text) => Dispatcher.Invoke(() => StatusText.Text = text);

    private void AppendChat(string role, string markdown)
    {
        Dispatcher.Invoke(() =>
        {
            var html = HttpUtility.HtmlDecode(Markdown.ToHtml(markdown));
            var cls = role == "user" ? "user" : "assistant";
            var label =
                role == "user"
                    ? Environment.UserName
                    : (ModelComboBox.SelectedItem as string ?? "Agent");
            var ts = DateTime.Now.ToString("HH:mm:ss");
            var block =
                $"<div class='msg {cls}'>"
                + $"<div class='role'>{label} <span class='ts'>{ts}</span></div>"
                + $"<div class='content'>{html}</div></div>";
            var script =
                $"document.body.insertAdjacentHTML('beforeend','{HttpUtility.JavaScriptStringEncode(block)}');"
                + "window.scrollTo(0,document.body.scrollHeight);";
            ChatOutput.InvokeScript("eval", script);
        });
    }
}
