using System.Runtime.Versioning;
using System.Web;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Markdig;
using Brushes = System.Windows.Media.Brushes;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

[SupportedOSPlatform("windows")]
public partial class AiChatTaskPanel
{
    private readonly string _userName = Environment.UserName;
    private string _currentResponseId;

    // 多轮上下文，最近 20 条
    private readonly List<object> _history = [];
    private const int MaxHistoryRounds = 20;
    private const int HistoryPageSize = 50;
    private int _historyOffset; // 已加载条数（从最新往旧算）

    private const string DefaultPromptText =
        "Enter 发送，Shift+Enter 换行，聊天框内容右键复制\n首字 ### 会把当前选中单元格值一并输入";

    public AiChatTaskPanel()
    {
        InitializeComponent();

        var enterBinding = PromptInput.TextArea.DefaultInputHandler.CommandBindings.FirstOrDefault(
            b => b.Command == EditingCommands.EnterParagraphBreak
        );
        if (enterBinding is not null)
            PromptInput.TextArea.DefaultInputHandler.CommandBindings.Remove(enterBinding);

        PromptInput.Text = DefaultPromptText;
        PromptInput.Foreground = Brushes.Gray;

        InitializeHtmlTemplate();
        PopulateModelList();
    }

    // ── 初始化 ────────────────────────────────────────────────────────────────

    private void PopulateModelList()
    {
        ModelComboBox.Items.Clear();
        var models = AppServices.Config.Llm.ModelList;
        if (models.Count == 0)
            models = [AppServices.Config.Llm.Model];
        foreach (var m in models)
            ModelComboBox.Items.Add(m);

        var current = AppServices.Config.Llm.Model;
        ModelComboBox.SelectedItem = ModelComboBox.Items.Contains(current)
            ? current
            : ModelComboBox.Items[0];
    }

    private void InitializeHtmlTemplate()
    {
        ResponseOutput.NavigateToString(
            @"<html>
<head>
<meta charset='utf-8'>
<style>
body{background:#1c1c1c;color:#e0e0e0;font-family:微软雅黑,monospace;line-height:1.6;margin:0;padding:10px;overflow-y:auto}
.message-container{display:flex;flex-direction:column;align-items:flex-start;margin:10px 0}
.message{padding:10px;border-radius:8px;max-width:90%;word-wrap:break-word}
.user{background:#2d2d30;border:1px solid #3e3e42;margin-left:auto;margin-right:10px}
.system{background:#3e3e42;border:1px solid #5a5a5e;margin-left:10px}
.role{font-weight:bold;margin-bottom:5px;font-size:.85em;color:#aaa}
.timestamp{font-size:.75em;color:gray;margin-top:5px;margin-left:10px;margin-right:10px}
.user+.timestamp{text-align:right;margin-left:auto;margin-right:10px}
pre{background:#2d2d30;color:#dcdcdc;padding:10px;border-radius:8px;overflow-x:auto}
code{font-family:Consolas,monospace;background:#2d2d30;color:#dcdcdc;padding:2px 4px;border-radius:4px}
</style>
<script>
function scrollToBottom(){window.scrollTo(0,document.body.scrollHeight)}
function replaceContent(id,html){
  var c=document.getElementById(id);
  if(c){var d=c.querySelector('.content');if(d)d.innerHTML=html}
}
function clearAll(){document.body.innerHTML=''}
</script>
</head>
<body></body>
</html>"
        );
        LoadChatHistory();
    }

    // ── 事件处理 ──────────────────────────────────────────────────────────────

    private void ModelComboBox_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e
    )
    {
        if (ModelComboBox.SelectedItem is string model)
        {
            AppServices.Config.Llm.Model = model;
            AppServices.GlobalValue.SaveValue("LiteLLMModel", model);
        }
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        _history.Clear();
        _historyOffset = 0;
        ResponseOutput.InvokeScript("eval", "clearAll()");
    }

    private void LoadMoreButton_Click(object sender, RoutedEventArgs e)
    {
        var db = new ChatHistoryManager();
        var total = db.GetHistoryCount();
        var nextOffset = _historyOffset + HistoryPageSize;
        if (nextOffset > total)
            nextOffset = total;
        var olderCount = nextOffset - _historyOffset;
        if (olderCount <= 0)
            return;
        // 取更老的一批：跳过最新 _historyOffset 条，再取 olderCount 条
        var older = db.LoadChatHistory(limit: nextOffset).Take(olderCount).ToList();
        _historyOffset = nextOffset;
        // 批量注入到顶部
        var sb = new System.Text.StringBuilder();
        foreach (var m in older)
            sb.Append(BuildMessageHtml(m.Role, m.Message, m.IsUser, m.Timestamp));
        var escaped = System.Web.HttpUtility.JavaScriptStringEncode(sb.ToString());
        ResponseOutput.InvokeScript(
            "eval",
            $"document.body.insertAdjacentHTML('afterbegin','{escaped}');"
        );
    }

    private void SendButton_Click(object sender, RoutedEventArgs e) => ProcessInput();

    private void PromptInput_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        if ((e.KeyboardDevice.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            PromptInput.Document.Insert(PromptInput.CaretOffset, Environment.NewLine);
        else
            ProcessInput();
    }

    private void PromptInput_GotFocus(object sender, RoutedEventArgs e)
    {
        if (PromptInput.Text != DefaultPromptText)
            return;
        PromptInput.Text = string.Empty;
        PromptInput.Foreground = Brushes.White;
    }

    private void PromptInput_LostFocus(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrWhiteSpace(PromptInput.Text))
            return;
        PromptInput.Text = DefaultPromptText;
        PromptInput.Foreground = Brushes.Gray;
    }

    // ── 核心发送逻辑 ──────────────────────────────────────────────────────────

    private async void ProcessInput()
    {
        var apiKey = AppServices.Config.Llm.ApiKey;
        var apiUrl = AppServices.Config.Llm.ChatCompletionsUrl;
        var model = AppServices.Config.Llm.Model;

        var userInput = PromptInput.Document.Text.Trim();

        if (userInput.StartsWith("###"))
        {
            var sel = AppServices.App.Selection;
            var val = PubMetToExcel.ArrayToArrayStr(sel.Value2);
            userInput = val + "," + userInput["###".Length..];
        }

        if (string.IsNullOrEmpty(userInput))
            return;

        // 构建带上下文的消息列表
        var sysContent = AppServices.Config.AiPrompts.ExcelAssistant;
        var messages = new List<object> { new { role = "system", content = sysContent }, };
        messages.AddRange(_history.TakeLast(MaxHistoryRounds * 2));
        messages.Add(new { role = "user", content = userInput });

        // 追加用户消息到历史
        _history.Add(new { role = "user", content = userInput });

        AppendMessage(_userName, userInput, isUser: true, DateTime.Now);

        _currentResponseId = null;
        var streamMessage = new ChatMessage
        {
            Role = model,
            Message = "",
            IsUser = false
        };

        try
        {
            await ChatApiClient.CallApiStreamAsync(
                model,
                messages,
                apiKey,
                apiUrl,
                chunk =>
                {
                    streamMessage.Message += chunk;
                    AppendStreamingChunk(chunk);
                },
                () =>
                {
                    var now = DateTime.Now;
                    streamMessage.Timestamp = now;
                    Dispatcher.Invoke(() =>
                    {
                        var script =
                            $"var c=document.getElementById('{_currentResponseId}');"
                            + $"var t=c.querySelector('.timestamp');"
                            + $"t.innerHTML='{now:yyyy-MM-dd HH:mm:ss}';";
                        ResponseOutput.InvokeScript("eval", script);
                    });
                }
            );

            var htmlMessage = HttpUtility.HtmlDecode(Markdown.ToHtml(streamMessage.Message));
            streamMessage.Message = htmlMessage;

            // 追加 AI 回复到历史
            _history.Add(new { role = "assistant", content = streamMessage.Message });

            await new ChatHistoryManager().SaveChatMessageAsync(streamMessage);
            ResponseOutput.InvokeScript("replaceContent", _currentResponseId, htmlMessage);
            ResponseOutput.InvokeScript("eval", "scrollToBottom()");
        }
        catch (Exception ex)
        {
            AppendMessage(model, $"调用 AI 时出错：{ex.Message}", isUser: false, DateTime.Now);
        }

        PromptInput.Document.Text = string.Empty;
        PromptInput.Focus();
    }

    // ── 流式内容追加 ──────────────────────────────────────────────────────────

    private void AppendStreamingChunk(string chunk)
    {
        Dispatcher.Invoke(() =>
        {
            if (string.IsNullOrEmpty(_currentResponseId))
            {
                _currentResponseId = $"msg-{DateTime.Now.Ticks}";
                var model = AppServices.Config.Llm.Model;
                AppendRawHtml(
                    $"<div id='{_currentResponseId}' class='message-container'>"
                        + $"<div class='message system'>"
                        + $"<div class='role'>{model}</div>"
                        + $"<div class='content'>思考中……</div></div>"
                        + $"<div class='timestamp'></div></div>"
                );
            }

            var script =
                $"var c=document.getElementById('{_currentResponseId}');"
                + $"var d=c.querySelector('.content');"
                + $"d.innerHTML+='{HttpUtility.JavaScriptStringEncode(chunk)}';"
                + "scrollToBottom();";
            ResponseOutput.InvokeScript("eval", script);
        });
    }

    private void AppendRawHtml(string html)
    {
        var script =
            $"document.body.insertAdjacentHTML('beforeend','{HttpUtility.JavaScriptStringEncode(html)}');scrollToBottom();";
        ResponseOutput.InvokeScript("eval", script);
    }

    // ── 消息渲染（历史加载 + 普通追加共用） ──────────────────────────────────

    private void AppendMessage(string role, string message, bool isUser, DateTime? timestamp)
    {
        Dispatcher.BeginInvoke(() =>
        {
            var htmlMessage = HttpUtility.HtmlDecode(Markdown.ToHtml(message));
            var ts = timestamp ?? DateTime.Now;
            AppendRawHtml(BuildMessageHtml(role, htmlMessage, isUser, ts));

            if (isUser)
                _ = new ChatHistoryManager().SaveChatMessageAsync(
                    new ChatMessage
                    {
                        Role = role,
                        Message = htmlMessage,
                        IsUser = true,
                        Timestamp = ts,
                    }
                );
        });
    }

    // ── 历史加载 ──────────────────────────────────────────────────────────────

    private void LoadChatHistory()
    {
        var db = new ChatHistoryManager();
        _historyOffset = HistoryPageSize;
        var history = db.LoadChatHistory(HistoryPageSize);
        if (history.Count == 0)
            return;

        // 批量拼接后一次注入，避免逐条 InvokeScript
        var sb = new System.Text.StringBuilder();
        foreach (var m in history)
            sb.Append(BuildMessageHtml(m.Role, m.Message, m.IsUser, m.Timestamp));

        var escaped = System.Web.HttpUtility.JavaScriptStringEncode(sb.ToString());
        Dispatcher.BeginInvoke(() =>
        {
            ResponseOutput.InvokeScript(
                "eval",
                $"document.body.insertAdjacentHTML('beforeend','{escaped}');scrollToBottom();"
            );
        });
    }

    private static string BuildMessageHtml(
        string role,
        string message,
        bool isUser,
        DateTime? timestamp
    )
    {
        var ts = (timestamp ?? DateTime.Now).ToString("yyyy-MM-dd HH:mm:ss");
        var cls = isUser ? "user" : "system";
        return $"<div class='message-container'>"
            + $"<div class='message {cls}'>"
            + $"<div class='role'>{role}</div>"
            + $"<div>{message}</div></div>"
            + $"<div class='timestamp'>{ts}</div></div>";
    }
}
