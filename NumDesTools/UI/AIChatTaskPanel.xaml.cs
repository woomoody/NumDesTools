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
    private static readonly MarkdownPipeline MdPipeline = new MarkdownPipelineBuilder()
        .UseAdvancedExtensions()
        .Build();

    private readonly string _userName = Environment.UserName;
    private string _currentResponseId;
    private string _streamBuffer = ""; // 流式累积文本
    private int _chunkCount; // 已收到 chunk 数
    private const int ReRenderEvery = 8; // 每 N 个 chunk 重渲染一次 MD

    // 多轮上下文，最近 20 条
    private readonly List<object> _history = [];
    private const int MaxHistoryRounds = 20;
    private const int HistoryPageSize = 50;
    private int _historyOffset; // 已加载条数（从最新往旧算）

    private string _sessionId = Guid.NewGuid().ToString("N")[..12];

    private record SessionItem(string SessionId, string Display)
    {
        public override string ToString() => Display;
    }

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

        // 动态高度：根据内容行数自动扩展（无滚动条）
        PromptInput.Document.Changed += (_, _) =>
        {
            var lineCount = PromptInput.Document.LineCount;
            var lineHeight = PromptInput.FontSize * 1.5 + 2;
            var newHeight = Math.Max(36, Math.Min(200, lineCount * lineHeight + 12));
            PromptInput.Height = newHeight;
        };

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
body{background:#1c1c1c;color:#d4d4d4;font-family:'微软雅黑',sans-serif;font-size:10pt;line-height:1.5;margin:0;padding:8px 10px;overflow-y:auto}
.message-container{display:flex;flex-direction:column;align-items:flex-start;margin:5px 0}
.message{padding:6px 10px;border-radius:7px;max-width:92%;word-wrap:break-word}
.message p{margin:3px 0}
.user{background:#1e3a5f;border:1px solid #2a4a6f;margin-left:auto;margin-right:6px;color:#d4d4d4}
.system{background:#333337;border:1px solid #444;margin-left:6px;color:#d4d4d4}
.role{font-weight:bold;margin-bottom:3px;font-size:.72em;color:#888}
.timestamp{font-size:.72em;color:#555;margin-top:3px;margin-left:8px;margin-right:8px}
.user+.timestamp{text-align:right;margin-left:auto;margin-right:8px}
pre{background:#252526;color:#dcdcdc;padding:7px;border-radius:5px;overflow-x:auto;font-size:10pt;margin:4px 0}
code{font-family:Consolas,monospace;background:#252526;color:#dcdcdc;padding:1px 3px;border-radius:3px;font-size:10pt}
table{border-collapse:collapse;font-size:.88em;margin:4px 0;width:auto}
th,td{border:1px solid #555;padding:3px 8px;text-align:left;white-space:nowrap}
th{background:#2a2d2e;color:#c6c6c6;font-weight:bold}
tr:nth-child(even) td{background:#2a2a2a}
ul,ol{margin:3px 0;padding-left:18px}
li{margin:1px 0}
h1,h2,h3{font-size:1em;font-weight:bold;margin:4px 0 2px}
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

    private async void CompressButton_Click(object sender, RoutedEventArgs e)
    {
        if (_history.Count < 2)
            return;
        CompressButton.IsEnabled = false;
        CompressButton.Content = "压缩中…";
        try
        {
            var model = AppServices.Config.Llm.Model;
            var apiKey = AppServices.Config.Llm.ApiKey;
            var apiUrl = AppServices.Config.Llm.ChatCompletionsUrl;
            var msgs = new List<object>();
            msgs.AddRange(_history);
            msgs.Add(new
            {
                role = "user",
                content = "请将上面的完整对话内容压缩为一段结构化摘要，保留所有关键数据、结论和操作记录，供后续对话参考。直接输出摘要，不加解释。",
            });
            var sb = new System.Text.StringBuilder();
            await ChatApiClient.CallApiStreamAsync(model, msgs, apiKey, apiUrl,
                chunk => sb.Append(chunk));
            var summary = sb.ToString();
            _history.Clear();
            _history.Add(new { role = "assistant", content = $"[对话摘要]\n{summary}" });
            ResponseOutput.InvokeScript("eval", "clearAll()");
            AppendMessage("系统(摘要)", $"**上下文已压缩**\n\n{summary}", isUser: false, DateTime.Now);
        }
        catch (Exception ex)
        {
            AppendMessage("系统", $"压缩失败: {ex.Message}", isUser: false, DateTime.Now);
        }
        finally
        {
            CompressButton.IsEnabled = true;
            CompressButton.Content = "压缩上下文";
        }
    }

    private void RenameSessionButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_sessionId))
            return;
        var current = (SessionComboBox.SelectedItem as SessionItem)?.Display ?? "";
        var newTitle = ShowRenameDialog(current);
        if (string.IsNullOrEmpty(newTitle))
            return;
        new ChatHistoryManager().SaveSessionTitle(_sessionId, newTitle, isAgent: false);
        RefreshSessionList();
    }

    private void DeleteSessionButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(_sessionId))
            return;
        var preview = (SessionComboBox.SelectedItem as SessionItem)?.Display ?? _sessionId;
        var confirm = System.Windows.MessageBox.Show(
            $"确认删除会话？\n{preview}",
            "删除会话",
            System.Windows.MessageBoxButton.OKCancel,
            System.Windows.MessageBoxImage.Warning
        );
        if (confirm != System.Windows.MessageBoxResult.OK)
            return;
        new ChatHistoryManager().DeleteSession(_sessionId);
        _sessionId = Guid.NewGuid().ToString("N")[..12];
        _history.Clear();
        _historyOffset = 0;
        ResponseOutput.InvokeScript("eval", "clearAll()");
        RefreshSessionList();
    }

    private void ImportButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "SQLite 数据库|*.db|所有文件|*.*",
            Title = "选择要导入的会话数据库",
        };
        if (dialog.ShowDialog() != true)
            return;
        try
        {
            var count = new ChatHistoryManager().ImportSessionsFromDb(dialog.FileName);
            RefreshSessionList();
            System.Windows.MessageBox.Show($"已导入 {count} 条会话", "导入成功", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"导入失败: {ex.Message}", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void ClearAllButton_Click(object sender, RoutedEventArgs e)
    {
        var confirm = System.Windows.MessageBox.Show(
            "确认清空全部 Chat 历史记录？此操作不可恢复。",
            "清空全部",
            System.Windows.MessageBoxButton.OKCancel,
            System.Windows.MessageBoxImage.Warning
        );
        if (confirm != System.Windows.MessageBoxResult.OK)
            return;
        new ChatHistoryManager().DeleteAllHistory(isAgent: false);
        _sessionId = Guid.NewGuid().ToString("N")[..12];
        _history.Clear();
        _historyOffset = 0;
        ResponseOutput.InvokeScript("eval", "clearAll()");
        RefreshSessionList();
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
        var messages = new List<object> { new { role = "system", content = sysContent } };
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
            IsUser = false,
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

            _streamBuffer = "";
            _chunkCount = 0;
            var htmlMessage = HttpUtility.HtmlDecode(
                Markdown.ToHtml(streamMessage.Message, MdPipeline)
            );
            streamMessage.Message = htmlMessage;

            // 追加 AI 回复到历史
            _history.Add(new { role = "assistant", content = streamMessage.Message });

            streamMessage.SessionId = _sessionId;
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
        Dispatcher.BeginInvoke(() => RefreshSessionList());
    }

    // ── 流式内容追加 ──────────────────────────────────────────────────────────

    private void AppendStreamingChunk(string chunk)
    {
        _streamBuffer += chunk;
        _chunkCount++;
        Dispatcher.Invoke(() =>
        {
            if (string.IsNullOrEmpty(_currentResponseId))
            {
                _currentResponseId = $"msg-{DateTime.Now.Ticks}";
                _streamBuffer = "";
                _chunkCount = 0;
                var model = AppServices.Config.Llm.Model;
                AppendRawHtml(
                    $"<div id='{_currentResponseId}' class='message-container'>"
                        + $"<div class='message system'>"
                        + $"<div class='role'>{model}</div>"
                        + $"<div class='content'></div></div>"
                        + $"<div class='timestamp'></div></div>"
                );
            }

            // 每 ReRenderEvery 个 chunk 重渲染一次完整 MD，实现流式 MD 效果
            if (_chunkCount % ReRenderEvery == 0 || chunk.Contains('\n'))
            {
                var rendered = HttpUtility.JavaScriptStringEncode(
                    HttpUtility.HtmlDecode(Markdown.ToHtml(_streamBuffer, MdPipeline))
                );
                var script =
                    $"var c=document.getElementById('{_currentResponseId}');"
                    + $"var d=c.querySelector('.content');"
                    + $"d.innerHTML='{rendered}';"
                    + "scrollToBottom();";
                ResponseOutput.InvokeScript("eval", script);
            }
            else
            {
                // 非渲染帧：追加纯文本（速度快，不影响最终 MD）
                var script =
                    $"var c=document.getElementById('{_currentResponseId}');"
                    + $"var d=c.querySelector('.content');"
                    + $"d.innerHTML+='{HttpUtility.JavaScriptStringEncode(chunk)}';"
                    + "scrollToBottom();";
                ResponseOutput.InvokeScript("eval", script);
            }
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
            var htmlMessage = HttpUtility.HtmlDecode(Markdown.ToHtml(message, MdPipeline));
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
                        SessionId = _sessionId,
                    }
                );
        });
    }

    // ── 历史加载 ──────────────────────────────────────────────────────────────

    private void LoadChatHistory()
    {
        var db = new ChatHistoryManager();
        var sessions = db.ListSessionsWithPreview(isAgent: false);
        if (sessions.Count > 0)
            _sessionId = sessions[0].SessionId;

        _historyOffset = HistoryPageSize;
        var history = db.LoadChatHistory(HistoryPageSize, sessionId: _sessionId, isAgent: false);
        if (history.Count == 0)
        {
            Dispatcher.BeginInvoke(() => RefreshSessionList());
            return;
        }

        // 追加历史到多轮上下文
        foreach (var m in history.TakeLast(MaxHistoryRounds * 2))
            _history.Add(new { role = m.IsUser ? "user" : "assistant", content = m.Message });

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
            RefreshSessionList();
        });
    }

    private void RefreshSessionList()
    {
        SessionComboBox.SelectionChanged -= SessionComboBox_SelectionChanged;
        SessionComboBox.Items.Clear();
        var sessions = new ChatHistoryManager().ListSessionsWithPreview(isAgent: false);
        foreach (var s in sessions)
        {
            var label = !string.IsNullOrEmpty(s.Title)
                ? s.Title
                : $"{s.LastTime:MM-dd HH:mm}  {s.Preview}";
            SessionComboBox.Items.Add(new SessionItem(s.SessionId, label));
        }
        var current = SessionComboBox
            .Items.OfType<SessionItem>()
            .FirstOrDefault(x => x.SessionId == _sessionId);
        SessionComboBox.SelectedItem = current;
        SessionComboBox.SelectionChanged += SessionComboBox_SelectionChanged;
    }

    private void NewChatButton_Click(object sender, RoutedEventArgs e)
    {
        _sessionId = Guid.NewGuid().ToString("N")[..12];
        _history.Clear();
        _historyOffset = 0;
        ResponseOutput.InvokeScript("eval", "clearAll()");
        RefreshSessionList();
    }

    private void SessionComboBox_SelectionChanged(
        object sender,
        System.Windows.Controls.SelectionChangedEventArgs e
    )
    {
        if (SessionComboBox.SelectedItem is not SessionItem item)
            return;
        if (item.SessionId == _sessionId)
            return;
        SwitchToSession(item.SessionId);
    }

    private void SwitchToSession(string sessionId)
    {
        _sessionId = sessionId;
        _history.Clear();
        _historyOffset = 0;
        ResponseOutput.InvokeScript("eval", "clearAll()");
        var db = new ChatHistoryManager();
        var messages = db.LoadChatHistory(HistoryPageSize, sessionId: sessionId, isAgent: false);
        _historyOffset = messages.Count;
        foreach (var m in messages.TakeLast(MaxHistoryRounds * 2))
            _history.Add(new { role = m.IsUser ? "user" : "assistant", content = m.Message });
        var sb = new System.Text.StringBuilder();
        foreach (var m in messages)
            sb.Append(BuildMessageHtml(m.Role, m.Message, m.IsUser, m.Timestamp));
        var escaped = System.Web.HttpUtility.JavaScriptStringEncode(sb.ToString());
        ResponseOutput.InvokeScript(
            "eval",
            $"document.body.insertAdjacentHTML('beforeend','{escaped}');scrollToBottom();"
        );
    }

    internal static string? ShowRenameDialog(string current)
    {
        using var form = new System.Windows.Forms.Form
        {
            Text = "重命名会话", Width = 420, Height = 120,
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false,
        };
        var tb = new System.Windows.Forms.TextBox
        {
            Location = new System.Drawing.Point(10, 15), Width = 384, Text = current,
        };
        var ok = new System.Windows.Forms.Button
        {
            Text = "确定", DialogResult = System.Windows.Forms.DialogResult.OK,
            Location = new System.Drawing.Point(230, 50), Width = 80,
        };
        var cancel = new System.Windows.Forms.Button
        {
            Text = "取消", DialogResult = System.Windows.Forms.DialogResult.Cancel,
            Location = new System.Drawing.Point(314, 50), Width = 80,
        };
        form.Controls.AddRange([tb, ok, cancel]);
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        tb.SelectAll();
        return form.ShowDialog() == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(tb.Text)
            ? tb.Text.Trim()
            : null;
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
