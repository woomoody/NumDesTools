using System.Web;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Markdig;
using Brushes = System.Windows.Media.Brushes;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI;

public partial class AiChatTaskPanel
{
    private string _apiKey;
    private string _apiUrl;
    private string _apiModel;
    private string _sysContent;

    private readonly string _userName = Environment.UserName;

    private const string DefaultPromptText =
        "Enter发送，Shift + Enter换行，聊天框内容复制使用右键\n首字输入###会把当前选择单元格值一并输入"; // 输入框默认文本

    private string _currentResponseId; // 用于跟踪当前响应消息

    public AiChatTaskPanel()
    {
        InitializeComponent();

        // 禁用 AvalonEdit 默认的 Enter 键行为
        var enterCommandBinding =
            PromptInput.TextArea.DefaultInputHandler.CommandBindings.FirstOrDefault(binding =>
                binding.Command == EditingCommands.EnterParagraphBreak
            );
        if (enterCommandBinding != null)
            PromptInput.TextArea.DefaultInputHandler.CommandBindings.Remove(
                enterCommandBinding
            );

        // 初始化输入框和输出框
        InitializeTextEditors();

        InitializeHtmlTemplate();
    }

    private void InitializeHtmlTemplate()
    {
        ResponseOutput.NavigateToString(
            @"
<html>
<head>
    <meta charset='utf-8'>
    <style>
        body {
            background-color: #1c1c1c;
            color: white;
            font-family: 微软雅黑, monospace;
            line-height: 1.6;
            margin: 0;
            padding: 10px;
            overflow-y: auto;
        }

        /* 消息容器 */
        .message-container {
            display: flex;
            flex-direction: column;
            align-items: flex-start; /* 默认左对齐 */
            margin: 10px 0;
        }

        /* 消息框 */
        .message {
            padding: 10px;
            border-radius: 8px;
            max-width: 90%;
            word-wrap: break-word;
        }

        /* 用户消息样式 */
        .user {
            background-color: #2d2d30;
            border: 1px solid #3e3e42;
            margin-left: auto;
            margin-right: 10px;
            text-align: left;
        }

        /* 系统消息样式 */
        .system {
            background-color: #3e3e42;
            border: 1px solid #5a5a5e;
            text-align: left;
            margin-left: 10px;
        }

        /* 角色名称 */
        .role {
            font-weight: bold;
            margin-bottom: 5px;
        }

        /* 时间戳样式 */
        .timestamp {
            font-size: 0.75em; /* 字体更小 */
            color: gray;
            margin-top: 5px;
            margin-left: 10px; /* 默认左对齐 */
            margin-right: 10px; /* 默认右边距 */
            text-align: left; /* 默认左对齐 */
        }

        /* 用户消息的时间戳右对齐 */
        .user + .timestamp {
            text-align: right;
            margin-left: auto; /* 自动调整左边距 */
            margin-right: 10px; /* 与框体右边距一致 */
        }

        pre {
            background-color: #2d2d30;
            color: #dcdcdc;
            padding: 10px;
            border-radius: 8px;
            overflow-x: auto;
        }

        code {
            font-family: 微软雅黑, monospace;
            background-color: #2d2d30;
            color: #dcdcdc;
            padding: 2px 4px;
            border-radius: 4px;
        }
    </style>
    <script>
        function scrollToBottom() {
            window.scrollTo(0, document.body.scrollHeight);
        }
        function replaceContent(responseId, newContent) {
            // 找到对应ID的消息容器div
            var messageContainer = document.getElementById(responseId);
            if (messageContainer) {
                // 在消息容器内找到类名为'content'的div
                var contentDiv = messageContainer.querySelector('.content');
                if (contentDiv) {
                    // 替换内容
                    contentDiv.innerHTML = newContent;
                }
            }
        }
    </script>
</head>
<body>
    <!-- 示例消息 
    <div class=""message-container"">
        <div class=""message user"">
            <div class=""role"">cent</div>
            <div>你好</div>
        </div>
        <div class=""timestamp"">2025-01-08 17:13:24</div>
    </div>

    <div class=""message-container"">
        <div class=""message system"">
            <div class=""role"">gpt-4o</div>
            <div>你好！有什么我可以帮助你的吗？如果你有关于Excel公式或C#代码的问题，请随时告诉我。</div>
        </div>
        <div class=""timestamp"">2025-01-08 17:13:25</div>
    </div>-->
</body>
</html>
                    "
        );
        // 加载本地聊天记录
        LoadChatHistory();
    }

    private void LoadChatHistory()
    {
        var chatRecord = new ChatHistoryManager();
        var chatHistory = chatRecord.LoadChatHistory();
        foreach (var message in chatHistory)
            LoadChatHistoryOutPut(message.Role, message.Message, message.IsUser, message.Timestamp);
    }

    private void InitializeTextEditors()
    {
        // 输入框默认文本
        PromptInput.Text = DefaultPromptText;
        PromptInput.Foreground = Brushes.Gray;
    }

    private void SendButton_Click(object sender, RoutedEventArgs e)
    {
        ProcessInput();
    }

    private void PromptInput_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            if ((e.KeyboardDevice.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                // 允许换行
                e.Handled = true; // 阻止默认行为
                PromptInput.Document.Insert(PromptInput.CaretOffset, Environment.NewLine);
            }
            else
            {
                // 阻止默认行为并发送消息
                e.Handled = true;
                SendButton_Click(SendButton, new RoutedEventArgs());
            }
        }
    }

    private void PromptInput_GotFocus(object sender, RoutedEventArgs e)
    {
        // 清空默认文本
        if (PromptInput.Text == DefaultPromptText)
        {
            PromptInput.Text = string.Empty;
            PromptInput.Foreground = Brushes.White;
        }
    }

    private void PromptInput_LostFocus(object sender, RoutedEventArgs e)
    {
        // 恢复默认文本
        if (string.IsNullOrWhiteSpace(PromptInput.Text))
        {
            PromptInput.Text = DefaultPromptText;
            PromptInput.Foreground = Brushes.Gray;
        }
    }

    private async void ProcessInput()
    {
        _apiKey = NumDesAddIn.ApiKey;
        _apiUrl = NumDesAddIn.ApiUrl;
        _apiModel = NumDesAddIn.ApiModel;

        var userInput = PromptInput.Document.Text.Trim();

        //新增当前单元格选中内容，输入的首字符为/时，识别
        if (userInput.StartsWith("###"))
        {
            var app = NumDesAddIn.App;
            var selectRange = app.Selection;
            var selectValue = selectRange.Value2;

            string selectValueStr = PubMetToExcel.ArrayToArrayStr(selectValue);

            userInput = selectValueStr + "," + userInput.Replace("###", "");
        }

        if (string.IsNullOrEmpty(userInput))
            return;

        try
        {
            object requestBody = null;
            if (_apiModel.Contains("gpt"))
                requestBody = CreateRequestBody(userInput);
            else if (_apiModel.Contains("deepseek")) requestBody = CreateRequestBodyDeepSeek(userInput);

            AppendToOutput(_userName, userInput, true);

            // 重置当前响应ID
            _currentResponseId = null;

            // 初始化流式消息
            var streamMessage = new ChatMessage
            {
                Role = _apiModel,
                Message = "",
                IsUser = false
            };
            await ChatApiClient.CallApiStreamAsync(
                requestBody,
                _apiKey,
                _apiUrl,
                chunk =>
                {
                    // 更新消息
                    streamMessage.Message += chunk;
                    AppendStreamingContent(chunk);
                },
                () =>
                {
                    var nowTime = DateTime.Now;
                    streamMessage.Timestamp = nowTime;

                    // 流结束时添加时间戳
                    Dispatcher.Invoke(() =>
                    {
                        var script = $@"
                        var container = document.getElementById('{_currentResponseId}');
                        var timestampDiv = container.querySelector('.timestamp');
                        timestampDiv.innerHTML = '{nowTime:yyyy-MM-dd HH:mm:ss}';";

                        ResponseOutput.InvokeScript("eval", script);

                    });
                });

            // 转换新消息为 HTML
            var htmlMessage = Markdown.ToHtml(streamMessage.Message);
            // **解码 HTML 实体**
            htmlMessage = HttpUtility.HtmlDecode(htmlMessage);

            // 保存完整消息
            streamMessage.Message = htmlMessage;
            new ChatHistoryManager().SaveChatMessage(streamMessage);

            ResponseOutput.InvokeScript("replaceContent",
                new object[] { _currentResponseId, htmlMessage });

        }
        catch (Exception ex)
        {
            AppendToOutput(_apiModel, $"调用 AI API 时出错：{ex.Message}", false);
        }

        // 清空输入框
        PromptInput.Document.Text = string.Empty;
        PromptInput.Focus(); // 自动聚焦到输入框
    }


    private void AppendStreamingContent(string chunk)
    {
        Dispatcher.Invoke(() =>
        {
            var doc = ResponseOutput.Document;
            if (doc == null) return;

            // 获取或创建消息容器
            if (string.IsNullOrEmpty(_currentResponseId))
            {
                _currentResponseId = $"msg-{DateTime.Now.Ticks}";

                var newMessage = $@"
                <div id='{_currentResponseId}' class='message-container'>
                    <div class='message system'>
                        <div class='role'>{_apiModel}</div>
                        <div class='content'</div>
                        <div class='content loading-dots' style='color:#888'>思考中……</div>
                    </div>
                    <div class='timestamp'></div>
                </div>";
                AppendRawHtml(newMessage);
            }

            // 更新内容
            var script = $@"
            var container = document.getElementById('{_currentResponseId}');
            var contentDiv = container.querySelector('.content');
            contentDiv.innerHTML += '{HttpUtility.JavaScriptStringEncode(chunk)}';
            scrollToBottom();";

            ResponseOutput.InvokeScript("eval", script);
        });
    }

    private void AppendRawHtml(string html)
    {
        var script = $@"
        document.body.insertAdjacentHTML('beforeend', '{HttpUtility.JavaScriptStringEncode(html)}');
        scrollToBottom();";

        ResponseOutput.InvokeScript("eval", script);
    }

    private object CreateRequestBody(string userInput)
    {
        _apiModel = NumDesAddIn.ApiModel;
        _sysContent = NumDesAddIn.ChatGptSysContentExcelAss;

        return new
        {
            model = _apiModel,
            messages = new[]
            {
                new { role = "system", content = _sysContent },
                new { role = "user", content = userInput }
            },
            max_tokens = 10000, // 最大生成的 token 数量
            temperature = 0.5, // 控制生成的随机性
            top_p = 1, // 核采样参数
            stream = true // 流式输出
        };
    }

    private object CreateRequestBodyDeepSeek(string userInput)
    {
        _apiModel = NumDesAddIn.ApiModel;
        _sysContent = NumDesAddIn.ChatGptSysContentExcelAss;

        return new
        {
            model = _apiModel,
            messages = new[]
            {
                new { content = "", role = "system" },
                new { content = userInput, role = "user" }
            },
            max_tokens = 10000,
            stream = true // 流式输出
        };
    }

    private void AppendToOutput(string role, string message, bool isUser, DateTime? timestamp = null)
    {
        Dispatcher.BeginInvoke(() =>
        {
            var doc = ResponseOutput.Document;

            if (doc != null)
            {
                // 使用反射获取 body 对象
                var body = doc.GetType()
                    .InvokeMember("body", BindingFlags.GetProperty, null, doc, null);

                if (body != null)
                {
                    // 获取当前的 innerHTML
                    var currentHtml = body.GetType()
                        .InvokeMember("innerHTML", BindingFlags.GetProperty, null, body, null)
                        ?.ToString();

                    // 转换新消息为 HTML
                    var htmlMessage = Markdown.ToHtml(
                        HttpUtility.HtmlEncode(message)
                    );

                    // **解码 HTML 实体**
                    htmlMessage = HttpUtility.HtmlDecode(htmlMessage);

                    // 如果未传递时间戳，则使用当前时间
                    var displayTimestamp = timestamp?.ToString("yyyy-MM-dd HH:mm:ss") ??
                                           DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                    // 生成消息 HTML
                    var messageHtml =
                        $@"
                        <div class='message-container'>
                            <div class='message {(isUser ? "user" : "system")}'>
                                <div class='role'>{role}</div>
                                <div>{htmlMessage}</div>
                            </div>
                            <div class='timestamp'>{displayTimestamp}</div>
                        </div>";

                    // 追加新消息到现有内容
                    var updatedHtml = currentHtml + messageHtml;

                    // 设置新的 innerHTML
                    body.GetType()
                        .InvokeMember(
                            "innerHTML",
                            BindingFlags.SetProperty,
                            null,
                            body,
                            new object[] { updatedHtml }
                        );

                    // 调用 JavaScript 函数 scrollToBottom
                    ResponseOutput.InvokeScript("scrollToBottom");

                    // 保存消息到本地文件
                    var chatRecord = new ChatHistoryManager();
                    chatRecord.SaveChatMessage(new ChatMessage
                    {
                        Role = role,
                        Message = htmlMessage,
                        IsUser = isUser,
                        Timestamp = DateTime.Now // 保存时间戳
                    });
                }
            }
        });
    }

    private void LoadChatHistoryOutPut(string role, string message, bool isUser, DateTime? timestamp = null)
    {
        Dispatcher.BeginInvoke(() =>
        {
            var doc = ResponseOutput.Document;

            if (doc != null)
            {
                // 使用反射获取 body 对象
                var body = doc.GetType()
                    .InvokeMember("body", BindingFlags.GetProperty, null, doc, null);

                if (body != null)
                {
                    // 获取当前的 innerHTML
                    var currentHtml = body.GetType()
                        .InvokeMember("innerHTML", BindingFlags.GetProperty, null, body, null)
                        ?.ToString();

                    // 生成消息 HTML
                    var messageHtml =
                        $@"
                        <div class='message-container'>
                            <div class='message {(isUser ? "user" : "system")}'>
                                <div class='role'>{role}</div>
                                <div>{message}</div>
                            </div>
                            <div class='timestamp'>{timestamp}</div>
                        </div>";

                    // 追加新消息到现有内容
                    var updatedHtml = currentHtml + messageHtml;

                    // 设置新的 innerHTML
                    body.GetType()
                        .InvokeMember(
                            "innerHTML",
                            BindingFlags.SetProperty,
                            null,
                            body,
                            new object[] { updatedHtml }
                        );

                    // 调用 JavaScript 函数 scrollToBottom
                    ResponseOutput.InvokeScript("scrollToBottom");
                }
            }
        });
    }
}