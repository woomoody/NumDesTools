using System.Web;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Markdig;
using Brushes = System.Windows.Media.Brushes;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI
{
    public partial class AiChatTaskPanel
    {
        private string _apiKey;
        private string _apiUrl;
        private string _apiModel;
        private string _sysContent;

        private readonly string _userName = Environment.UserName;

        private const string DefaultPromptText =
            "Enter发送，Shift + Enter换行，聊天框内容复制使用右键\n首字输入###会把当前选择单元格值一并输入"; // 输入框默认文本

        public AiChatTaskPanel()
        {
            InitializeComponent();

            // 禁用 AvalonEdit 默认的 Enter 键行为
            var enterCommandBinding =
                PromptInput.TextArea.DefaultInputHandler.CommandBindings.FirstOrDefault(binding =>
                    binding.Command == EditingCommands.EnterParagraphBreak
                );
            if (enterCommandBinding != null)
            {
                PromptInput.TextArea.DefaultInputHandler.CommandBindings.Remove(
                    enterCommandBinding
                );
            }

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
            {
                LoadChatHistoryOutPut(message.Role, message.Message, message.IsUser , message.Timestamp);
            }
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

            string userInput = PromptInput.Document.Text.Trim();

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
                AppendToOutput(_userName, userInput, isUser: true);

                object requestBody = null;
                if (_apiModel.Contains("gpt"))
                {
                    requestBody = CreateRequestBody(userInput);
                }
                else if (_apiModel.Contains("deepseek"))
                {
                    requestBody = CreateRequestBodyDeepSeek(userInput);
                }

                string response = await ChatApiClient.CallApiStreamAsync(requestBody, _apiKey, _apiUrl);

                AppendToOutput(_apiModel, response, isUser: false);
            }
            catch (Exception ex)
            {
                AppendToOutput(_apiModel, $"调用 AI API 时出错：{ex.Message}", isUser: false);
            }

            // 清空输入框
            PromptInput.Document.Text = string.Empty;
            PromptInput.Focus(); // 自动聚焦到输入框
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
                max_tokens = 10000,     // 最大生成的 token 数量
                temperature = 0.5,     // 控制生成的随机性
                top_p = 1,             // 核采样参数
                stream = true          // 流式输出
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
                stream = true          // 流式输出
            };
        }

        private void AppendToOutput(string role, string message, bool isUser , DateTime? timestamp = null)
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
                        string currentHtml = body.GetType()
                            .InvokeMember("innerHTML", BindingFlags.GetProperty, null, body, null)
                            ?.ToString();

                        // 转换新消息为 HTML
                        string htmlMessage = Markdown.ToHtml(
                            HttpUtility.HtmlEncode(message)
                        );

                        // **解码 HTML 实体**
                        htmlMessage = HttpUtility.HtmlDecode(htmlMessage);

                        // 如果未传递时间戳，则使用当前时间
                        string displayTimestamp = timestamp?.ToString("yyyy-MM-dd HH:mm:ss") ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        // 生成消息 HTML
                        string messageHtml =
                            $@"
                        <div class='message-container'>
                            <div class='message {(isUser ? "user" : "system")}'>
                                <div class='role'>{role}</div>
                                <div>{htmlMessage}</div>
                            </div>
                            <div class='timestamp'>{displayTimestamp}</div>
                        </div>";

                        // 追加新消息到现有内容
                        string updatedHtml = currentHtml + messageHtml;

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
                            Message = message,
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
                        string currentHtml = body.GetType()
                            .InvokeMember("innerHTML", BindingFlags.GetProperty, null, body, null)
                            ?.ToString();

                        // 转换新消息为 HTML
                        string htmlMessage = Markdown.ToHtml(
                            HttpUtility.HtmlEncode(message)
                        );

                        // **解码 HTML 实体**
                        htmlMessage = HttpUtility.HtmlDecode(htmlMessage);

                        // 如果未传递时间戳，则使用当前时间
                        string displayTimestamp = timestamp?.ToString("yyyy-MM-dd HH:mm:ss") ?? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                        // 生成消息 HTML
                        string messageHtml =
                            $@"
                        <div class='message-container'>
                            <div class='message {(isUser ? "user" : "system")}'>
                                <div class='role'>{role}</div>
                                <div>{htmlMessage}</div>
                            </div>
                            <div class='timestamp'>{displayTimestamp}</div>
                        </div>";

                        // 追加新消息到现有内容
                        string updatedHtml = currentHtml + messageHtml;

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
}
