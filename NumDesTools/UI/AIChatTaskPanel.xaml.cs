using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Brushes = System.Windows.Media.Brushes;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI
{
    public partial class AiChatTaskPanel
    {
        private readonly string _apiKey;
        private readonly string _apiUrl;
        private readonly string _apiModel;
        private readonly string _userName = Environment.UserName;
        private readonly string _sysContent;

        private const string DefaultPromptText = "Enter发送，Shift + Enter换行"; // 输入框默认文本

        public AiChatTaskPanel()
        {
            InitializeComponent();
            _apiKey = NumDesAddIn.ApiKey;
            _apiUrl = NumDesAddIn.ApiUrl;
            _apiModel = NumDesAddIn.ApiModel;
            _sysContent = NumDesAddIn.ChatGptSysContentExcelAss;

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
            ResponseOutput.NavigateToString(@"
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {
                    background-color: #1e1e1e;
                    color: white;
                    font-family: 微软雅黑, monospace;
                    line-height: 1.6;
                    margin: 0;
                    padding: 10px;
                }
                .message {
                    margin: 10px 0;
                    padding: 10px;
                    border-radius: 8px;
                    max-width: 90%;
                    word-wrap: break-word;
                }
                .user {
                    background-color: #2d2d30;
                    border: 1px solid #3e3e42;
                    text-align: right;
                    margin-left: 10px;
                }
                .system {
                    background-color: #3e3e42;
                    border: 1px solid #5a5a5e;
                    text-align: left;
                    margin-left: 10px;
                }
                .role {
                    font-weight: bold;
                    margin-bottom: 5px;
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
        </head>
        <body></body>
        </html>
        ");
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

        private void ProcessInput()
        {
            string userInput = PromptInput.Document.Text.Trim();
            if (string.IsNullOrEmpty(userInput))
                return;

            try
            {
                var requestBody = CreateRequestBody(userInput);

                string response = Task.Run(
                    () => ChatGptApiClient.CallApiAsync(requestBody, _apiKey, _apiUrl)
                ).Result;

                AppendToOutput(_userName, userInput, isUser: true);
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

        private object CreateRequestBody(string prompt)
        {
            return new
            {
                model = _apiModel,
                messages = new[]
                {
                    new { role = "system", content = _sysContent },
                    new { role = "user", content = prompt }
                },
                max_tokens = 2048
            };
        }

        private void AppendToOutput(string role, string message, bool isUser)
        {
            Dispatcher.Invoke(() =>
            {
                dynamic doc = ResponseOutput.Document;
                dynamic body = doc?.body;

                if (body != null)
                {
                    // 将 Markdown 转换为 HTML
                    string htmlMessage = Markdig.Markdown.ToHtml(message);

                    string messageHtml = $@"
                    <div class='message {(isUser ? "user" : "system")}'>
                        <div class='role'>{role}</div>
                        <div>{htmlMessage}</div>
                    </div>";

                    body.innerHTML += messageHtml;
                }
            });
        }



    }
}
