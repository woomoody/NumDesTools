using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace NumDesTools.UI
{
    /// <summary>
    /// GptTaskPanel.xaml 的交互逻辑
    /// </summary>
    public partial class GptTaskPanel
    {
        private readonly string _apiKey;
        private readonly string _userName = Environment.UserName;
        private readonly string _sysName = "gpt-4o";
        private readonly string _sysContent;

        public GptTaskPanel()
        {
            InitializeComponent();
            InitializeHtmlTemplate();
            _apiKey = NumDesAddIn.ChatGptApiKey;
            _sysContent = NumDesAddIn.ChatGptSysContentExcelAss;
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

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            ProcessInput();
        }

        private void PromptInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SendButton_Click(SendButton, new RoutedEventArgs());
            }
        }


        private  void ProcessInput()
        {
            string userInput = PromptInput.Text.Trim();
            if (string.IsNullOrEmpty(userInput))
                return;

            PromptInput.Clear();
            try
            {
                var requestBody = CreateRequestBody(userInput);

                string response = Task.Run(() => ChatGptApiClient.CallApiAsync(requestBody, _apiKey)).Result;

                AppendToOutput(_userName, userInput, isUser: true);
                AppendToOutput(_sysName, response, isUser: false);
            }
            catch (Exception ex)
            {
                AppendToOutput(_sysName, $"调用 GPT API 时出错：{ex.Message}", isUser: false);
            }
        }



        //Gpt配置参数
        private object CreateRequestBody(string prompt)
        {
            //model role是保留字段，不能自定义修改，可以修改content内容
            return new
            {
                model = "gpt-4o",
                messages = new[]
                {
                    new { role = "system", content = _sysContent},
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