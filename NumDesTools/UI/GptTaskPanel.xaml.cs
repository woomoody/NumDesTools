using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Newtonsoft.Json;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using UserControl = System.Windows.Controls.UserControl;

namespace NumDesTools.UI
{
    /// <summary>
    /// GptTaskPanel.xaml 的交互逻辑
    /// </summary>
    public partial class GptTaskPanel : UserControl
    {
        private readonly string _apiKey;

        public GptTaskPanel()
        {
            InitializeComponent();
            InitializeHtmlTemplate();
            _apiKey = Environment.GetEnvironmentVariable("API_KEY");
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
                    font-family: Consolas, monospace;
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
                    font-family: Consolas, monospace;
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


        private async void ProcessInput()
        {
            string userInput = PromptInput.Text.Trim();
            if (string.IsNullOrEmpty(userInput))
                return;

            PromptInput.Clear();
            GeneratingHint.Visibility = Visibility.Visible;

            try
            {
                string response = await CallChatGptApi(userInput, _apiKey);
                AppendToOutput("用户", userInput, isUser: true);
                AppendToOutput("系统", response, isUser: false);
            }
            catch (Exception ex)
            {
                AppendToOutput("系统", $"调用 GPT API 时出错：{ex.Message}", isUser: false);
            }
            finally
            {
                GeneratingHint.Visibility = Visibility.Collapsed;
            }
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

        private async Task<string> CallChatGptApi(string prompt, string apiKey)
        {
            string apiUrl = "https://api.openai.com/v1/chat/completions";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                var requestBody = new
                {
                    model = "gpt-4o",
                    messages = new[]
                    {
                    new { role = "system", content = "You are an assistant." },
                    new { role = "user", content = prompt }
                },
                    max_tokens = 2048
                };

                string jsonBody = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    dynamic jsonResponse = JsonConvert.DeserializeObject(responseContent);
                    return jsonResponse.choices[0].message.content.ToString();
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"API 调用失败，状态码：{response.StatusCode}，错误信息：{errorContent}");
                }
            }
        }
    }
}