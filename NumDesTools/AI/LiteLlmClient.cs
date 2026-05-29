using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NumDesTools.AI;

/// <summary>
/// OpenAI-兼容 HTTP 客户端（LiteLLM 后端）。
/// </summary>
public class LiteLlmClient : ILlmClient
{
    private static readonly HttpClient Client = new() { Timeout = TimeSpan.FromMinutes(3) };

    private static object BuildRequestBody(
        string model,
        IEnumerable<object> messages,
        bool stream = true
    ) =>
        new
        {
            model,
            messages,
            max_tokens = 10000,
            temperature = 0.5,
            stream,
        };

    public async Task<string> CallAsync(
        string model,
        string systemContent,
        string userContent,
        string apiKey,
        string apiUrl,
        CancellationToken ct = default
    )
    {
        if (string.IsNullOrEmpty(apiKey))
            throw new ArgumentException("API 密钥不能为空。");

        var messages = new object[]
        {
            new { role = "system", content = systemContent ?? "" },
            new { role = "user", content = userContent },
        };
        var body = BuildRequestBody(model, messages, stream: false);

        using var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
        request.Content = new StringContent(
            JsonConvert.SerializeObject(body),
            Encoding.UTF8,
            "application/json"
        );
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        var response = await Client.SendAsync(request, ct);
        var responseContent = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
            throw new Exception($"API 调用失败 {response.StatusCode}：{responseContent}");

        var json = JObject.Parse(responseContent);
        return json["choices"]?[0]?["message"]?["content"]?.ToString() ?? "";
    }

    public async Task CallStreamAsync(
        string model,
        IReadOnlyList<object> messages,
        string apiKey,
        string apiUrl,
        System.Action<string> onChunkReceived,
        System.Action? onCompleted = null,
        CancellationToken ct = default
    )
    {
        var body = BuildRequestBody(model, messages);
        using var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
        request.Content = new StringContent(
            JsonConvert.SerializeObject(body),
            Encoding.UTF8,
            "application/json"
        );
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

        using var response = await Client.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            ct
        );
        response.EnsureSuccessStatusCode();

        using var stream = await response.Content.ReadAsStreamAsync();
        using var reader = new StreamReader(stream);

        while (!reader.EndOfStream)
        {
            var line = await reader.ReadLineAsync();
            if (line?.StartsWith("data: ") == true)
                ProcessLine(line["data: ".Length..], onChunkReceived);
        }
        onCompleted?.Invoke();
    }

    public async Task<List<string>> FetchModelsAsync(string apiKey, string apiUrl)
    {
        try
        {
            var modelsUrl = apiUrl.Replace("/chat/completions", "/models");
            using var request = new HttpRequestMessage(HttpMethod.Get, modelsUrl);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
            using var response = await Client.SendAsync(request);
            if (!response.IsSuccessStatusCode)
                return [];
            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
            return json["data"]
                    ?.Select(m => m["id"]?.ToString())
                    .Where(id => !string.IsNullOrEmpty(id))
                    .ToList()
                ?? [];
        }
        catch (Exception ex)
        {
            PluginLog.Write($"FetchModels 失败: {ex.Message}");
            return [];
        }
    }

    private static void ProcessLine(string json, Action<string> handler)
    {
        if (json == "[DONE]")
            return;
        try
        {
            var obj = JObject.Parse(json);
            var content = obj["choices"]?[0]?["delta"]?["content"]?.ToString();
            if (!string.IsNullOrEmpty(content))
                handler(content);
        }
        catch (Exception ex)
        {
            PluginLog.Write($"解析失败: {ex.Message}");
        }
    }
}
