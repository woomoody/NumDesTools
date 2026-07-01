using System.Net;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NumDesTools.Scanner;

/// <summary>
/// 飞书项目 MCP HTTP 调用封装，对应 Python 版本的 mcp_call()。
/// </summary>
public static class FeishuMcpClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(120) };

    public static string McpToken { get; set; } = string.Empty;
    public static string McpUrl { get; set; } = "https://project.feishu.cn/mcp_server/v1";
    public static string ProjectKey { get; set; } = string.Empty;

    // ponytail: 手写指数退避，不引入 Polly，3次已够用
    private static readonly int[] RetryDelaysMs = [1000, 2000, 4000];

    /// <summary>
    /// 带指数退避重试的 HTTP 发送。只对网络异常和 5xx 重试，4xx 直接抛出。
    /// 使用 requestFactory 在每次重试时重新构造 HttpRequestMessage（避免 disposed 复用）。
    /// </summary>
    private static async Task<HttpResponseMessage> SendWithRetryAsync(
        Func<HttpRequestMessage> requestFactory,
        CancellationToken ct = default
    )
    {
        for (var attempt = 0; ; attempt++)
        {
            using var request = requestFactory();
            HttpResponseMessage response;
            try
            {
                response = await Http.SendAsync(request, ct);
            }
            catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException)
            {
                if (attempt >= RetryDelaysMs.Length)
                    throw;
                await Task.Delay(RetryDelaysMs[attempt], ct);
                continue;
            }

            // 4xx → 鉴权失败/请求错误，不重试直接抛
            if ((int)response.StatusCode is >= 400 and < 500)
            {
                response.Dispose();
                throw new HttpRequestException(
                    $"飞书 MCP 请求失败：HTTP {(int)response.StatusCode} {response.ReasonPhrase}"
                );
            }

            // 5xx → 服务端错误，重试
            if (response.StatusCode >= HttpStatusCode.InternalServerError)
            {
                response.Dispose();
                if (attempt >= RetryDelaysMs.Length)
                    throw new HttpRequestException(
                        $"飞书 MCP 服务端错误（已重试 {RetryDelaysMs.Length} 次）：HTTP {(int)response.StatusCode}"
                    );
                await Task.Delay(RetryDelaysMs[attempt], ct);
                continue;
            }

            return response;
        }
    }

    /// <summary>
    /// 调用飞书 MCP 工具，返回解析后的 JToken（对象或数组），失败抛出异常。
    /// </summary>
    public static async Task<JToken> CallAsync(string toolName, object arguments)
    {
        var body = JsonConvert.SerializeObject(
            new
            {
                jsonrpc = "2.0",
                method = "tools/call",
                @params = new { name = toolName, arguments },
                id = 1,
            }
        );

        HttpRequestMessage Factory() =>
            new HttpRequestMessage(HttpMethod.Post, McpUrl)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
                Headers = { { "X-Mcp-Token", McpToken } },
            };

        using var response = await SendWithRetryAsync(Factory);
        var raw = await response.Content.ReadAsStringAsync();
        var root = JObject.Parse(raw);

        var result = root["result"];
        if (result?["isError"]?.Value<bool>() == true)
        {
            var msgs = result["content"]?.Select(c => c["text"]?.ToString()) ?? [];
            throw new InvalidOperationException("MCP error: " + string.Join(" | ", msgs));
        }

        foreach (var content in result?["content"] ?? new JArray())
        {
            var text = content["text"]?.ToString();
            if (string.IsNullOrWhiteSpace(text))
                continue;
            try
            {
                return JToken.Parse(text);
            }
            catch
            { /* 不是 JSON，继续 */
            }
        }

        return new JObject();
    }

    /// <summary>按 MQL 查询工作项，返回原始 JToken。</summary>
    public static Task<JToken> SearchByMqlAsync(string mql) =>
        CallAsync("search_by_mql", new { project_key = ProjectKey, mql });

    /// <summary>获取工作项评论列表。</summary>
    public static Task<JToken> ListCommentsAsync(string workItemId) =>
        CallAsync(
            "list_workitem_comments",
            new { project_key = ProjectKey, work_item_id = workItemId }
        );

    /// <summary>
    /// 更新工作项指定字段（multi-text 类型直接传纯文本）。不触发推送通知。
    /// story 用 field_d9a4cd（需求内容），issue 用 description（缺陷描述）。
    /// </summary>
    public static Task<JToken> UpdateTextField(
        string workItemId,
        string workItemTypeKey,
        string fieldKey,
        string text
    ) =>
        CallAsync(
            "update_field",
            new
            {
                project_key = ProjectKey,
                work_item_id = workItemId,
                work_item_type_key = workItemTypeKey,
                fields = new[] { new { field_key = fieldKey, field_value = text } },
            }
        );

    public static string StoryContentFieldKey => "field_d9a4cd";
    public static string IssueDescFieldKey => "description";

    /// <summary>
    /// 拉取工作项指定字段的当前文本值，用于写入前保留人工原文。
    /// </summary>
    public static async Task<string> GetCurrentFieldValueAsync(string workItemId, string fieldKey)
    {
        try
        {
            var result = await CallAsync(
                "get_workitem_brief",
                new
                {
                    project_key = ProjectKey,
                    work_item_id = workItemId,
                    fields = new[] { fieldKey },
                }
            );
            var fieldList = result["work_item_fields"] as Newtonsoft.Json.Linq.JArray;
            if (fieldList != null)
                foreach (var f in fieldList)
                    if (f["key"]?.ToString() == fieldKey)
                        return f["value"]?.ToString() ?? "";
            return result[fieldKey]?.ToString() ?? "";
        }
        catch { }
        return "";
    }
}
