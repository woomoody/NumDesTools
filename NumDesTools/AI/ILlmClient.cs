using System.Threading;
using System.Threading.Tasks;

namespace NumDesTools.AI;

/// <summary>
/// LLM 客户端抽象——用于测试替换或切换不同后端。
/// </summary>
public interface ILlmClient
{
    Task<string> CallAsync(
        string model,
        string systemContent,
        string userContent,
        string apiKey,
        string apiUrl,
        CancellationToken ct = default
    );

    Task CallStreamAsync(
        string model,
        IReadOnlyList<object> messages,
        string apiKey,
        string apiUrl,
        System.Action<string> onChunkReceived,
        System.Action? onCompleted = null,
        CancellationToken ct = default
    );

    Task<List<string>> FetchModelsAsync(string apiKey, string apiUrl);
}
