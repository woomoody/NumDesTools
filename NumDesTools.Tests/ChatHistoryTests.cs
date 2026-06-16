using NumDesTools.AI;

namespace NumDesTools.Tests;

/// <summary>
/// 通过 NUMDES_CHATHISTORY_TEST_DB 环境变量注入临时 DB 路径，
/// 让 ChatHistoryManager 默认构造函数使用测试 DB。
/// </summary>
public class ChatHistoryTests : IDisposable
{
    private readonly string _dbPath;
    private readonly ChatHistoryManager _manager;

    public ChatHistoryTests()
    {
        _dbPath = Path.GetTempFileName() + ".db";
        Environment.SetEnvironmentVariable(
            ChatHistoryManager.TestDbEnvVar,
            $"Data Source={_dbPath}"
        );
        _manager = new ChatHistoryManager();
    }

    public void Dispose()
    {
        Environment.SetEnvironmentVariable(ChatHistoryManager.TestDbEnvVar, null);
        try
        {
            if (File.Exists(_dbPath))
                File.Delete(_dbPath);
        }
        catch { }
    }

    [Fact]
    public void ListSessionsWithPreview_NoSessions_ReturnsEmpty()
    {
        var result = _manager.ListSessionsWithPreview(isAgent: false);

        Assert.Empty(result);
    }

    [Fact]
    public async Task ListSessionsWithPreview_WithAgentMessages_ReturnsSessions()
    {
        var sessionId = Guid.NewGuid().ToString();
        var userMsg = "用户问题内容";

        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "assistant",
                Message = "AI回答",
                IsUser = false,
                Timestamp = DateTime.Now,
                SessionId = sessionId,
                IsAgent = true,
            }
        );
        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "assistant",
                Message = "AI回答2",
                IsUser = false,
                Timestamp = DateTime.Now.AddSeconds(1),
                SessionId = sessionId,
                IsAgent = true,
            }
        );
        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "user",
                Message = userMsg,
                IsUser = true,
                Timestamp = DateTime.Now.AddSeconds(2),
                SessionId = sessionId,
                IsAgent = true,
            }
        );

        var result = _manager.ListSessionsWithPreview(isAgent: true);

        Assert.Single(result);
        Assert.Contains(userMsg, result[0].Preview);
    }

    [Fact]
    public async Task ListSessionsWithPreview_AgentAndChatIsolated()
    {
        var chatSessionId = Guid.NewGuid().ToString();
        var agentSessionId = Guid.NewGuid().ToString();

        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "user",
                Message = "chat消息",
                IsUser = true,
                Timestamp = DateTime.Now,
                SessionId = chatSessionId,
                IsAgent = false,
            }
        );
        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "user",
                Message = "agent消息",
                IsUser = true,
                Timestamp = DateTime.Now,
                SessionId = agentSessionId,
                IsAgent = true,
            }
        );

        var chatResult = _manager.ListSessionsWithPreview(isAgent: false);
        var agentResult = _manager.ListSessionsWithPreview(isAgent: true);

        Assert.Single(chatResult);
        Assert.Equal(chatSessionId, chatResult[0].SessionId);

        Assert.Single(agentResult);
        Assert.Equal(agentSessionId, agentResult[0].SessionId);
    }

    [Fact]
    public async Task ListSessionsWithPreview_PreviewTruncated()
    {
        var sessionId = Guid.NewGuid().ToString();
        var longMessage = new string('X', 80); // 超过 40 字

        await _manager.SaveChatMessageAsync(
            new ChatMessage
            {
                Role = "user",
                Message = longMessage,
                IsUser = true,
                Timestamp = DateTime.Now,
                SessionId = sessionId,
                IsAgent = true,
            }
        );

        var result = _manager.ListSessionsWithPreview(isAgent: true);

        Assert.Single(result);
        Assert.True(
            result[0].Preview.Length <= 40,
            $"Preview 长度 {result[0].Preview.Length} 超过 40"
        );
    }
}
