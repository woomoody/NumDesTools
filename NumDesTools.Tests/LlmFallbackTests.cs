using NumDesTools.AI;
using Xunit;

namespace NumDesTools.Tests;

public class LlmFallbackTests
{
    // ── Test 1: 识别 deepseek 内容过滤错误 ───────────────────────────────────

    [Fact]
    public void IsContentFilterError_DeepseekInappropriateContent_ReturnsTrue()
    {
        var ex = new HttpRequestException(
            "litellm.BadRequestError: DashscopeException - Output data may contain " +
            "inappropriate content. For details, see: https://help.aliyun.com/zh/model-studio/error");

        Assert.True(LiteLlmClient.IsContentFilterError(ex));
    }

    // ── Test 2: 普通网络错误不触发 fallback ───────────────────────────────────

    [Fact]
    public void IsContentFilterError_NetworkError_ReturnsFalse()
    {
        var ex = new HttpRequestException("Connection refused");
        Assert.False(LiteLlmClient.IsContentFilterError(ex));
    }

    // ── Test 3: 认证错误不触发 fallback ──────────────────────────────────────

    [Fact]
    public void IsContentFilterError_AuthError_ReturnsFalse()
    {
        var ex = new HttpRequestException("401 Unauthorized");
        Assert.False(LiteLlmClient.IsContentFilterError(ex));
    }

    // ── Test 4: null message 不崩溃 ──────────────────────────────────────────

    [Fact]
    public void IsContentFilterError_NullMessage_ReturnsFalse()
    {
        var ex = new Exception();
        Assert.False(LiteLlmClient.IsContentFilterError(ex));
    }
}
