using PerfumeShop.Core.AI;

namespace PerfumeShop.IntegrationTests.AI;

/// <summary>
/// 智能客服引擎测试 — 意图识别和多轮对话
/// </summary>
public class ChatbotEngineTests
{
    private readonly IChatbotEngine _engine = new ChatbotEngine();

    [Theory]
    [InlineData("你好")]
    [InlineData("hello")]
    [InlineData("嗨")]
    [InlineData("thanks")]
    [InlineData("bye")]
    [InlineData("再见")]
    public void ProcessMessage_Greetings_ReturnsGreetingCategory(string message)
    {
        var response = _engine.ProcessMessage(message);

        Assert.NotNull(response);
        Assert.Equal("greeting", response.Category);
        Assert.NotEmpty(response.Reply);
        Assert.False(response.NeedsHuman);
    }

    [Fact]
    public void ProcessMessage_FallbackUnknown_ReturnsNeedsHuman()
    {
        var response = _engine.ProcessMessage("你们公司的股票怎么样呀");

        Assert.Equal("fallback", response.Category);
        Assert.True(response.NeedsHuman);
        Assert.NotEmpty(response.Suggestions);
    }

    [Fact]
    public void ProcessMessage_AlwaysReturnsNonEmptyReply()
    {
        var inputs = new[]
        {
            "你好", "退货", "定制香水", "积分", "未知问题xyz", "",
            "如何查询订单状态", "运费多少", "香水保质期", "会员权益"
        };

        foreach (var input in inputs)
        {
            var response = _engine.ProcessMessage(input);
            Assert.NotNull(response);
            Assert.NotEmpty(response.Reply);
            Assert.NotNull(response.Category);
        }
    }

    [Fact]
    public void ProcessMessage_ReturnRefund_ContainsRelevantContent()
    {
        var response = _engine.ProcessMessage("我想申请退货退款");

        // Should match return/refund pattern
        Assert.Contains("退", response.Reply);
        Assert.Equal("order", response.Category);
        Assert.NotEmpty(response.Suggestions);
    }

    [Fact]
    public void ProcessMessage_CustomizeKeywords_ContainsCustomizeContent()
    {
        var response = _engine.ProcessMessage("前调中调后调怎么调配");

        Assert.Equal("customize", response.Category);
        Assert.Contains("定制", response.Reply);
    }

    [Fact]
    public void ProcessMessage_SessionHistory_Works()
    {
        var sessionId = $"test-{Guid.NewGuid()}";

        var greet = _engine.ProcessMessage("你好", sessionId);
        Assert.Equal("greeting", greet.Category);

        // Second message in same session
        var followUp = _engine.ProcessMessage("我想申请退货退款", sessionId);
        Assert.Equal("order", followUp.Category);
    }

    [Fact]
    public void GetFaqList_ReturnsFiveCategories()
    {
        var faqList = _engine.GetFaqList();

        Assert.NotNull(faqList);
        Assert.Equal(5, faqList.Count);

        var categories = faqList.Select(f => f.Category).ToHashSet();
        Assert.Contains("order", categories);
        Assert.Contains("shipping", categories);
        Assert.Contains("product", categories);
        Assert.Contains("customize", categories);
        Assert.Contains("account", categories);
    }

    [Fact]
    public void GetFaqList_EachCategoryHasQuestions()
    {
        var faqList = _engine.GetFaqList();

        foreach (var cat in faqList)
        {
            Assert.NotEmpty(cat.Questions);
            Assert.NotEmpty(cat.Name);
        }
    }

    [Fact]
    public void ProcessMessage_FaqExactMatch_ReturnsCorrectCategory()
    {
        var response = _engine.ProcessMessage("定制流程是怎样的？");

        Assert.Equal("customize", response.Category);
    }

    [Fact]
    public void ProcessMessage_FaqShippingExact_ReturnsShipping()
    {
        var response = _engine.ProcessMessage("配送范围和费用？");

        Assert.Equal("shipping", response.Category);
    }

    [Fact]
    public void ProcessMessage_FaqProductExact_ReturnsProduct()
    {
        var response = _engine.ProcessMessage("香水能保存多久？");

        Assert.Equal("product", response.Category);
    }

    [Fact]
    public void ProcessMessage_FaqAccountExact_ReturnsAccount()
    {
        var response = _engine.ProcessMessage("如何获取积分？");

        Assert.Equal("account", response.Category);
    }

    [Fact]
    public void Engine_IsReady()
    {
        Assert.True(_engine.IsReady);
    }
}
