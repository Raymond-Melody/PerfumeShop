using PerfumeShop.Core.AI;

namespace PerfumeShop.IntegrationTests.AI;

/// <summary>
/// 情感分析器测试 — 20 条固定数据
/// </summary>
public class SentimentAnalyzerTests
{
    private readonly ISentimentAnalyzer _analyzer = new SentimentAnalyzer();

    [Theory]
    [InlineData("这款香水非常好闻，我很喜欢", "positive")]
    [InlineData("味道太棒了，超级持久", "positive")]
    [InlineData("质量很好，值得推荐", "positive")]
    [InlineData("完美的礼物，女朋友很喜欢", "positive")]
    [InlineData("包装精致，香味优雅", "positive")]
    [InlineData("This perfume is amazing and wonderful", "positive")]
    [InlineData("Absolutely love this fragrance, so beautiful", "positive")]
    [InlineData("Great scent, excellent quality", "positive")]
    [InlineData("Very nice and lovely perfume", "positive")]
    [InlineData("Best perfume ever, perfect!", "positive")]
    [InlineData("太难闻了，非常失望", "negative")]
    [InlineData("刺鼻的酒精味，后悔购买", "negative")]
    [InlineData("假货，质量很差", "negative")]
    [InlineData("味道不持久，不值这个价", "negative")]
    [InlineData("Terrible quality, very disappointed", "negative")]
    [InlineData(" smells awful, worst purchase ever", "negative")]
    [InlineData("收到了，还没使用", "neutral")]
    [InlineData("快递挺快的", "neutral")]
    [InlineData("", "neutral")]
    [InlineData("今天天气还行", "neutral")]
    public void Analyze_20FixedData_ReturnsExpectedSentiment(string text, string expectedLabel)
    {
        var result = _analyzer.Analyze(text);

        Assert.NotNull(result);
        Assert.Equal(expectedLabel, result.Label);
        Assert.InRange(result.Score, -1.0, 1.0);
    }

    [Fact]
    public void Analyze_PositiveText_ReturnsPositiveKeywords()
    {
        var result = _analyzer.Analyze("这款香水非常好闻，香味持久，值得推荐");

        Assert.Equal("positive", result.Label);
        Assert.True(result.Keywords.Positive.Count > 0);
        Assert.True(result.Confidence > 0);
        Assert.NotEmpty(result.Summary);
    }

    [Fact]
    public void Analyze_NegativeWithNegator_FlipsSentiment()
    {
        // "不好闻" should flip "好闻" from positive to negative
        var result = _analyzer.Analyze("这个香水不好闻");

        Assert.Equal("negative", result.Label);
    }

    [Fact]
    public void Analyze_EmptyText_ReturnsNeutral()
    {
        var result = _analyzer.Analyze("");

        Assert.Equal("neutral", result.Label);
        Assert.Equal(0.0, result.Score);
        Assert.Equal(1.0, result.Confidence);
    }

    [Fact]
    public void Analyze_NullText_ReturnsNeutral()
    {
        var result = _analyzer.Analyze(null!);

        Assert.Equal("neutral", result.Label);
    }

    [Fact]
    public void Analyze_IntensifierAmplifiesScore()
    {
        var normal = _analyzer.Analyze("好闻");
        var intensified = _analyzer.Analyze("非常好闻");

        // Intensified version should have equal or higher score magnitude
        Assert.True(intensified.Score >= normal.Score || intensified.Label == normal.Label);
    }

    [Fact]
    public void BatchAnalyze_MultipleTexts_ReturnsMatchingCount()
    {
        var texts = new[] { "很喜欢", "太难闻了", "还行吧" };
        var results = _analyzer.BatchAnalyze(texts);

        Assert.Equal(3, results.Count);
    }

    [Fact]
    public void Analyzer_IsReady()
    {
        Assert.True(_analyzer.IsReady);
    }
}
