namespace PerfumeShop.Core.AI;

/// <summary>
/// 情感分析结果
/// </summary>
public class SentimentResult
{
    public double Score { get; set; }
    public string Label { get; set; } = "";
    public double Confidence { get; set; }
    public SentimentKeywords Keywords { get; set; } = new();
    public string Summary { get; set; } = "";
}

public class SentimentKeywords
{
    public List<string> Positive { get; set; } = new();
    public List<string> Negative { get; set; } = new();
}

/// <summary>
/// 情感分析器接口
/// </summary>
public interface ISentimentAnalyzer
{
    bool IsReady { get; }
    SentimentResult Analyze(string text);
    List<SentimentResult> BatchAnalyze(IEnumerable<string> texts);
}
