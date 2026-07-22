using System.Text.RegularExpressions;
using JiebaNet.Segmenter;

namespace PerfumeShop.Core.AI;

/// <summary>
/// 基于词典的情感分析器（移植自 ai-service/sentiment_analyzer.py）
/// 使用 jieba.NET 进行中文分词，内置情感词库 + 否定词处理 + 程度副词
/// </summary>
public class SentimentAnalyzer : ISentimentAnalyzer
{
    private static readonly HashSet<string> PositiveWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "喜欢", "不错", "好闻", "香", "推荐", "满意", "棒", "赞", "完美", "惊喜",
        "优雅", "清新", "持久", "高级", "自然", "舒服", "柔和", "温柔", "迷人",
        "好评", "超值", "回购", "性价比", "经典", "独特", "精致", "好", "爱",
        "nice", "love", "good", "great", "beautiful", "excellent", "perfect",
        "wonderful", "amazing", "fantastic", "best", "lovely"
    };

    private static readonly HashSet<string> NegativeWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "失望", "不好", "难闻", "刺鼻", "不持久", "太淡", "太浓", "不值", "差",
        "后悔", "过敏", "头晕", "恶心", "酒精", "假货", "劣质", "粗糙", "呛",
        "bad", "poor", "terrible", "awful", "worst", "horrible", "dislike",
        "cheap", "fake", "weak", "strong", "headache", "allergy"
    };

    private static readonly Dictionary<string, double> Intensifiers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["非常"] = 2.0, ["特别"] = 2.0, ["超级"] = 2.5, ["极其"] = 2.5, ["太"] = 1.8,
        ["很"] = 1.5, ["挺"] = 1.3, ["有点"] = 0.7, ["稍微"] = 0.6, ["略微"] = 0.6,
        ["very"] = 2.0, ["extremely"] = 2.5, ["really"] = 1.8, ["quite"] = 1.5,
        ["so"] = 1.5, ["absolutely"] = 2.5
    };

    private static readonly HashSet<string> Negators = new(StringComparer.OrdinalIgnoreCase)
    {
        "不", "没", "无", "非", "not", "no", "don't", "doesn't", "didn't"
    };

    private static readonly Regex ChinesePattern = new(@"[\u4e00-\u9fff]", RegexOptions.Compiled);
    private static readonly Regex CleanPattern = new(@"[^\u4e00-\u9fff\w\s]", RegexOptions.Compiled);

    private readonly JiebaSegmenter? _segmenter;

    public bool IsReady { get; }

    public SentimentAnalyzer()
    {
        try
        {
            _segmenter = new JiebaSegmenter();
            IsReady = true;
        }
        catch
        {
            IsReady = true; // 降级到空格分词也能工作
        }
    }

    public SentimentResult Analyze(string text)
    {
        text = (text ?? "").Trim();
        if (string.IsNullOrEmpty(text))
        {
            return new SentimentResult
            {
                Label = "neutral", Score = 0.0, Confidence = 1.0,
                Keywords = new SentimentKeywords(), Summary = ""
            };
        }

        var words = Tokenize(text);

        double posCount = 0, negCount = 0;
        var posKeywords = new List<string>();
        var negKeywords = new List<string>();

        for (int i = 0; i < words.Count; i++)
        {
            var word = words[i];
            bool negatorBefore = i > 0 && Negators.Contains(words[i - 1]);
            double intensifier = 1.0;
            if (i > 0 && Intensifiers.TryGetValue(words[i - 1], out var intens))
                intensifier = intens;

            if (PositiveWords.Contains(word))
            {
                if (negatorBefore)
                {
                    negCount += intensifier;
                    negKeywords.Add($"不{word}");
                }
                else
                {
                    posCount += intensifier;
                    posKeywords.Add(word);
                }
            }
            else if (NegativeWords.Contains(word))
            {
                if (negatorBefore)
                {
                    posCount += intensifier * 0.7;
                    posKeywords.Add($"不{word}");
                }
                else
                {
                    negCount += intensifier;
                    negKeywords.Add(word);
                }
            }
        }

        double total = posCount + negCount;
        double score = total == 0 ? 0.0 : (posCount - negCount) / total;

        string label = score > 0.2 ? "positive" : score < -0.2 ? "negative" : "neutral";

        int keywordCount = posKeywords.Count + negKeywords.Count;
        double confidence = keywordCount > 0 ? Math.Min(1.0, keywordCount / 5.0) : 0.3;

        return new SentimentResult
        {
            Score = Math.Round(score, 3),
            Label = label,
            Confidence = Math.Round(confidence, 3),
            Keywords = new SentimentKeywords
            {
                Positive = posKeywords.Take(5).ToList(),
                Negative = negKeywords.Take(5).ToList()
            },
            Summary = GenerateSummary(label, posKeywords.Take(5).ToList(), negKeywords.Take(5).ToList())
        };
    }

    public List<SentimentResult> BatchAnalyze(IEnumerable<string> texts)
    {
        return texts.Select(Analyze).ToList();
    }

    private List<string> Tokenize(string text)
    {
        text = CleanPattern.Replace(text, " ");

        if (_segmenter != null && ChinesePattern.IsMatch(text))
        {
            return _segmenter.Cut(text)
                .Where(w => !string.IsNullOrWhiteSpace(w))
                .Select(w => w.Trim())
                .ToList();
        }

        return text.ToLower().Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
    }

    private static string GenerateSummary(string sentiment, List<string> posWords, List<string> negWords)
    {
        if (sentiment == "positive")
            return posWords.Count > 0
                ? $"用户评价积极，关键词: {string.Join(", ", posWords.Take(3))}"
                : "用户评价偏正面";
        if (sentiment == "negative")
            return negWords.Count > 0
                ? $"用户评价消极，关键词: {string.Join(", ", negWords.Take(3))}"
                : "用户评价偏负面";
        return "用户评价中性";
    }
}
