using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Core.AI;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 情感分析 API — 对应 V18 ai-service/sentiment_analyzer.py
/// </summary>
[ApiController]
[Route("api/v2/sentiment")]
public class SentimentController : ControllerBase
{
    private readonly ISentimentAnalyzer _analyzer;

    public SentimentController(ISentimentAnalyzer analyzer)
    {
        _analyzer = analyzer;
    }

    /// <summary>POST /api/v2/sentiment/analyze — 单条文本情感分析</summary>
    [HttpPost("analyze")]
    public IActionResult Analyze([FromBody] SentimentRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.Text))
            return BadRequest(new { success = false, message = "文本不能为空" });

        var result = _analyzer.Analyze(req.Text);
        return Ok(new
        {
            success = true,
            data = new
            {
                score = result.Score,
                label = result.Label,
                confidence = result.Confidence,
                keywords = result.Keywords,
                summary = result.Summary
            }
        });
    }

    /// <summary>POST /api/v2/sentiment/batch — 批量文本情感分析</summary>
    [HttpPost("batch")]
    public IActionResult Batch([FromBody] SentimentBatchRequest req)
    {
        if (req.Texts == null || req.Texts.Count == 0)
            return BadRequest(new { success = false, message = "文本列表不能为空" });

        if (req.Texts.Count > 50)
            return BadRequest(new { success = false, message = "单次最多分析50条文本" });

        var results = _analyzer.BatchAnalyze(req.Texts);
        return Ok(new
        {
            success = true,
            count = results.Count,
            data = results.Select(r => new
            {
                score = r.Score,
                label = r.Label,
                confidence = r.Confidence,
                keywords = r.Keywords,
                summary = r.Summary
            })
        });
    }
}

public class SentimentRequest
{
    public string Text { get; set; } = "";
}

public class SentimentBatchRequest
{
    public List<string> Texts { get; set; } = new();
}
