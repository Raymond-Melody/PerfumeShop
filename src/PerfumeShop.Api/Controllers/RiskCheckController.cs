using Microsoft.AspNetCore.Mvc;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 风险检查 API — 对应 V18 includes/api_guard.asp
/// </summary>
[ApiController]
[Route("api/v2/risk")]
public class RiskCheckController : ControllerBase
{
    /// <summary>POST /api/v2/risk/check — 请求风险评估</summary>
    [HttpPost("check")]
    public IActionResult Check([FromBody] RiskCheckRequest req)
    {
        var riskLevel = "low";
        var reasons = new List<string>();

        // IP frequency check (simplified)
        if (!string.IsNullOrEmpty(req.Ip))
        {
            if (req.Ip.StartsWith("10.") || req.Ip.StartsWith("192.168.") || req.Ip == "127.0.0.1")
            {
                reasons.Add("internal_ip");
            }
        }

        // Amount-based risk
        if (req.Amount > 10000)
        {
            riskLevel = "high";
            reasons.Add("high_amount");
        }
        else if (req.Amount > 5000)
        {
            riskLevel = "medium";
            reasons.Add("medium_amount");
        }

        // New user risk
        if (req.AccountAgeDays < 7)
        {
            if (riskLevel == "low") riskLevel = "medium";
            reasons.Add("new_account");
        }

        // Multiple rapid requests
        if (req.RequestsPerMinute > 30)
        {
            riskLevel = "high";
            reasons.Add("rate_exceeded");
        }

        return Ok(new
        {
            success = true,
            data = new
            {
                riskLevel,
                score = riskLevel switch { "high" => 0.8, "medium" => 0.5, _ => 0.2 },
                reasons,
                action = riskLevel == "high" ? "block" : riskLevel == "medium" ? "review" : "allow",
                timestamp = DateTime.UtcNow
            }
        });
    }
}

public class RiskCheckRequest
{
    public string Ip { get; set; } = "";
    public decimal Amount { get; set; }
    public int AccountAgeDays { get; set; }
    public int RequestsPerMinute { get; set; }
    public string Action { get; set; } = "";
}
