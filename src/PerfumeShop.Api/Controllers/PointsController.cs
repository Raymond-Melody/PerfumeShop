using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/points")]
public class PointsController : ControllerBase
{
    private readonly IPointsService _svc;
    public PointsController(IPointsService svc) => _svc = svc;

    /// <summary>GET /api/points/balance/{userId} — 积分余额</summary>
    [HttpGet("balance/{userId}")]
    public async Task<IActionResult> GetBalance(int userId)
    {
        var balance = await _svc.GetBalanceAsync(userId);
        return Ok(new { success = true, data = balance });
    }

    /// <summary>GET /api/points/ledger/{userId} — 积分流水</summary>
    [HttpGet("ledger/{userId}")]
    public async Task<IActionResult> GetLedger(int userId, [FromQuery] int page = 1, [FromQuery] int pageSize = 20)
    {
        var (items, total) = await _svc.GetLedgerAsync(userId, page, pageSize);
        return Ok(new { success = true, data = items, total, page, pageSize });
    }

    /// <summary>POST /api/points/award — 发放积分</summary>
    [HttpPost("award")]
    public async Task<IActionResult> Award([FromBody] PointsActionRequest req)
    {
        var points = await _svc.AwardPointsAsync(req.UserId, req.Points, req.Source ?? "manual", req.Description, req.ReferenceId);
        return Ok(new { success = true, message = $"已发放{points}积分", points });
    }

    /// <summary>POST /api/points/deduct — 扣减积分</summary>
    [HttpPost("deduct")]
    public async Task<IActionResult> Deduct([FromBody] PointsActionRequest req)
    {
        var ok = await _svc.DeductPointsAsync(req.UserId, req.Points, req.Source ?? "manual", req.Description, req.ReferenceId);
        return Ok(new { success = ok, message = ok ? "积分扣减成功" : "积分不足或操作失败" });
    }

    /// <summary>GET /api/points/redemptions — 兑换商城商品</summary>
    [HttpGet("redemptions")]
    public async Task<IActionResult> GetRedemptions()
    {
        var items = await _svc.GetRedemptionItemsAsync();
        return Ok(new { success = true, data = items });
    }

    /// <summary>POST /api/points/redeem — 积分兑换</summary>
    [HttpPost("redeem")]
    public async Task<IActionResult> Redeem([FromBody] RedeemRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var result = await _svc.RedeemAsync(req.UserId, req.RedemptionId);
        return Ok(new { success = result.Success, message = result.Message, pointsRemaining = result.PointsRemaining });
    }

    /// <summary>GET /api/points/rules — 积分规则</summary>
    [HttpGet("rules")]
    public async Task<IActionResult> GetRules()
    {
        var rules = await _svc.GetRulesAsync();
        return Ok(new { success = true, data = rules });
    }
}

public class PointsActionRequest
{
    public int UserId { get; set; }
    public int Points { get; set; }
    public string? Source { get; set; }
    public string? Description { get; set; }
    public int? ReferenceId { get; set; }
}

public class RedeemRequest
{
    public int UserId { get; set; }
    public int RedemptionId { get; set; }
}
