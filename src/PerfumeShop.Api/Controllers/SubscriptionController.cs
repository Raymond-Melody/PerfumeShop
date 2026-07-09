using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/subscription")]
public class SubscriptionController : ControllerBase
{
    private readonly ISubscriptionService _svc;
    public SubscriptionController(ISubscriptionService svc) => _svc = svc;

    /// <summary>GET /api/subscription/plans — 活跃订阅计划列表</summary>
    [HttpGet("plans")]
    public async Task<IActionResult> GetPlans()
    {
        var plans = await _svc.GetActivePlansAsync();
        return Ok(new { success = true, data = plans });
    }

    /// <summary>GET /api/subscription/user/{userId} — 用户当前订阅</summary>
    [HttpGet("user/{userId}")]
    public async Task<IActionResult> GetUserSubscription(int userId)
    {
        var sub = await _svc.GetUserSubscriptionAsync(userId);
        if (sub == null) return Ok(new { success = true, data = (object?)null, message = "暂无订阅" });
        return Ok(new { success = true, data = sub });
    }

    /// <summary>POST /api/subscription/subscribe — 创建订阅</summary>
    [HttpPost("subscribe")]
    public async Task<IActionResult> Subscribe([FromBody] SubscribeRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var result = await _svc.SubscribeAsync(req.UserId, req.PlanId, req.AutoRenew);
        return Ok(new { success = result.Success, message = result.Message, subscriptionId = result.SubscriptionId });
    }

    /// <summary>POST /api/subscription/{id}/cancel — 取消订阅</summary>
    [HttpPost("{id}/cancel")]
    public async Task<IActionResult> Cancel(int id, [FromBody] UserAction req)
    {
        var ok = await _svc.CancelSubscriptionAsync(id, req.UserId);
        return Ok(new { success = ok, message = ok ? "已取消订阅" : "取消失败" });
    }

    /// <summary>POST /api/subscription/{id}/auto-renew — 切换自动续费</summary>
    [HttpPost("{id}/auto-renew")]
    public async Task<IActionResult> ToggleAutoRenew(int id, [FromBody] AutoRenewRequest req)
    {
        var ok = await _svc.ToggleAutoRenewAsync(id, req.UserId, req.AutoRenew);
        return Ok(new { success = ok, message = ok ? "设置已更新" : "操作失败" });
    }

    /// <summary>GET /api/subscription/{id}/deliveries — 配送历史</summary>
    [HttpGet("{id}/deliveries")]
    public async Task<IActionResult> GetDeliveries(int id)
    {
        var deliveries = await _svc.GetDeliveryHistoryAsync(id);
        return Ok(new { success = true, data = deliveries });
    }

    // ========== 管理后台 ==========

    /// <summary>GET /api/subscription/admin/plans — 全部订阅计划</summary>
    [HttpGet("admin/plans")]
    public async Task<IActionResult> GetAllPlans()
    {
        var plans = await _svc.GetAllPlansAsync();
        return Ok(new { success = true, data = plans });
    }

    /// <summary>POST /api/subscription/admin/{planId}/toggle — 切换状态</summary>
    [HttpPost("admin/{planId}/toggle")]
    public async Task<IActionResult> Toggle(int planId)
    {
        var ok = await _svc.TogglePlanActiveAsync(planId);
        return Ok(new { success = ok, message = ok ? "状态已切换" : "操作失败" });
    }

    /// <summary>DELETE /api/subscription/admin/{planId} — 删除订阅计划</summary>
    [HttpDelete("admin/{planId}")]
    public async Task<IActionResult> Delete(int planId)
    {
        var ok = await _svc.DeletePlanAsync(planId);
        return Ok(new { success = ok, message = ok ? "已删除" : "删除失败" });
    }
}

public class SubscribeRequest
{
    public int UserId { get; set; }
    public int PlanId { get; set; }
    public bool AutoRenew { get; set; } = true;
}

public class AutoRenewRequest
{
    public int UserId { get; set; }
    public bool AutoRenew { get; set; }
}

public class UserAction
{
    public int UserId { get; set; }
}
