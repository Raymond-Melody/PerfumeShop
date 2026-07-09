using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/group-buy")]
public class GroupBuyController : ControllerBase
{
    private readonly IGroupBuyService _svc;
    public GroupBuyController(IGroupBuyService svc) => _svc = svc;

    /// <summary>GET /api/group-buy — 进行中的拼团计划</summary>
    [HttpGet]
    public async Task<IActionResult> GetActivePlans()
    {
        var plans = await _svc.GetActivePlansAsync();
        return Ok(new { success = true, data = plans });
    }

    /// <summary>GET /api/group-buy/{planId}/groups — 某计划下可加入的团</summary>
    [HttpGet("{planId}/groups")]
    public async Task<IActionResult> GetOpenGroups(int planId)
    {
        var groups = await _svc.GetOpenGroupsAsync(planId);
        return Ok(new { success = true, data = groups });
    }

    /// <summary>GET /api/group-buy/{planId}/stats — 计划统计</summary>
    [HttpGet("{planId}/stats")]
    public async Task<IActionResult> GetStats(int planId)
    {
        var stats = await _svc.GetPlanStatsAsync(planId);
        return Ok(new { success = true, data = stats });
    }

    /// <summary>GET /api/group-buy/group/{groupId} — 团详情</summary>
    [HttpGet("group/{groupId}")]
    public async Task<IActionResult> GetGroupDetail(int groupId)
    {
        var detail = await _svc.GetGroupDetailAsync(groupId);
        if (detail == null) return NotFound(new { success = false, message = "团不存在" });
        return Ok(new { success = true, data = detail });
    }

    /// <summary>POST /api/group-buy/start — 发起新团</summary>
    [HttpPost("start")]
    public async Task<IActionResult> StartGroup([FromBody] GroupBuyActionRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var result = await _svc.StartGroupAsync(req.PlanId, req.UserId);
        return Ok(new { success = result.Success, message = result.Message, groupId = result.GroupId, groupSn = result.GroupSn });
    }

    /// <summary>POST /api/group-buy/join — 加入已有团</summary>
    [HttpPost("join")]
    public async Task<IActionResult> JoinGroup([FromBody] GroupBuyJoinRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var result = await _svc.JoinGroupAsync(req.GroupId, req.UserId);
        return Ok(new { success = result.Success, message = result.Message, isGroupComplete = result.IsGroupComplete });
    }

    // ========== 管理后台 ==========

    /// <summary>POST /api/group-buy/admin/{planId}/toggle — 切换计划状态</summary>
    [HttpPost("admin/{planId}/toggle")]
    public async Task<IActionResult> Toggle(int planId)
    {
        var ok = await _svc.TogglePlanActiveAsync(planId);
        return Ok(new { success = ok, message = ok ? "状态已切换" : "操作失败" });
    }

    /// <summary>DELETE /api/group-buy/admin/{planId} — 删除拼团计划</summary>
    [HttpDelete("admin/{planId}")]
    public async Task<IActionResult> Delete(int planId)
    {
        var ok = await _svc.DeletePlanAsync(planId);
        return Ok(new { success = ok, message = ok ? "已删除" : "删除失败" });
    }
}

public class GroupBuyActionRequest
{
    public int PlanId { get; set; }
    public int UserId { get; set; }
}

public class GroupBuyJoinRequest
{
    public int GroupId { get; set; }
    public int UserId { get; set; }
}
