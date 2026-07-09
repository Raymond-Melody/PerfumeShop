using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

[ApiController]
[Route("api/flash-sale")]
public class FlashSaleController : ControllerBase
{
    private readonly IFlashSaleService _svc;
    public FlashSaleController(IFlashSaleService svc) => _svc = svc;

    /// <summary>GET /api/flash-sale — 当前进行中的秒杀活动(分页)</summary>
    [HttpGet]
    public async Task<IActionResult> GetActive([FromQuery] int page = 1, [FromQuery] int pageSize = 12)
    {
        var (items, total) = await _svc.GetActiveFlashSalesAsync(page, pageSize);
        return Ok(new { success = true, data = items, total, page, pageSize });
    }

    /// <summary>GET /api/flash-sale/upcoming — 即将开始的秒杀</summary>
    [HttpGet("upcoming")]
    public async Task<IActionResult> GetUpcoming([FromQuery] int top = 6)
    {
        var items = await _svc.GetUpcomingFlashSalesAsync(top);
        return Ok(new { success = true, data = items });
    }

    /// <summary>GET /api/flash-sale/{id} — 单个秒杀详情</summary>
    [HttpGet("{id}")]
    public async Task<IActionResult> GetDetail(int id)
    {
        var item = await _svc.GetFlashSaleByIdAsync(id);
        if (item == null) return NotFound(new { success = false, message = "秒杀活动不存在" });
        return Ok(new { success = true, data = item });
    }

    /// <summary>POST /api/flash-sale/{id}/purchase — 抢购</summary>
    [HttpPost("{id}/purchase")]
    public async Task<IActionResult> Purchase(int id, [FromBody] PurchaseRequest req)
    {
        if (req.UserId <= 0) return BadRequest(new { success = false, message = "请先登录" });
        var result = await _svc.PurchaseAsync(id, req.UserId, req.Quantity > 0 ? req.Quantity : 1);
        return Ok(new { success = result.Success, message = result.Message, orderId = result.OrderId });
    }

    // ========== 管理后台端点 ==========

    /// <summary>GET /api/flash-sale/admin/stats — 秒杀统计</summary>
    [HttpGet("admin/stats")]
    public async Task<IActionResult> GetStats()
    {
        var stats = await _svc.GetAdminStatsAsync();
        return Ok(new { success = true, data = stats });
    }

    /// <summary>POST /api/flash-sale/admin/{id}/toggle — 切换启用状态</summary>
    [HttpPost("admin/{id}/toggle")]
    public async Task<IActionResult> Toggle(int id)
    {
        var ok = await _svc.ToggleActiveAsync(id);
        return Ok(new { success = ok, message = ok ? "状态已切换" : "操作失败" });
    }

    /// <summary>DELETE /api/flash-sale/admin/{id} — 删除秒杀活动</summary>
    [HttpDelete("admin/{id}")]
    public async Task<IActionResult> Delete(int id)
    {
        var ok = await _svc.DeleteFlashSaleAsync(id);
        return Ok(new { success = ok, message = ok ? "已删除" : "删除失败" });
    }
}

public class PurchaseRequest
{
    public int UserId { get; set; }
    public int Quantity { get; set; } = 1;
}
