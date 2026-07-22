using Microsoft.AspNetCore.Mvc;
using PerfumeShop.Data.Interfaces;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 生产工单 API — V19 M4-C
/// 对齐 V18 api/sync_production_orders.asp, api/fix_production_status.asp
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class ProductionController : ControllerBase
{
    private readonly IProductionRepository _repo;

    public ProductionController(IProductionRepository repo) => _repo = repo;

    /// <summary>生产工单列表（分页+筛选）</summary>
    [HttpGet]
    public async Task<IActionResult> GetOrders(
        [FromQuery] string? status = null,
        [FromQuery] string? search = null,
        [FromQuery] int page = 1,
        [FromQuery] int pageSize = 20)
    {
        var (items, total) = await _repo.GetProductionOrdersAsync(status, search, page, pageSize);
        return Ok(new { items, total, page, pageSize });
    }

    /// <summary>工单详情</summary>
    [HttpGet("{id:int}")]
    public async Task<IActionResult> GetOrder(int id)
    {
        var po = await _repo.GetProductionOrderDetailAsync(id);
        if (po == null) return NotFound(new { message = "工单不存在" });

        var logs = await _repo.GetProductionLogsAsync(id);
        return Ok(new
        {
            po.ProductionId, po.OrderId, po.DetailId, po.WorkOrderNo,
            po.BottleIndex, po.TotalBottles, po.Status, po.Priority,
            po.RecipeId, po.RecipeName, po.AssignedTo, po.Notes,
            po.StartedAt, po.CompletedAt, po.CreatedAt, po.UpdatedAt,
            po.Qcnotes, po.QcpassedAt, po.ShippedOutAt, po.WarehouseInAt,
            logs
        });
    }

    /// <summary>状态更新</summary>
    [HttpPut("{id:int}/status")]
    public async Task<IActionResult> UpdateStatus(int id, [FromBody] StatusRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.NewStatus))
            return BadRequest(new { message = "状态不能为空" });

        var ok = await _repo.UpdateProductionStatusAsync(id, request.NewStatus, request.Operator);
        return ok
            ? Ok(new { message = $"状态已更新为 {request.NewStatus}" })
            : NotFound(new { message = "工单不存在" });
    }

    /// <summary>同步已付款订单→生产工单（幂等）</summary>
    [HttpPost("sync-orders")]
    public async Task<IActionResult> SyncOrders(CancellationToken ct)
    {
        var (synced, errors, message) = await _repo.SyncProductionOrdersAsync(ct);
        return Ok(new { success = errors == 0, synced, errors, message });
    }

    /// <summary>修复生产工单状态（中文→英文迁移）</summary>
    [HttpPost("{id:int}/fix-status")]
    public async Task<IActionResult> FixStatus(CancellationToken ct)
    {
        var (updated, message) = await _repo.FixProductionStatusAsync(ct);
        return Ok(new { success = true, updated, message });
    }

    /// <summary>质检记录列表</summary>
    [HttpGet("qc")]
    public async Task<IActionResult> GetQualityChecks(
        [FromQuery] string? status = null,
        [FromQuery] int page = 1,
        [FromQuery] int pageSize = 20)
    {
        var (items, total) = await _repo.GetQualityChecksAsync(status, page, pageSize);
        return Ok(new { items, total, page, pageSize });
    }

    /// <summary>生产报表数据</summary>
    [HttpGet("report")]
    public async Task<IActionResult> GetReport(
        [FromQuery] DateTime? startDate = null,
        [FromQuery] DateTime? endDate = null)
    {
        var start = startDate ?? DateTime.Now.AddMonths(-1);
        var end = endDate ?? DateTime.Now;
        var data = await _repo.GetProductionReportDataAsync(start, end);

        // 按日期聚合产量
        var grouped = data.GroupBy(p => p.CreatedAt!.Value.Date)
            .Select(g => new
            {
                date = g.Key.ToString("yyyy-MM-dd"),
                count = g.Count(),
                completed = g.Count(p => p.Status == "Completed"),
                passRate = g.Count(p => p.Status == "Completed") * 100.0 / Math.Max(g.Count(), 1)
            }).ToList();

        return Ok(new { data = grouped, startDate = start.ToString("yyyy-MM-dd"), endDate = end.ToString("yyyy-MM-dd") });
    }
}

public class StatusRequest
{
    public string NewStatus { get; set; } = "";
    public string? Operator { get; set; }
}
