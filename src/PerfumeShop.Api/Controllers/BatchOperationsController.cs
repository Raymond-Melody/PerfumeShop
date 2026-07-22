using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 批量操作 API — V19 M4-C
/// 对齐 V18 api/batch_operations.asp
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class BatchOperationsController : ControllerBase
{
    private readonly PerfumeShopContext _db;

    public BatchOperationsController(PerfumeShopContext db) => _db = db;

    /// <summary>批量发货</summary>
    [HttpPost("ship")]
    public async Task<IActionResult> BatchShip([FromBody] BatchShipRequest request)
    {
        if (request.Ids == null || request.Ids.Length == 0)
            return BadRequest(new { message = "请选择要发货的订单" });

        int success = 0, fail = 0;
        foreach (var id in request.Ids)
        {
            var order = await _db.Orders.FirstOrDefaultAsync(o => o.OrderId == id);
            if (order != null && order.Status == "Paid")
            {
                order.Status = "Shipped";
                order.TrackingNumber = request.TrackingNo;
                order.ShippedAt = DateTime.Now;
                order.UpdatedAt = DateTime.Now;
                success++;
            }
            else fail++;
        }
        await _db.SaveChangesAsync();
        return Ok(new { action = "batch_ship", total = request.Ids.Length, successCount = success, failCount = fail });
    }

    /// <summary>批量取消</summary>
    [HttpPost("cancel")]
    public async Task<IActionResult> BatchCancel([FromBody] BatchIdsRequest request)
    {
        if (request.Ids == null || request.Ids.Length == 0)
            return BadRequest(new { message = "请选择要取消的订单" });

        int success = 0, fail = 0;
        foreach (var id in request.Ids)
        {
            var order = await _db.Orders.FirstOrDefaultAsync(o => o.OrderId == id);
            if (order != null && (order.Status == "Pending" || order.Status == "Paid"))
            {
                order.Status = "Cancelled";
                order.UpdatedAt = DateTime.Now;
                success++;
            }
            else fail++;
        }
        await _db.SaveChangesAsync();
        return Ok(new { action = "batch_cancel", total = request.Ids.Length, successCount = success, failCount = fail });
    }

    /// <summary>批量上下架</summary>
    [HttpPost("status")]
    public async Task<IActionResult> BatchStatus([FromBody] BatchStatusRequest request)
    {
        if (request.Ids == null || request.Ids.Length == 0)
            return BadRequest(new { message = "请选择要操作的商品" });

        bool activate = request.Action == "list";
        int success = 0, fail = 0;
        foreach (var id in request.Ids)
        {
            var product = await _db.Products.FirstOrDefaultAsync(p => p.ProductId == id);
            if (product != null)
            {
                product.IsActive = activate;
                success++;
            }
            else fail++;
        }
        await _db.SaveChangesAsync();
        return Ok(new { action = request.Action, total = request.Ids.Length, successCount = success, failCount = fail });
    }
}

public class BatchShipRequest
{
    public int[] Ids { get; set; } = Array.Empty<int>();
    public string? TrackingNo { get; set; }
}

public class BatchIdsRequest
{
    public int[] Ids { get; set; } = Array.Empty<int>();
}

public class BatchStatusRequest
{
    public int[] Ids { get; set; } = Array.Empty<int>();
    public string Action { get; set; } = "list"; // list=上架, unlist=下架
}
