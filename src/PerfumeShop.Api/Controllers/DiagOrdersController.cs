using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 订单诊断 API — 对应 V18 api/diag_orders.asp
/// </summary>
[ApiController]
[Route("api/v2/diag")]
public class DiagOrdersController : ControllerBase
{
    private readonly PerfumeShopContext _db;

    public DiagOrdersController(PerfumeShopContext db)
    {
        _db = db;
    }

    /// <summary>GET /api/v2/diag/orders — 订单数据诊断</summary>
    [HttpGet("orders")]
    public async Task<IActionResult> GetOrderDiag()
    {
        try
        {
            // Order status distribution
            var orderStatuses = await _db.Orders
                .GroupBy(o => o.Status ?? "Unknown")
                .Select(g => new { status = g.Key, count = g.Count() })
                .ToListAsync();

            var totalOrders = await _db.Orders.CountAsync();

            return Ok(new
            {
                totalOrders,
                orderStatuses,
                generatedAt = DateTime.UtcNow
            });
        }
        catch (Exception ex)
        {
            return Ok(new
            {
                status = "error",
                message = ex.Message,
                generatedAt = DateTime.UtcNow
            });
        }
    }
}
