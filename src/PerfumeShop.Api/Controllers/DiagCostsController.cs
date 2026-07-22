using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Controllers;

/// <summary>
/// 成本诊断 API — 对应 V18 api/diag_costs.asp
/// </summary>
[ApiController]
[Route("api/v2/diag")]
public class DiagCostsController : ControllerBase
{
    private readonly PerfumeShopContext _db;

    public DiagCostsController(PerfumeShopContext db)
    {
        _db = db;
    }

    /// <summary>GET /api/v2/diag/costs — 成本数据诊断</summary>
    [HttpGet("costs")]
    public async Task<IActionResult> GetCostDiag()
    {
        try
        {
            var totalProducts = await _db.Products.CountAsync();
            var productsWithPrice = await _db.Products.CountAsync(p => p.BasePrice > 0);

            // Revenue summary from orders (non-cancelled)
            var revenueOrders = await _db.Orders
                .Where(o => o.Status != "Cancelled" && o.Status != "Pending")
                .Select(o => new { o.TotalAmount, CostAmt = (double?)o.CostAmount })
                .ToListAsync();

            var revenue = revenueOrders.Sum(o => (double)o.TotalAmount);
            var totalCost = revenueOrders.Sum(o => o.CostAmt ?? 0);
            var ordersWithCost = revenueOrders.Count(o => (o.CostAmt ?? 0) > 0);

            // Product samples
            var productSamples = await _db.Products
                .Take(5)
                .Select(p => new
                {
                    id = p.ProductId,
                    name = p.ProductName,
                    basePrice = (double)p.BasePrice,
                    type = p.ProductType
                })
                .ToListAsync();

            return Ok(new
            {
                totalProducts,
                productsWithPrice,
                productSamples,
                financeSummary = new
                {
                    revenue = Math.Round(revenue, 2),
                    totalCost = Math.Round(totalCost, 2)
                },
                ordersWithCost,
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
