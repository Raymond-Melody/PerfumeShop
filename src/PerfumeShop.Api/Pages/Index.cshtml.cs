using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "HomePage")]
public class IndexModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public IndexModel(PerfumeShopContext db) => _db = db;

    public List<Product> FeaturedProducts { get; set; } = new();
    public List<FlashSale> ActiveFlashSales { get; set; } = new();
    public Dictionary<int, Product> FlashSaleProducts { get; set; } = new();

    public async Task OnGetAsync()
    {
        // 热门商品 (最新8件在售)
        FeaturedProducts = await _db.Products
            .AsNoTracking()
            .Where(p => p.IsActive == true)
            .OrderByDescending(p => p.CreatedAt)
            .Take(8)
            .ToListAsync();

        // 活跃秒杀
        var now = DateTime.Now;
        ActiveFlashSales = await _db.FlashSales
            .AsNoTracking()
            .Where(f => f.IsActive && f.StartTime <= now && f.EndTime >= now)
            .OrderBy(f => f.SortOrder)
            .Take(4)
            .ToListAsync();

        // 补充商品信息
        var productIds = ActiveFlashSales.Select(f => f.ProductId).Distinct().ToList();
        if (productIds.Any())
        {
            FlashSaleProducts = await _db.Products
                .AsNoTracking()
                .Where(p => productIds.Contains(p.ProductId))
                .ToDictionaryAsync(p => p.ProductId);
        }
    }
}
