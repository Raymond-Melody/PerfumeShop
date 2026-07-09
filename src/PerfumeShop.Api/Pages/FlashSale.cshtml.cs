using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "MarketingPage")]
public class FlashSaleModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public FlashSaleModel(PerfumeShopContext db) => _db = db;

    public List<Data.Models.FlashSale> Sales { get; set; } = new();
    public Dictionary<int, Product> Products { get; set; } = new();

    public async Task OnGetAsync()
    {
        var now = DateTime.Now;
        Sales = await _db.FlashSales
            .AsNoTracking()
            .Where(f => f.IsActive && f.StartTime <= now && f.EndTime >= now)
            .OrderBy(f => f.SortOrder)
            .ToListAsync();

        var productIds = Sales.Select(f => f.ProductId).Distinct().ToList();
        if (productIds.Any())
            Products = await _db.Products.AsNoTracking().Where(p => productIds.Contains(p.ProductId)).ToDictionaryAsync(p => p.ProductId);
    }
}
