using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "MarketingPage")]
public class GroupBuyModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public GroupBuyModel(PerfumeShopContext db) => _db = db;

    public List<GroupBuyPlan> Plans { get; set; } = new();
    public Dictionary<int, Product> Products { get; set; } = new();

    public async Task OnGetAsync()
    {
        Plans = await _db.GroupBuyPlans
            .AsNoTracking()
            .Where(p => p.IsActive)
            .OrderBy(p => p.SortOrder)
            .ToListAsync();

        var productIds = Plans.Select(p => p.ProductId).Distinct().ToList();
        if (productIds.Any())
            Products = await _db.Products
                .AsNoTracking()
                .Where(p => productIds.Contains(p.ProductId))
                .ToDictionaryAsync(p => p.ProductId);
    }
}
