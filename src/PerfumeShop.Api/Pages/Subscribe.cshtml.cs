using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "MarketingPage")]
public class SubscribeModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public SubscribeModel(PerfumeShopContext db) => _db = db;

    public List<SubscriptionPlan> Plans { get; set; } = new();

    public async Task OnGetAsync()
    {
        Plans = await _db.SubscriptionPlans
            .AsNoTracking()
            .Where(p => p.IsActive)
            .OrderBy(p => p.SortOrder)
            .ToListAsync();
    }
}
