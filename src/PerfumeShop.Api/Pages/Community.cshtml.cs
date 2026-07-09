using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "MarketingPage")]
public class CommunityModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public CommunityModel(PerfumeShopContext db) => _db = db;

    public List<CommunityPost> Posts { get; set; } = new();
    public int TotalPages { get; set; }
    public int CurrentPage { get; set; }
    public string? Filter { get; set; }
    private const int PageSize = 12;

    public async Task OnGetAsync(string? type, int page = 1)
    {
        Filter = type;
        CurrentPage = page;

        var query = _db.CommunityPosts
            .AsNoTracking()
            .Where(p => p.IsPublic && p.IsActive);

        if (!string.IsNullOrEmpty(type))
            query = query.Where(p => p.PostType == type);

        var total = await query.CountAsync();
        TotalPages = (int)Math.Ceiling(total / (double)PageSize);

        Posts = await query
            .OrderByDescending(p => p.IsPinned)
            .ThenByDescending(p => p.CreatedAt)
            .Skip((page - 1) * PageSize)
            .Take(PageSize)
            .ToListAsync();
    }
}
