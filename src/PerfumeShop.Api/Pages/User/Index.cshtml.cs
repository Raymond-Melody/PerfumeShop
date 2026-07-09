using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class IndexModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public IndexModel(PerfumeShopContext db) => _db = db;

    public int OrderCount { get; set; }
    public int FavoriteCount { get; set; }
    public int Points { get; set; }
    public List<Order> RecentOrders { get; set; } = new();

    public async Task OnGetAsync()
    {
        // 简单演示: 获取统计数据 (实际应用从 cookie/session 获取 userId)
        OrderCount = await _db.Orders.CountAsync();
        FavoriteCount = await _db.UserFavorites.CountAsync();
        Points = await _db.Users.SumAsync(u => (int?)u.Points) ?? 0;
        RecentOrders = await _db.Orders.OrderByDescending(o => o.CreatedAt).Take(5).ToListAsync();
    }
}
