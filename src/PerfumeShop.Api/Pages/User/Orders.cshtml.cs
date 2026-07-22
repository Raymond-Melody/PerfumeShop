using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

[Authorize]
public class OrdersModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public OrdersModel(PerfumeShopContext db) => _db = db;

    public List<Order> Orders { get; set; } = new();
    public int CurrentPage { get; set; } = 1;
    public int TotalPages { get; set; } = 1;
    public string? StatusFilter { get; set; }
    public int UserId { get; set; }

    public async Task<IActionResult> OnGetAsync(int page = 1, string? status = null)
    {
        var userIdClaim = User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
        if (string.IsNullOrEmpty(userIdClaim) || !int.TryParse(userIdClaim, out var userId))
            return RedirectToPage("/Login");

        UserId = userId;
        StatusFilter = status;
        CurrentPage = Math.Max(1, page);
        const int pageSize = 10;

        var query = _db.Orders
            .Where(o => o.UserId == userId && o.Status != "Deleted");

        if (!string.IsNullOrEmpty(status))
            query = query.Where(o => o.Status == status);

        var total = await query.CountAsync();
        TotalPages = (int)Math.Ceiling((double)total / pageSize);
        if (TotalPages < 1) TotalPages = 1;
        if (CurrentPage > TotalPages) CurrentPage = TotalPages;

        Orders = await query
            .OrderByDescending(o => o.CreatedAt)
            .Skip((CurrentPage - 1) * pageSize)
            .Take(pageSize)
            .ToListAsync();

        return Page();
    }
}
