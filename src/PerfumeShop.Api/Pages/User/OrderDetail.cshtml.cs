using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

[Authorize]
public class OrderDetailModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public OrderDetailModel(PerfumeShopContext db) => _db = db;

    public Order? Order { get; set; }
    public List<OrderDetail> Details { get; set; } = new();
    public List<NoteSelectionView> NoteSelections { get; set; } = new();
    public List<OrderIngredient> Ingredients { get; set; } = new();
    public bool HasCustomOrKol { get; set; }

    public async Task<IActionResult> OnGetAsync(int id)
    {
        var userIdClaim = User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
        if (string.IsNullOrEmpty(userIdClaim) || !int.TryParse(userIdClaim, out var userId))
            return RedirectToPage("/Login");

        Order = await _db.Orders
            .FirstOrDefaultAsync(o => o.OrderId == id && o.UserId == userId);

        if (Order == null)
            return RedirectToPage("/User/Orders");

        // 获取订单商品明细
        Details = await _db.OrderDetails
            .Where(d => d.OrderId == id)
            .ToListAsync();

        // 获取香调配比 (JOIN FragranceNotes 获取名称)
        var detailIds = Details.Select(d => d.DetailId).ToList();
        NoteSelections = await _db.OrderDetailNoteSelections
            .Where(s => detailIds.Contains(s.DetailId))
            .Join(_db.FragranceNotes,
                s => s.NoteId,
                n => n.NoteId,
                (s, n) => new NoteSelectionView
                {
                    DetailId = s.DetailId,
                    NoteType = s.NoteType,
                    NoteName = n.NoteName,
                    Percentage = s.Percentage
                })
            .ToListAsync();

        // 获取成分信息
        Ingredients = await _db.OrderIngredients
            .Where(i => i.OrderId == id)
            .ToListAsync();

        // 检查是否有定制或 KOL 产品
        HasCustomOrKol = await _db.OrderDetails
            .Join(_db.Products, od => od.ProductId, p => p.ProductId, (od, p) => new { od.OrderId, p.ProductType })
            .AnyAsync(x => x.OrderId == id &&
                (x.ProductType == "Custom" || x.ProductType == "KOL"));

        return Page();
    }
}

public class NoteSelectionView
{
    public int DetailId { get; set; }
    public string? NoteType { get; set; }
    public string? NoteName { get; set; }
    public int Percentage { get; set; }
}
