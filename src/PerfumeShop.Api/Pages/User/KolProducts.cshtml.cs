using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages.User;

public class KolProductsModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public KolProductsModel(PerfumeShopContext db) => _db = db;

    public List<KolProductItem> Products { get; set; } = new();
    public string? FilterCategory { get; set; }
    public string? FilterKeyword { get; set; }

    public class KolProductItem
    {
        public int ProductId { get; set; }
        public string ProductName { get; set; } = "";
        public decimal BasePrice { get; set; }
        public string? ImageUrl { get; set; }
        public string? Description { get; set; }
        public string? Category { get; set; }
        public DateTime? CreatedAt { get; set; }
        public int? Kolid { get; set; }
    }

    public async Task OnGetAsync(string? category = null, string? keyword = null)
    {
        FilterCategory = category;
        FilterKeyword = keyword;

        var query = _db.Products.Where(p => p.Kolid != null && p.ProductType == "KOL");
        if (!string.IsNullOrEmpty(category))
            query = query.Where(p => p.Category == category);
        if (!string.IsNullOrEmpty(keyword))
            query = query.Where(p => p.ProductName.Contains(keyword) || (p.Description != null && p.Description.Contains(keyword)));

        Products = await query.OrderByDescending(p => p.CreatedAt).Take(50)
            .Select(p => new KolProductItem
            {
                ProductId = p.ProductId,
                ProductName = p.ProductName,
                BasePrice = p.BasePrice,
                ImageUrl = p.ImageUrl,
                Description = p.Description,
                Category = p.Category,
                CreatedAt = p.CreatedAt,
                Kolid = p.Kolid,
            }).ToListAsync();
    }
}
