using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.OutputCaching;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

[OutputCache(PolicyName = "ProductDetail")]
public class ProductModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public ProductModel(PerfumeShopContext db) => _db = db;

    public Product? Product { get; set; }
    public FlashSale? ActiveFlashSale { get; set; }

    public async Task<IActionResult> OnGetAsync(int? id)
    {
        if (id == null || id <= 0) return RedirectToPage("/Products");

        Product = await _db.Products
            .AsNoTracking()
            .FirstOrDefaultAsync(p => p.ProductId == id && p.IsActive == true);

        if (Product == null) return Page(); // 模板会显示"商品不存在"

        // 检查是否有活跃秒杀
        var now = DateTime.Now;
        ActiveFlashSale = await _db.FlashSales
            .AsNoTracking()
            .FirstOrDefaultAsync(f => f.ProductId == id && f.IsActive
                && f.StartTime <= now && f.EndTime >= now);

        return Page();
    }
}
