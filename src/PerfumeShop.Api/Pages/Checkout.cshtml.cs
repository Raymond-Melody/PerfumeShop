using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

public class CheckoutModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public CheckoutModel(PerfumeShopContext db) => _db = db;

    public List<CheckoutItem> Items { get; set; } = new();
    public decimal SubTotal { get; set; }
    public decimal ShippingFee { get; set; } = 0;
    public decimal Total => SubTotal + ShippingFee;

    public async Task<IActionResult> OnGetAsync()
    {
        var userId = GetUserId();
        if (userId == 0) return RedirectToPage("/Login", new { returnUrl = "/checkout" });

        var cartItems = await _db.Carts
            .Where(c => c.UserId == userId)
            .ToListAsync();

        if (!cartItems.Any()) return RedirectToPage("/Cart");

        var productIds = cartItems.Select(c => c.ProductId).Distinct().ToList();
        var products = await _db.Products
            .Where(p => productIds.Contains(p.ProductId))
            .ToDictionaryAsync(p => p.ProductId);

        foreach (var ci in cartItems)
        {
            if (products.TryGetValue(ci.ProductId, out var prod))
            {
                Items.Add(new CheckoutItem
                {
                    CartId = ci.CartId,
                    ProductId = prod.ProductId,
                    ProductName = prod.ProductName,
                    ImageUrl = prod.ImageUrl,
                    Quantity = ci.Quantity ?? 1,
                    UnitPrice = prod.BasePrice
                });
            }
        }

        SubTotal = Items.Sum(i => i.UnitPrice * i.Quantity);
        ShippingFee = SubTotal >= 299 ? 0 : 15;
        return Page();
    }

    private int GetUserId()
    {
        if (Request.Cookies.TryGetValue("UserId", out var val) && int.TryParse(val, out var id))
            return id;
        return 0;
    }
}

public class CheckoutItem
{
    public int CartId { get; set; }
    public int ProductId { get; set; }
    public string ProductName { get; set; } = "";
    public string? ImageUrl { get; set; }
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
