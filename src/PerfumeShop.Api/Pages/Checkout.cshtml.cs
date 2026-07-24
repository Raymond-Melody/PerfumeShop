using System.Security.Claims;
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
        // V21 修复：从 Cookie Authentication 的 Claims 读取（与 Login/Cart 一致）
        // 原实现错误地读不存在的 "UserId" cookie，导致已登录用户被踢回登录页
        var uidStr = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        return int.TryParse(uidStr, out var id) ? id : 0;
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
