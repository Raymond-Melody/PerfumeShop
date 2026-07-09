using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Api.Pages;

public class CartModel : PageModel
{
    private readonly PerfumeShopContext _db;
    public CartModel(PerfumeShopContext db) => _db = db;

    public List<CartItemView> Items { get; set; } = new();
    public decimal TotalAmount { get; set; }

    public async Task OnGetAsync()
    {
        // 简单获取所有购物车项目 (实际应基于 Session 或 UserId)
        var carts = await _db.Carts.Take(50).ToListAsync();
        var productIds = carts.Select(c => c.ProductId).Distinct().ToList();
        var products = productIds.Any()
            ? await _db.Products.Where(p => productIds.Contains(p.ProductId)).ToDictionaryAsync(p => p.ProductId)
            : new Dictionary<int, Product>();

        foreach (var c in carts)
        {
            var prod = products.GetValueOrDefault(c.ProductId);
            Items.Add(new CartItemView
            {
                CartId = c.CartId,
                Product = prod,
                Quantity = c.Quantity ?? 1,
                UnitPrice = c.UnitPrice
            });
        }
        TotalAmount = Items.Sum(i => i.UnitPrice * i.Quantity);
    }
}

public class CartItemView
{
    public int CartId { get; set; }
    public Product? Product { get; set; }
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
