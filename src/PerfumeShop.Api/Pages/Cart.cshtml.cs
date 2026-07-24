using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
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

    /// <summary>
    /// V21 修复：处理 ?add=/?remove= 并按当前登录用户过滤购物车
    /// 原实现忽略 add 参数且读取所有人购物车（越权），导致加购无效
    /// </summary>
    public async Task<IActionResult> OnGetAsync(int? add = null, int? remove = null)
    {
        // 未登录 → 引导登录（购物车绑定用户）
        var uidStr = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        if (!int.TryParse(uidStr, out var userId) || userId <= 0)
        {
            if (add.HasValue || remove.HasValue)
                return Redirect("/login?returnUrl=/cart");
        }

        // 加入购物车（合并同商品数量）— Cart 为 keyless 实体，必须用原生 SQL 写入
        if (add.HasValue && add.Value > 0 && userId > 0)
        {
            var product = await _db.Products.AsNoTracking().FirstOrDefaultAsync(p => p.ProductId == add.Value && p.IsActive == true);
            if (product != null)
            {
                var existingQty = await _db.Carts.AsNoTracking()
                    .Where(c => c.UserId == userId && c.ProductId == add.Value)
                    .Select(c => (int?)c.Quantity).FirstOrDefaultAsync();
                if (existingQty.HasValue)
                {
                    await _db.Database.ExecuteSqlInterpolatedAsync(
                        $"UPDATE Cart SET Quantity = COALESCE(Quantity,0) + 1 WHERE UserID = {userId} AND ProductID = {add.Value}");
                }
                else
                {
                    var price = product.BasePrice;
                    var now = DateTime.Now;
                    await _db.Database.ExecuteSqlInterpolatedAsync(
                        $"INSERT INTO Cart (UserID, ProductID, Quantity, UnitPrice, CreatedAt) VALUES ({userId}, {add.Value}, 1, {price}, {now})");
                }
            }
            return Redirect("/cart"); // PRG：避免刷新重复加购
        }

        // 移除购物车项（按 CartId）
        if (remove.HasValue && remove.Value > 0 && userId > 0)
        {
            await _db.Database.ExecuteSqlInterpolatedAsync(
                $"DELETE FROM Cart WHERE CartID = {remove.Value} AND UserID = {userId}");
            return Redirect("/cart");
        }

        // 展示：仅当前用户的购物车
        if (userId > 0)
        {
            var carts = await _db.Carts.Where(c => c.UserId == userId).ToListAsync();
            var productIds = carts.Select(c => c.ProductId).Distinct().ToList();
            var products = productIds.Any()
                ? await _db.Products.Where(p => productIds.Contains(p.ProductId)).ToDictionaryAsync(p => p.ProductId)
                : new Dictionary<int, Product>();

            foreach (var c in carts)
            {
                Items.Add(new CartItemView
                {
                    CartId = c.CartId,
                    Product = products.GetValueOrDefault(c.ProductId),
                    Quantity = c.Quantity ?? 1,
                    UnitPrice = c.UnitPrice
                });
            }
            TotalAmount = Items.Sum(i => (i.Product?.BasePrice ?? i.UnitPrice) * i.Quantity);
        }

        return Page();
    }
}

public class CartItemView
{
    public int CartId { get; set; }
    public Product? Product { get; set; }
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
}
