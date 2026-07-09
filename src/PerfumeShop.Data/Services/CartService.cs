using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

public class CartService : ICartService
{
    private readonly PerfumeShopContext _db;
    public CartService(PerfumeShopContext db) => _db = db;

    public async Task<CartSummary> GetCartAsync(int userId)
    {
        // Cart is keyless — use LINQ queries only
        var cartItems = await (from c in _db.Carts
                               join p in _db.Products on c.ProductId equals p.ProductId
                               where c.UserId == userId
                               select new CartItem
                               {
                                   ProductId = c.ProductId,
                                   ProductName = p.ProductName ?? "",
                                   ImageUrl = p.ImageUrl,
                                   UnitPrice = c.UnitPrice,
                                   Quantity = c.Quantity ?? 1,
                                   ProductType = p.ProductType,
                                   Size = c.CustomLabel
                               }).ToListAsync();

        return new CartSummary
        {
            UserId = userId,
            Items = cartItems
        };
    }

    public async Task<int> AddItemAsync(int userId, int productId, int quantity, string? size = null)
    {
        if (quantity <= 0) quantity = 1;

        // 检查购物车中是否已有该商品
        var existing = await _db.Carts
            .FirstOrDefaultAsync(c => c.UserId == userId && c.ProductId == productId);

        if (existing != null)
        {
            existing.Quantity = (existing.Quantity ?? 0) + quantity;
            await _db.SaveChangesAsync();
            return existing.Quantity ?? 0;
        }

        // 获取商品价格
        var product = await _db.Products.FirstOrDefaultAsync(p => p.ProductId == productId);
        if (product == null) return 0;

        var cart = new Cart
        {
            UserId = userId,
            ProductId = productId,
            Quantity = quantity,
            UnitPrice = product.BasePrice,
            CustomLabel = size,
            CreatedAt = DateTime.Now
        };
        _db.Carts.Add(cart);
        await _db.SaveChangesAsync();
        return quantity;
    }

    public async Task<bool> UpdateQuantityAsync(int userId, int productId, int quantity)
    {
        if (quantity <= 0) return await RemoveItemAsync(userId, productId);

        var item = await _db.Carts
            .FirstOrDefaultAsync(c => c.UserId == userId && c.ProductId == productId);
        if (item == null) return false;

        item.Quantity = quantity;
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> RemoveItemAsync(int userId, int productId)
    {
        var item = await _db.Carts
            .FirstOrDefaultAsync(c => c.UserId == userId && c.ProductId == productId);
        if (item == null) return false;

        _db.Carts.Remove(item);
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<bool> ClearCartAsync(int userId)
    {
        var items = await _db.Carts.Where(c => c.UserId == userId).ToListAsync();
        _db.Carts.RemoveRange(items);
        await _db.SaveChangesAsync();
        return true;
    }

    public async Task<int> GetCartCountAsync(int userId)
    {
        return await _db.Carts
            .Where(c => c.UserId == userId)
            .SumAsync(c => c.Quantity ?? 0);
    }
}
