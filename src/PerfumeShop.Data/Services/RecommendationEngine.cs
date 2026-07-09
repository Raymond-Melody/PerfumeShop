using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 推荐引擎 — 个性化/热门/关联/新品推荐
/// </summary>
public class RecommendationEngine : IRecommendationEngine
{
    private readonly PerfumeShopContext _db;

    public RecommendationEngine(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    public async Task<IEnumerable<int>> GetPersonalizedRecommendationsAsync(int userId, int count = 6, CancellationToken ct = default)
    {
        // 获取用户偏好
        var user = await _db.Users
            .AsNoTracking()
            .Select(u => new { u.UserId, u.FavoriteCategory, u.PreferredNote })
            .FirstOrDefaultAsync(u => u.UserId == userId, ct);

        if (user == null)
            return await GetPopularProductsAsync(count, ct);

        var query = _db.Products.AsNoTracking().Where(p => p.IsActive == true);

        // 按偏好分类过滤
        if (!string.IsNullOrEmpty(user.FavoriteCategory))
            query = query.Where(p => p.Category == user.FavoriteCategory);

        // 按偏好香型取Top
        return await query
            .OrderByDescending(p => p.CreatedAt)
            .Take(count)
            .Select(p => p.ProductId)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<int>> GetPopularProductsAsync(int count = 10, CancellationToken ct = default)
    {
        // 按购买量排序 (通过关联订单明细统计)
        return await (from p in _db.Products
                      join od in _db.OrderDetails on p.ProductId equals od.ProductId into odJoin
                      from od in odJoin.DefaultIfEmpty()
                      where p.IsActive == true
                      group od by new { p.ProductId, p.CreatedAt } into g
                      orderby g.Sum(x => x != null ? x.Quantity : 0) descending, g.Key.CreatedAt descending
                      select g.Key.ProductId)
                      .Take(count)
                      .ToListAsync(ct);
    }

    public async Task<IEnumerable<int>> GetRelatedProductsAsync(int productId, int count = 4, CancellationToken ct = default)
    {
        var product = await _db.Products
            .AsNoTracking()
            .Select(p => new { p.ProductId, p.Category, p.ProductType })
            .FirstOrDefaultAsync(p => p.ProductId == productId, ct);

        if (product == null) return Enumerable.Empty<int>();

        return await _db.Products
            .AsNoTracking()
            .Where(p => p.ProductId != productId && p.IsActive == true &&
                        (p.Category == product.Category || p.ProductType == product.ProductType))
            .OrderBy(p => Guid.NewGuid()) // 随机
            .Take(count)
            .Select(p => p.ProductId)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<int>> GetNewArrivalsAsync(int count = 6, CancellationToken ct = default)
    {
        return await _db.Products
            .AsNoTracking()
            .Where(p => p.IsActive == true)
            .OrderByDescending(p => p.CreatedAt)
            .Take(count)
            .Select(p => p.ProductId)
            .ToListAsync(ct);
    }
}
