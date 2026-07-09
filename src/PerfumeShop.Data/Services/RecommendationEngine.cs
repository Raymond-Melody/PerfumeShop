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
        // 按订单明细中的总销量排序（EF Core 可翻译的子查询）
        var popularIds = await _db.OrderDetails
            .GroupBy(od => od.ProductId)
            .Select(g => new { ProductId = g.Key, TotalSold = g.Sum(od => od.Quantity) })
            .OrderByDescending(x => x.TotalSold)
            .Take(count)
            .Select(x => x.ProductId)
            .ToListAsync(ct);

        if (popularIds.Count < count)
        {
            // 销量不足时用最新上架商品补齐
            var existingIds = new HashSet<int>(popularIds);
            var newProducts = await _db.Products
                .AsNoTracking()
                .Where(p => p.IsActive == true && !existingIds.Contains(p.ProductId))
                .OrderByDescending(p => p.CreatedAt)
                .Take(count - popularIds.Count)
                .Select(p => p.ProductId)
                .ToListAsync(ct);
            popularIds.AddRange(newProducts);
        }

        return popularIds;
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
