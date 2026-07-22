using Microsoft.EntityFrameworkCore;
using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 推荐引擎 — 协同过滤 + 内容推荐 + 趋势分析
/// 对应 ASP: recommendation_engine.asp + ai-service/recommendation_engine.py
/// V18 函数映射:
///   RE_GetPersonalizedProducts → GetPersonalizedAsync
///   RE_GetPopularProducts      → GetPopularProductsAsync
///   RE_GetRelatedProducts      → GetRelatedAsync / GetRelatedProductsAsync
///   RE_GetTrendingNow          → GetTrendingAsync
///   RE_GetNewProducts          → GetNewArrivalsAsync
///   RE_GetSimilarFragrances    → GetRelatedAsync
///   Python get_personalized    → GetPersonalizedAsync
///   Python get_similar_products→ GetRelatedAsync
///   Python get_trending        → GetTrendingAsync
///   Python _find_similar_users → FindSimilarUsersAsync (private)
///   Python _cosine_similarity  → CosineSimilarity (private)
///   Python _compute_user_profile→ ComputeUserProfile (private)
/// </summary>
public class RecommendationEngine : IRecommendationEngine
{
    private readonly PerfumeShopContext _db;

    // V18: 协同过滤权重 (对应 Python: sim_score * 1.0)
    private const double CollaborativeWeight = 1.0;
    // V18: 内容推荐权重 (对应 Python: content_score * 0.5)
    private const double ContentBasedWeight = 0.5;
    // V18: 趋势加权 (对应 Python: trend_score * 0.3)
    private const double TrendingWeight = 0.3;
    // V18: 相似用户数量上限 (对应 Python: top_k=20)
    private const int SimilarUsersTopK = 20;
    // V18: 趋势分析时间窗口（天）
    private const int TrendingWindowDays = 30;

    public RecommendationEngine(PerfumeShopContext db)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
    }

    // ==================== V18: RE_GetPersonalizedProducts + Python get_personalized ====================

    /// <summary>协同过滤 + 内容推荐 + 趋势加权 — 对标 Python RecommendationEngine.get_personalized</summary>
    public async Task<IEnumerable<RecommendationItem>> GetPersonalizedAsync(int userId, int count = 6, CancellationToken ct = default)
    {
        if (count <= 0) count = 6;

        // 获取目标用户购买历史
        var userProductIds = await GetUserPurchaseHistoryAsync(userId, ct);
        var favoriteProductIds = await GetUserFavoriteIdsAsync(userId, ct);
        var excludeSet = new HashSet<int>(userProductIds);
        excludeSet.UnionWith(favoriteProductIds);

        var scores = new Dictionary<int, double>();

        // ====== 策略1: 协同过滤 (V18: Collaborative — Find similar users) ======
        var similarUsers = await FindSimilarUsersAsync(userId, SimilarUsersTopK, ct);
        foreach (var (simUserId, simScore) in similarUsers)
        {
            var simUserProducts = await GetUserPurchaseHistoryAsync(simUserId, ct);
            foreach (var pid in simUserProducts)
            {
                if (!excludeSet.Contains(pid))
                {
                    scores.TryGetValue(pid, out var current);
                    scores[pid] = current + simScore * CollaborativeWeight;
                }
            }
        }

        // ====== 策略2: 内容推荐 (V18: Content-based — Cosine similarity) ======
        if (userProductIds.Count > 0)
        {
            var userProfile = await ComputeUserProfileAsync(userProductIds, ct);
            if (userProfile.Count > 0)
            {
                // 获取所有活跃商品的特征向量
                var allProducts = await _db.Products
                    .AsNoTracking()
                    .Where(p => p.IsActive == true && !excludeSet.Contains(p.ProductId))
                    .Select(p => new { p.ProductId, p.Category, p.ProductType })
                    .ToListAsync(ct);

                foreach (var p in allProducts)
                {
                    var productFeatures = BuildProductFeatureVector(p.Category, p.ProductType);
                    if (productFeatures.Count > 0)
                    {
                        var contentScore = CosineSimilarity(userProfile, productFeatures);
                        if (contentScore > 0)
                        {
                            scores.TryGetValue(p.ProductId, out var current);
                            scores[p.ProductId] = current + contentScore * ContentBasedWeight;
                        }
                    }
                }
            }
        }

        // ====== 策略3: 趋势加权 (V18: Boost trending items) ======
        var trendingScores = await ComputeTrendingScoresAsync(ct);
        foreach (var (pid, trendScore) in trendingScores)
        {
            if (!excludeSet.Contains(pid))
            {
                scores.TryGetValue(pid, out var current);
                scores[pid] = current + trendScore * TrendingWeight;
            }
        }

        // ====== 排序取 Top N ======
        var sortedItems = scores
            .OrderByDescending(kv => kv.Value)
            .Take(count)
            .Select(kv => new RecommendationItem
            {
                ProductId = kv.Key,
                Score = Math.Round(kv.Value, 4),
                Reason = GetRecommendationReason(kv.Key, trendingScores, userProductIds)
            })
            .ToList();

        // ====== 回退策略: 如果协同过滤+内容推荐不足，补充热门商品 ======
        if (sortedItems.Count < count)
        {
            var existingIds = new HashSet<int>(sortedItems.Select(x => x.ProductId));
            existingIds.UnionWith(excludeSet);

            var popular = await GetPopularFallbackAsync(count - sortedItems.Count, existingIds, ct);
            foreach (var item in popular)
            {
                sortedItems.Add(new RecommendationItem
                {
                    ProductId = item,
                    Score = 0.1,
                    Reason = "热门商品"
                });
            }
        }

        return sortedItems;
    }

    /// <summary>兼容旧接口 — 返回 ProductId 列表</summary>
    public async Task<IEnumerable<int>> GetPersonalizedRecommendationsAsync(int userId, int count = 6, CancellationToken ct = default)
    {
        var items = await GetPersonalizedAsync(userId, count, ct);
        return items.Select(i => i.ProductId);
    }

    // ==================== V18: RE_GetPopularProducts ====================

    /// <summary>热门商品推荐 — 按销量排序</summary>
    public async Task<IEnumerable<int>> GetPopularProductsAsync(int count = 10, CancellationToken ct = default)
    {
        var popularIds = await _db.OrderDetails
            .GroupBy(od => od.ProductId)
            .Select(g => new { ProductId = g.Key, TotalSold = g.Sum(od => od.Quantity) })
            .OrderByDescending(x => x.TotalSold)
            .Take(count)
            .Select(x => x.ProductId)
            .ToListAsync(ct);

        // V18: 补充 — 销量不足时用活跃商品补充（含无销量新品）
        if (popularIds.Count < count)
        {
            var existingIds = new HashSet<int>(popularIds);
            var fillProducts = await _db.Products
                .AsNoTracking()
                .Where(p => p.IsActive == true && !existingIds.Contains(p.ProductId))
                .OrderByDescending(p => p.CreatedAt)
                .Take(count - popularIds.Count)
                .Select(p => p.ProductId)
                .ToListAsync(ct);
            popularIds.AddRange(fillProducts);
        }

        return popularIds;
    }

    // ==================== V18: RE_GetRelatedProducts / RE_GetSimilarFragrances ====================

    /// <summary>基于商品的关联推荐 — 同类型/同分类 + 余弦相似度</summary>
    public async Task<IEnumerable<RecommendationItem>> GetRelatedAsync(int productId, int count = 4, CancellationToken ct = default)
    {
        var product = await _db.Products
            .AsNoTracking()
            .Select(p => new { p.ProductId, p.Category, p.ProductType })
            .FirstOrDefaultAsync(p => p.ProductId == productId, ct);

        if (product == null) return Enumerable.Empty<RecommendationItem>();

        var targetFeatures = BuildProductFeatureVector(product.Category, product.ProductType);

        // V18: 同类型产品推荐（排除当前产品）
        var candidates = await _db.Products
            .AsNoTracking()
            .Where(p => p.ProductId != productId && p.IsActive == true &&
                        (p.Category == product.Category || p.ProductType == product.ProductType))
            .Select(p => new { p.ProductId, p.Category, p.ProductType })
            .ToListAsync(ct);

        // V18: 用余弦相似度打分
        var scored = candidates
            .Select(p =>
            {
                var features = BuildProductFeatureVector(p.Category, p.ProductType);
                var sim = CosineSimilarity(targetFeatures, features);
                return new RecommendationItem
                {
                    ProductId = p.ProductId,
                    Score = Math.Round(sim, 4),
                    Reason = "相似产品"
                };
            })
            .OrderByDescending(x => x.Score)
            .Take(count)
            .ToList();

        // 如果同类型不足，补充随机推荐
        if (scored.Count < count)
        {
            var existingIds = new HashSet<int>(scored.Select(x => x.ProductId)) { productId };
            var fallback = await _db.Products
                .AsNoTracking()
                .Where(p => p.IsActive == true && !existingIds.Contains(p.ProductId))
                .OrderByDescending(p => p.CreatedAt)
                .Take(count - scored.Count)
                .Select(p => p.ProductId)
                .ToListAsync(ct);

            foreach (var pid in fallback)
            {
                scored.Add(new RecommendationItem { ProductId = pid, Score = 0.05, Reason = "推荐" });
            }
        }

        return scored;
    }

    /// <summary>兼容旧接口 — 返回 ProductId 列表</summary>
    public async Task<IEnumerable<int>> GetRelatedProductsAsync(int productId, int count = 4, CancellationToken ct = default)
    {
        var items = await GetRelatedAsync(productId, count, ct);
        return items.Select(i => i.ProductId);
    }

    // ==================== V18: RE_GetTrendingNow ====================

    /// <summary>趋势分析 — 近期销量增速排名</summary>
    public async Task<IEnumerable<RecommendationItem>> GetTrendingAsync(int count = 8, CancellationToken ct = default)
    {
        var trendingScores = await ComputeTrendingScoresAsync(ct);

        return trendingScores
            .OrderByDescending(kv => kv.Value)
            .Take(count)
            .Select(kv => new RecommendationItem
            {
                ProductId = kv.Key,
                Score = Math.Round(kv.Value, 4),
                Reason = "热门趋势"
            })
            .ToList();
    }

    // ==================== V18: RE_GetNewProducts ====================

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

    // ==================== 协同过滤核心算法 ====================

    /// <summary>
    /// V18: 查找相似用户 — Jaccard 相似度
    /// 对标 Python _find_similar_users: Jaccard = |intersection| / |union|
    /// </summary>
    private async Task<List<(int UserId, double Similarity)>> FindSimilarUsersAsync(int userId, int topK, CancellationToken ct)
    {
        // 获取目标用户购买历史
        var targetProducts = new HashSet<int>(await GetUserPurchaseHistoryAsync(userId, ct));
        if (targetProducts.Count == 0) return new List<(int, double)>();

        // 获取所有有购买记录的用户 — 客户端分组（SQLite 不支持 GroupBy + ToList 子查询）
        var rawData = await _db.OrderDetails
            .AsNoTracking()
            .Join(_db.Orders.AsNoTracking(), od => od.OrderId, o => o.OrderId, (od, o) => new { o.UserId, od.ProductId })
            .Where(x => x.UserId != userId)
            .ToListAsync(ct);

        var allUserProducts = rawData
            .GroupBy(x => x.UserId)
            .Select(g => new { UserId = g.Key, ProductIds = g.Select(x => x.ProductId).Distinct().ToList() })
            .ToList();

        var similarities = new List<(int UserId, double Similarity)>();

        foreach (var other in allUserProducts)
        {
            var otherProducts = new HashSet<int>(other.ProductIds);
            if (otherProducts.Count == 0) continue;

            // V18: Jaccard similarity = |intersection| / |union|
            var intersection = targetProducts.Intersect(otherProducts).Count();
            if (intersection > 0)
            {
                var union = targetProducts.Union(otherProducts).Count();
                var jaccard = union > 0 ? (double)intersection / union : 0;
                similarities.Add((other.UserId, jaccard));
            }
        }

        return similarities
            .OrderByDescending(x => x.Similarity)
            .Take(topK)
            .ToList();
    }

    /// <summary>
    /// V18: 计算用户偏好画像 — 对标 Python _compute_user_profile
    /// 从用户购买商品的特征向量取平均值
    /// </summary>
    private async Task<Dictionary<string, double>> ComputeUserProfileAsync(List<int> productIds, CancellationToken ct)
    {
        var products = await _db.Products
            .AsNoTracking()
            .Where(p => productIds.Contains(p.ProductId))
            .Select(p => new { p.Category, p.ProductType })
            .ToListAsync(ct);

        var profile = new Dictionary<string, double>();
        int count = 0;

        foreach (var p in products)
        {
            var features = BuildProductFeatureVector(p.Category, p.ProductType);
            foreach (var (key, value) in features)
            {
                profile.TryGetValue(key, out var current);
                profile[key] = current + value;
            }
            count++;
        }

        // V18: 取平均
        if (count > 0)
        {
            foreach (var key in profile.Keys.ToList())
            {
                profile[key] /= count;
            }
        }

        return profile;
    }

    /// <summary>
    /// V18: 余弦相似度 — 对标 Python _cosine_similarity
    /// cosine_sim = dot(A,B) / (||A|| * ||B||)
    /// </summary>
    private static double CosineSimilarity(Dictionary<string, double> vec1, Dictionary<string, double> vec2)
    {
        var allKeys = new HashSet<string>(vec1.Keys);
        allKeys.UnionWith(vec2.Keys);

        double dotProduct = 0;
        double norm1 = 0;
        double norm2 = 0;

        foreach (var key in allKeys)
        {
            vec1.TryGetValue(key, out var v1);
            vec2.TryGetValue(key, out var v2);
            dotProduct += v1 * v2;
            norm1 += v1 * v1;
            norm2 += v2 * v2;
        }

        norm1 = Math.Sqrt(norm1);
        norm2 = Math.Sqrt(norm2);

        if (norm1 == 0 || norm2 == 0) return 0;
        return dotProduct / (norm1 * norm2);
    }

    /// <summary>构建产品特征向量（基于 Category + ProductType）</summary>
    private static Dictionary<string, double> BuildProductFeatureVector(string? category, string? productType)
    {
        var features = new Dictionary<string, double>();

        if (!string.IsNullOrEmpty(category))
            features[$"cat:{category}"] = 1.0;

        if (!string.IsNullOrEmpty(productType))
            features[$"type:{productType}"] = 1.0;

        return features;
    }

    /// <summary>
    /// V18: 计算趋势评分 — 近期销量 / 历史平均销量
    /// 对标 Python _trending_scores
    /// </summary>
    private async Task<Dictionary<int, double>> ComputeTrendingScoresAsync(CancellationToken ct)
    {
        var recentCutoff = DateTime.Now.AddDays(-TrendingWindowDays);

        // 近期销量（30天内）
        var recentSales = await _db.OrderDetails
            .AsNoTracking()
            .Join(_db.Orders.AsNoTracking(), od => od.OrderId, o => o.OrderId, (od, o) => new { od.ProductId, o.Status, o.CreatedAt })
            .Where(x => x.Status == "Paid" && x.CreatedAt >= recentCutoff)
            .GroupBy(x => x.ProductId)
            .Select(g => new { ProductId = g.Key, RecentQty = g.Sum(x => 1) })
            .ToListAsync(ct);

        // 历史总销量
        var totalSales = await _db.OrderDetails
            .AsNoTracking()
            .Join(_db.Orders.AsNoTracking(), od => od.OrderId, o => o.OrderId, (od, o) => new { od.ProductId, o.Status })
            .Where(x => x.Status == "Paid")
            .GroupBy(x => x.ProductId)
            .Select(g => new { ProductId = g.Key, TotalQty = g.Count() })
            .ToListAsync(ct);

        var totalMap = totalSales.ToDictionary(x => x.ProductId, x => x.TotalQty);
        var result = new Dictionary<int, double>();

        foreach (var recent in recentSales)
        {
            totalMap.TryGetValue(recent.ProductId, out var total);
            // 趋势分 = 近期销量 / max(历史总销量, 1)
            var trendScore = total > 0 ? (double)recent.RecentQty / Math.Max(total, 1) : (double)recent.RecentQty;
            result[recent.ProductId] = Math.Round(trendScore, 4);
        }

        return result;
    }

    /// <summary>获取用户购买历史 — ProductId 列表</summary>
    private async Task<List<int>> GetUserPurchaseHistoryAsync(int userId, CancellationToken ct)
    {
        return await _db.OrderDetails
            .AsNoTracking()
            .Join(_db.Orders.AsNoTracking(), od => od.OrderId, o => o.OrderId, (od, o) => new { o.UserId, od.ProductId })
            .Where(x => x.UserId == userId)
            .Select(x => x.ProductId)
            .Distinct()
            .ToListAsync(ct);
    }

    /// <summary>获取用户收藏的 ProductId 列表</summary>
    private async Task<List<int>> GetUserFavoriteIdsAsync(int userId, CancellationToken ct)
    {
        return await _db.UserFavorites
            .AsNoTracking()
            .Where(f => f.UserId == userId)
            .Select(f => f.ProductId)
            .ToListAsync(ct);
    }

    /// <summary>热门商品回退</summary>
    private async Task<List<int>> GetPopularFallbackAsync(int count, HashSet<int> excludeIds, CancellationToken ct)
    {
        var ids = await _db.OrderDetails
            .AsNoTracking()
            .GroupBy(od => od.ProductId)
            .Select(g => new { ProductId = g.Key, TotalSold = g.Sum(od => od.Quantity) })
            .Where(x => !excludeIds.Contains(x.ProductId))
            .OrderByDescending(x => x.TotalSold)
            .Take(count)
            .Select(x => x.ProductId)
            .ToListAsync(ct);

        // V18: 无销量数据时回退到活跃商品
        if (ids.Count < count)
        {
            var existingIds = new HashSet<int>(ids);
            existingIds.UnionWith(excludeIds);
            var fillIds = await _db.Products
                .AsNoTracking()
                .Where(p => p.IsActive == true && !existingIds.Contains(p.ProductId))
                .OrderByDescending(p => p.CreatedAt)
                .Take(count - ids.Count)
                .Select(p => p.ProductId)
                .ToListAsync(ct);
            ids.AddRange(fillIds);
        }

        return ids;
    }

    /// <summary>生成推荐理由 — 对标 Python _get_recommendation_reason</summary>
    private static string GetRecommendationReason(int productId, Dictionary<int, double> trendingScores, List<int> userProducts)
    {
        var reasons = new List<string>();

        if (trendingScores.TryGetValue(productId, out var trendScore) && trendScore > 0.5)
            reasons.Add("热门趋势");

        if (userProducts.Count > 0)
            reasons.Add("其他用户也喜欢");

        reasons.Add("与你喜爱的产品相似");

        return reasons.Count > 0 ? string.Join("、", reasons) : "综合推荐";
    }
}
