namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 推荐引擎接口 — 个性化/热门/关联/趋势/新品推荐
/// 对应 ASP: recommendation_engine.asp + ai-service/recommendation_engine.py
/// </summary>
public interface IRecommendationEngine
{
    // ========== 个性化推荐 (V18: RE_GetPersonalizedProducts + Python get_personalized) ==========

    /// <summary>协同过滤 + 内容推荐 + 趋势加权 — 对标 Python RecommendationEngine.get_personalized</summary>
    Task<IEnumerable<RecommendationItem>> GetPersonalizedAsync(int userId, int count = 6, CancellationToken ct = default);

    /// <summary>基于用户偏好的推荐（兼容旧接口）</summary>
    Task<IEnumerable<int>> GetPersonalizedRecommendationsAsync(int userId, int count = 6, CancellationToken ct = default);

    // ========== 热门商品 (V18: RE_GetPopularProducts) ==========

    /// <summary>热门商品推荐 — 按销量排序</summary>
    Task<IEnumerable<int>> GetPopularProductsAsync(int count = 10, CancellationToken ct = default);

    // ========== 相关商品 (V18: RE_GetRelatedProducts / RE_GetSimilarFragrances) ==========

    /// <summary>基于商品的关联推荐 (同类型/同分类) + 余弦相似度</summary>
    Task<IEnumerable<RecommendationItem>> GetRelatedAsync(int productId, int count = 4, CancellationToken ct = default);

    /// <summary>基于商品的关联推荐 (兼容旧接口)</summary>
    Task<IEnumerable<int>> GetRelatedProductsAsync(int productId, int count = 4, CancellationToken ct = default);

    // ========== 趋势分析 (V18: RE_GetTrendingNow) ==========

    /// <summary>趋势分析 — 近期销量增速排名</summary>
    Task<IEnumerable<RecommendationItem>> GetTrendingAsync(int count = 8, CancellationToken ct = default);

    // ========== 新品推荐 (V18: RE_GetNewProducts) ==========

    /// <summary>新品推荐</summary>
    Task<IEnumerable<int>> GetNewArrivalsAsync(int count = 6, CancellationToken ct = default);
}

/// <summary>推荐项 — 包含评分和推荐理由</summary>
public class RecommendationItem
{
    public int ProductId { get; set; }
    public double Score { get; set; }
    public string Reason { get; set; } = "";
}
