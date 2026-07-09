namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 推荐引擎接口 — 个性化商品推荐
/// </summary>
public interface IRecommendationEngine
{
    /// <summary>基于用户偏好的推荐</summary>
    Task<IEnumerable<int>> GetPersonalizedRecommendationsAsync(int userId, int count = 6, CancellationToken ct = default);

    /// <summary>热门商品推荐</summary>
    Task<IEnumerable<int>> GetPopularProductsAsync(int count = 10, CancellationToken ct = default);

    /// <summary>基于商品的关联推荐 (同类型/同分类)</summary>
    Task<IEnumerable<int>> GetRelatedProductsAsync(int productId, int count = 4, CancellationToken ct = default);

    /// <summary>新品推荐</summary>
    Task<IEnumerable<int>> GetNewArrivalsAsync(int count = 6, CancellationToken ct = default);
}
