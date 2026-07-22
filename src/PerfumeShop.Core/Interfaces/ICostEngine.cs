namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 成本引擎接口 — 三级成本自动传导：原材料 → 香调 → 产品 → 订单
/// 对应 ASP 中的 cost_engine.asp (45KB)
/// </summary>
public interface ICostEngine
{
    // ==================== 缓存预加载 ====================

    /// <summary>全局缓存预加载（原材料→香调→产品三层预热）— 对应 CE_PreloadAllCostData()</summary>
    Task PreloadAllCostDataAsync(CancellationToken ct = default);

    /// <summary>预加载跨品类加权批次成本 — 对应 CE_PreloadBatchCosts()</summary>
    Task PreloadBatchCostsAsync(CancellationToken ct = default);

    // ==================== Level 1-4: 成本计算 ====================

    /// <summary>计算单个原料成本 (加权平均/供应商报价回退) — 对应 CE_CalculateMaterialCost()</summary>
    Task<decimal> CalculateMaterialCostAsync(int materialId, CancellationToken ct = default);

    /// <summary>计算单个香调成本 (Accord配方或成分聚合) — 对应 CE_CalculateNoteCost()</summary>
    Task<decimal> CalculateNoteCostAsync(int noteId, CancellationToken ct = default);

    /// <summary>计算产品 BOM 成本 (香调配比 + 瓶身) — 对应 CE_CalculateProductBOMCost()</summary>
    Task<decimal> CalculateProductBomCostAsync(int productId, CancellationToken ct = default);

    /// <summary>计算产品完整单位成本 (BOM + 包装 + 人工) — 对应 CE_CalculateProductUnitCost()</summary>
    Task<decimal> CalculateProductUnitCostAsync(int productId, CancellationToken ct = default);

    // ==================== 缓存查询（O(1)） ====================

    /// <summary>从缓存获取原料成本 — 对应 CE_GetCachedMaterialCost()</summary>
    Task<decimal> GetCachedMaterialCostAsync(int materialId, CancellationToken ct = default);

    /// <summary>从缓存获取香调成本 — 对应 CE_GetCachedNoteCost()</summary>
    Task<decimal> GetCachedNoteCostAsync(int noteId, CancellationToken ct = default);

    /// <summary>从缓存获取产品BOM成本 — 对应 CE_GetCachedProductBOMCost()</summary>
    Task<decimal> GetCachedProductBomCostAsync(int productId, CancellationToken ct = default);

    /// <summary>从缓存获取产品单位总成本 — 对应 CE_GetCachedProductUnitCost()</summary>
    Task<decimal> GetCachedProductUnitCostAsync(int productId, CancellationToken ct = default);

    /// <summary>从缓存获取品牌定香采购成本 — 对应 CE_GetCachedFixedBrandCost()</summary>
    Task<decimal> GetFixedBrandCostAsync(int productId, CancellationToken ct = default);

    // ==================== 批次加权成本 ====================

    /// <summary>获取指定物料的加权批次成本 — 对应 CE_GetCachedBatchCost()</summary>
    Task<decimal> GetBatchWeightedCostAsync(string itemCode, CancellationToken ct = default);

    /// <summary>获取物料加权平均成本(查询5个品类) — 对应 CE_GetBatchWeightedCost()</summary>
    Task<decimal> CalculateBatchWeightedCostAsync(string itemCode, CancellationToken ct = default);

    // ==================== Level 5-8: 成本更新 ====================

    /// <summary>更新指定产品的成本到数据库 — 对应 CE_UpdateProductCost()</summary>
    Task UpdateProductCostAsync(int productId, CancellationToken ct = default);

    /// <summary>批量更新所有启用产品的成本 — 对应 CE_UpdateAllProductCosts()</summary>
    Task<int> UpdateAllProductCostsAsync(CancellationToken ct = default);

    /// <summary>更新订单的成本和利润 — 对应 CE_UpdateOrderCosts()</summary>
    Task UpdateOrderCostsAsync(int orderId, CancellationToken ct = default);

    /// <summary>批量更新所有有效订单的成本和利润 — 对应 CE_UpdateAllOrderCosts()</summary>
    Task UpdateAllOrderCostsAsync(CancellationToken ct = default);

    // ==================== 摘要 ====================

    /// <summary>获取成本传导状态摘要 — 对应 CE_GetCostSummary()</summary>
    Task<CostSummaryDto> GetCostSummaryAsync(CancellationToken ct = default);

    /// <summary>清除引擎缓存（预加载后重新计算时使用）</summary>
    void ClearCache();
}

/// <summary>成本摘要 DTO</summary>
public class CostSummaryDto
{
    public int TotalProducts { get; set; }
    public int UpdatedProducts { get; set; }
    public int TotalOrders { get; set; }
    public int UpdatedOrders { get; set; }
    public DateTime? LastUpdate { get; set; }
    /// <summary>缓存中原材料条目数</summary>
    public int CachedMaterials { get; set; }
    /// <summary>缓存中加权批次条目数</summary>
    public int CachedBatchCosts { get; set; }
}
