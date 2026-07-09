namespace PerfumeShop.Core.Interfaces;

/// <summary>
/// 成本引擎接口 — 三级成本自动传导：原材料 → 香调 → 产品 → 订单
/// 对应 ASP 中的 cost_engine.asp
/// </summary>
public interface ICostEngine
{
    /// <summary>计算单个原料成本 (加权平均/供应商报价回退)</summary>
    Task<decimal> CalculateMaterialCostAsync(int materialId, CancellationToken ct = default);

    /// <summary>计算单个香调成本 (Accord配方或成分聚合)</summary>
    Task<decimal> CalculateNoteCostAsync(int noteId, CancellationToken ct = default);

    /// <summary>计算产品 BOM 成本 (香调配比 + 瓶身)</summary>
    Task<decimal> CalculateProductBomCostAsync(int productId, CancellationToken ct = default);

    /// <summary>计算产品完整单位成本 (BOM + 包装 + 人工) — 品牌定香使用采购成本</summary>
    Task<decimal> CalculateProductUnitCostAsync(int productId, CancellationToken ct = default);

    /// <summary>更新指定产品的成本到数据库</summary>
    Task UpdateProductCostAsync(int productId, CancellationToken ct = default);

    /// <summary>批量更新所有启用产品的成本</summary>
    Task<int> UpdateAllProductCostsAsync(CancellationToken ct = default);

    /// <summary>更新订单的成本和利润</summary>
    Task UpdateOrderCostsAsync(int orderId, CancellationToken ct = default);

    /// <summary>批量更新所有有效订单的成本和利润</summary>
    Task UpdateAllOrderCostsAsync(CancellationToken ct = default);

    /// <summary>获取成本传导状态摘要</summary>
    Task<CostSummaryDto> GetCostSummaryAsync(CancellationToken ct = default);
}

/// <summary>成本摘要 DTO</summary>
public class CostSummaryDto
{
    public int TotalProducts { get; set; }
    public int UpdatedProducts { get; set; }
    public int TotalOrders { get; set; }
    public int UpdatedOrders { get; set; }
    public DateTime? LastUpdate { get; set; }
}
