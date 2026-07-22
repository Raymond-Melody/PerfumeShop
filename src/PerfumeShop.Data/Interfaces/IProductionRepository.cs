using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 生产仓储接口 — 生产工单、质检、排程、库存
/// </summary>
public interface IProductionRepository : IRepository<ProductionOrder>
{
    /// <summary>获取生产工单列表（分页+筛选）</summary>
    Task<(IEnumerable<ProductionOrder> Items, int TotalCount)> GetProductionOrdersAsync(
        string? status = null, string? search = null,
        int page = 1, int pageSize = 20, CancellationToken ct = default);

    /// <summary>按ID获取工单详情（含日志）</summary>
    Task<ProductionOrder?> GetProductionOrderDetailAsync(int productionId, CancellationToken ct = default);

    /// <summary>获取工单关联的生产日志</summary>
    Task<IEnumerable<ProductionLog>> GetProductionLogsAsync(int productionId, CancellationToken ct = default);

    /// <summary>同步已付款订单→生产工单（事务+幂等）</summary>
    Task<(int Synced, int Errors, string Message)> SyncProductionOrdersAsync(CancellationToken ct = default);

    /// <summary>更新生产工单状态</summary>
    Task<bool> UpdateProductionStatusAsync(int productionId, string newStatus, string? operatorName = null, CancellationToken ct = default);

    /// <summary>修复生产工单状态（中文→英文）</summary>
    Task<(int Updated, string Message)> FixProductionStatusAsync(CancellationToken ct = default);

    /// <summary>获取质检记录列表</summary>
    Task<(IEnumerable<ProductionOrder> Items, int TotalCount)> GetQualityChecksAsync(
        string? status = null, int page = 1, int pageSize = 20, CancellationToken ct = default);

    /// <summary>获取生产报表数据（按日期分组产量）</summary>
    Task<IEnumerable<ProductionOrder>> GetProductionReportDataAsync(
        DateTime startDate, DateTime endDate, CancellationToken ct = default);

    /// <summary>获取成品库存列表</summary>
    Task<IEnumerable<ProductInventory>> GetProductInventoryAsync(CancellationToken ct = default);

    /// <summary>获取瓶子库存列表</summary>
    Task<IEnumerable<BottleInventory>> GetBottleInventoryAsync(CancellationToken ct = default);

    /// <summary>获取包装库存列表</summary>
    Task<IEnumerable<PackagingInventory>> GetPackagingInventoryAsync(CancellationToken ct = default);
}
