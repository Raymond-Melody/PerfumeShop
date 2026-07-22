using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 半成品仓储接口 — 调配记录、半成品转移、库存预警
/// </summary>
public interface ISemiFinishedRepository
{
    // ===== 调配记录 (AccordProduction) =====

    /// <summary>获取调配记录列表（分页）</summary>
    Task<(IEnumerable<AccordProduction> Items, int TotalCount)> GetAccordProductionsAsync(
        string? status = null, int page = 1, int pageSize = 20, CancellationToken ct = default);

    /// <summary>获取调配记录详情（含明细）</summary>
    Task<AccordProduction?> GetAccordProductionDetailAsync(int productionId, CancellationToken ct = default);

    /// <summary>获取调配明细</summary>
    Task<IEnumerable<AccordProductionDetail>> GetAccordProductionDetailsAsync(int productionId, CancellationToken ct = default);

    /// <summary>新增调配记录</summary>
    Task<AccordProduction> CreateBlendRecordAsync(AccordProduction record, CancellationToken ct = default);

    /// <summary>更新调配记录状态</summary>
    Task<bool> UpdateBlendStatusAsync(int productionId, string newStatus, CancellationToken ct = default);

    // ===== 半成品库存 (NoteInventory + RawMaterialInventory) =====

    /// <summary>获取半成品库存列表（香料）</summary>
    Task<IEnumerable<NoteInventory>> GetNoteInventoryAsync(CancellationToken ct = default);

    /// <summary>获取原料库存列表</summary>
    Task<IEnumerable<RawMaterialInventory>> GetRawMaterialInventoryAsync(CancellationToken ct = default);

    /// <summary>获取库存预警（低于安全库存）</summary>
    Task<IEnumerable<object>> GetInventoryAlertsAsync(CancellationToken ct = default);

    // ===== 车间转移 (WorkshopTransfer) =====

    /// <summary>获取转移记录列表</summary>
    Task<(IEnumerable<WorkshopTransfer> Items, int TotalCount)> GetWorkshopTransfersAsync(
        string? status = null, int page = 1, int pageSize = 20, CancellationToken ct = default);

    /// <summary>执行半成品转移</summary>
    Task<WorkshopTransfer> TransferSemiFinishedAsync(WorkshopTransfer transfer, CancellationToken ct = default);

    /// <summary>完成转移（确认接收）</summary>
    Task<bool> FulfillTransferAsync(int transferId, CancellationToken ct = default);

    /// <summary>获取基础香料库存列表</summary>
    Task<IEnumerable<BaseNote>> GetBaseNotesAsync(CancellationToken ct = default);

    /// <summary>获取半成品报表数据</summary>
    Task<IEnumerable<AccordProduction>> GetReportDataAsync(DateTime start, DateTime end, CancellationToken ct = default);
}
