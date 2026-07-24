using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 采购模块仓储 — 封装采购相关实体的查询与业务操作
/// </summary>
public class PurchaseRepository
{
    private readonly PerfumeShopContext _db;
    private readonly IInventoryLedger _ledger;

    public PurchaseRepository(PerfumeShopContext db, IInventoryLedger ledger)
    {
        _db = db ?? throw new ArgumentNullException(nameof(db));
        _ledger = ledger;
    }

    // ==================== 采购订单 ====================

    /// <summary>获取采购订单列表（支持状态/类型/供应商筛选）</summary>
    public async Task<(List<PurchaseOrder> Items, int Total)> GetPurchaseOrdersAsync(
        string? status = null, string? orderType = null, int? supplierId = null,
        string? search = null, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        var q = _db.PurchaseOrders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(o => o.Status == status);
        if (!string.IsNullOrWhiteSpace(orderType)) q = q.Where(o => o.OrderType == orderType);
        if (supplierId.HasValue) q = q.Where(o => o.SupplierId == supplierId);
        if (!string.IsNullOrWhiteSpace(search))
            q = q.Where(o => (o.PurchaseNo ?? "").Contains(search) || (o.Remarks ?? "").Contains(search));
        var total = await q.CountAsync(ct);
        var items = await q.OrderByDescending(o => o.CreatedAt).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    /// <summary>获取单个采购订单</summary>
    public async Task<PurchaseOrder?> GetPurchaseOrderAsync(int id, CancellationToken ct = default)
        => await _db.PurchaseOrders.FirstOrDefaultAsync(o => o.PurchaseId == id, ct);

    /// <summary>获取采购订单明细</summary>
    public async Task<List<PurchaseOrderDetail>> GetPurchaseOrderDetailsAsync(int purchaseId, CancellationToken ct = default)
        => await _db.PurchaseOrderDetails.AsNoTracking().Where(d => d.PurchaseId == purchaseId).ToListAsync(ct);

    /// <summary>创建采购订单</summary>
    public async Task<PurchaseOrder> CreatePurchaseOrderAsync(PurchaseOrder order, List<PurchaseOrderDetail>? details = null, CancellationToken ct = default)
    {
        order.CreatedAt = DateTime.Now;
        order.Status = "draft";
        order.PurchaseNo = $"PO-{DateTime.Now:yyyyMMddHHmmss}";
        var entry = await _db.PurchaseOrders.AddAsync(order, ct);
        if (details?.Count > 0)
        {
            foreach (var d in details) d.PurchaseId = order.PurchaseId;
            await _db.PurchaseOrderDetails.AddRangeAsync(details, ct);
        }
        await _db.SaveChangesAsync(ct);
        return entry.Entity;
    }

    /// <summary>审批采购订单</summary>
    public async Task<bool> ApprovePurchaseOrderAsync(int id, int approvedBy, CancellationToken ct = default)
    {
        var order = await _db.PurchaseOrders.FindAsync(new object[] { id }, ct);
        if (order == null) return false;
        order.Status = "approved";
        order.ApprovedBy = approvedBy;
        order.ApprovedAt = DateTime.Now;
        order.UpdatedAt = DateTime.Now;
        _db.PurchaseOrderStatusLogs.Add(new PurchaseOrderStatusLog
        {
            PurchaseId = id, FromStatus = "draft", ToStatus = "approved",
            ChangedBy = approvedBy.ToString(), ChangedAt = DateTime.Now
        });
        await _db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>取消采购订单</summary>
    public async Task<bool> CancelPurchaseOrderAsync(int id, string? reason = null, CancellationToken ct = default)
    {
        var order = await _db.PurchaseOrders.FindAsync(new object[] { id }, ct);
        if (order == null) return false;
        var oldStatus = order.Status;
        order.Status = "cancelled";
        order.UpdatedAt = DateTime.Now;
        order.Remarks = reason ?? order.Remarks;
        _db.PurchaseOrderStatusLogs.Add(new PurchaseOrderStatusLog
        {
            PurchaseId = id, FromStatus = oldStatus, ToStatus = "cancelled",
            ChangedAt = DateTime.Now, Remarks = reason
        });
        await _db.SaveChangesAsync(ct);
        return true;
    }

    /// <summary>批量审批采购订单</summary>
    public async Task<int> BatchApproveAsync(IEnumerable<int> ids, int approvedBy, CancellationToken ct = default)
    {
        var orders = await _db.PurchaseOrders.Where(o => ids.Contains(o.PurchaseId) && o.Status == "draft").ToListAsync(ct);
        foreach (var o in orders)
        {
            o.Status = "approved"; o.ApprovedBy = approvedBy; o.ApprovedAt = DateTime.Now; o.UpdatedAt = DateTime.Now;
            _db.PurchaseOrderStatusLogs.Add(new PurchaseOrderStatusLog
            {
                PurchaseId = o.PurchaseId, FromStatus = "draft", ToStatus = "approved",
                ChangedBy = approvedBy.ToString(), ChangedAt = DateTime.Now
            });
        }
        return await _db.SaveChangesAsync(ct);
    }

    // ==================== 收货入库 ====================

    /// <summary>获取收货单列表</summary>
    public async Task<(List<PurchaseReceipt> Items, int Total)> GetReceiptsAsync(
        int? supplierId = null, string? status = null, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        var q = _db.PurchaseReceipts.AsNoTracking().AsQueryable();
        if (supplierId.HasValue) q = q.Where(r => r.SupplierId == supplierId);
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(r => r.Status == status);
        var total = await q.CountAsync(ct);
        var items = await q.OrderByDescending(r => r.CreatedAt).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    /// <summary>
    /// 创建收货单并更新库存 — 对标 V18 admin/purchase/receiving.asp
    /// 事务内：建收货单+明细 → 按 OrderType 更新对应库存 StockQty + 移动加权 WeightedUnitCost
    ///            → 写批次 + IN 流水 → 回写采购明细 ReceivedQty + 订单状态
    /// orderType: RawMaterial | Packaging | Bottle | Printing | SprayHead
    /// </summary>
    public async Task<PurchaseReceipt> ReceivePurchaseAsync(
        PurchaseReceipt receipt, List<PurchaseReceiptDetail>? details = null,
        string orderType = "RawMaterial", string? operatorName = null, CancellationToken ct = default)
    {
        using var tx = await _db.Database.BeginTransactionAsync(ct);
        try
        {
            receipt.CreatedAt = DateTime.Now;
            if (string.IsNullOrWhiteSpace(receipt.Status)) receipt.Status = "Complete";
            if (string.IsNullOrWhiteSpace(receipt.ReceiptNo))
                receipt.ReceiptNo = $"RCV-{DateTime.Now:yyyyMMddHHmmss}";
            var entry = await _db.PurchaseReceipts.AddAsync(receipt, ct);
            await _db.SaveChangesAsync(ct);

            if (details?.Count > 0)
            {
                foreach (var d in details) d.ReceiptId = receipt.ReceiptId;
                await _db.PurchaseReceiptDetails.AddRangeAsync(details, ct);
                await _db.SaveChangesAsync(ct);

                foreach (var d in details)
                {
                    var matId = d.MaterialId ?? 0;
                    var accepted = (decimal)(d.AcceptedQty ?? d.ReceivedQty ?? 0);
                    var unitPrice = d.UnitPrice ?? 0;
                    if (matId <= 0 || accepted <= 0) continue;

                    await ApplyStockInAsync(orderType, matId, accepted, unitPrice, receipt.ReceiptNo!, operatorName, ct);

                    // 回写采购明细已收数量
                    if (d.PurchaseDetailId.HasValue)
                        await _db.Database.ExecuteSqlInterpolatedAsync(
                            $"UPDATE PurchaseOrderDetails SET ReceivedQty = COALESCE(ReceivedQty,0) + {accepted} WHERE DetailID = {d.PurchaseDetailId.Value}", ct);
                }
            }

            // 更新采购订单状态
            if (receipt.PurchaseId.HasValue)
            {
                var poStatus = receipt.Status == "Partial" ? "PartialReceived" : "Received";
                var now = DateTime.Now;
                await _db.Database.ExecuteSqlInterpolatedAsync(
                    $"UPDATE PurchaseOrders SET Status = {poStatus}, UpdatedAt = {now} WHERE PurchaseID = {receipt.PurchaseId.Value}", ct);
            }

            await tx.CommitAsync(ct);
            return entry.Entity;
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }
    }

    /// <summary>OrderType → 库存表/主键列映射 (对标 V18 receiving.asp L326-393)</summary>
    private static (string Table, string IdCol, string ItemType) StockTableOf(string orderType) => orderType switch
    {
        "Packaging" => ("PackagingInventory", "PackagingID", "Packaging"),
        "Bottle" => ("BottleStyles", "BottleID", "Bottle"),
        "Printing" => ("PrintingInventory", "PrintingID", "Printing"),
        "SprayHead" => ("SprayHeadInventory", "SprayHeadID", "SprayHead"),
        _ => ("RawMaterialInventory", "MaterialID", "RawMaterial"),
    };

    /// <summary>入库单条：移动加权成本重算 + 更新库存 + 批次 + IN 流水</summary>
    private async Task ApplyStockInAsync(string orderType, int matId, decimal accepted, decimal unitPrice,
        string receiptNo, string? operatorName, CancellationToken ct)
    {
        var (table, idCol, itemType) = StockTableOf(orderType);

        // 读旧库存量与旧加权成本（表名/列名受控枚举，非用户输入）
        var rows = await _db.Database.SqlQueryRaw<StockCostRow>(
            $"SELECT COALESCE(StockQty,0) AS OldStock, COALESCE(WeightedUnitCost,0) AS OldCost, COALESCE(ItemName,'') AS ItemName, COALESCE(ItemCode,'') AS ItemCode FROM [{table}] WHERE [{idCol}] = {{0}}",
            matId).ToListAsync(ct);
        var old = rows.FirstOrDefault() ?? new StockCostRow();
        var oldStock = old.OldStock < 0 ? 0 : old.OldStock;

        // 移动加权：(旧量*旧价 + 收货量*单价) / (旧量+收货量)
        var newCost = (oldStock + accepted) > 0
            ? (oldStock * old.OldCost + accepted * unitPrice) / (oldStock + accepted)
            : unitPrice;

        var nowTs = DateTime.Now;
        // 更新库存量 + 加权成本
        await _db.Database.ExecuteSqlRawAsync(
            $"UPDATE [{table}] SET StockQty = COALESCE(StockQty,0) + {{0}}, WeightedUnitCost = {{1}}, UpdatedAt = {{2}} WHERE [{idCol}] = {{3}}",
            new object[] { accepted, newCost, nowTs, matId }, ct);

        // 批次记录（实际采购单价，保留批次差异化成本）
        try
        {
            await _db.Database.ExecuteSqlInterpolatedAsync($@"
INSERT INTO InventoryBatches (ItemType, ItemID, ItemCode, ItemName, BatchNo, UnitCost, StockQty, UpdatedAt, CreatedAt)
VALUES ({itemType}, {matId}, {old.ItemCode}, {old.ItemName}, {receiptNo}, {unitPrice}, {accepted}, {nowTs}, {nowTs})", ct);
        }
        catch { /* InventoryBatches 可能列不齐，不阻断 */ }

        // IN 流水（NoteID 非香调交易传 0，MaterialID 存品类主键）
        await _ledger.WriteTransactionAsync(new InvTxn(
            NoteId: 0, MaterialId: matId, ProductId: null, Quantity: accepted,
            TransactionType: "采购入库", Direction: "IN", ReferenceType: "PurchaseReceipt",
            ReferenceOrderId: null, UnitCost: unitPrice,
            Notes: $"收货单{receiptNo}", CreatedBy: operatorName ?? "SYSTEM"), ct);

        await _ledger.WriteMovementAsync(new StockMove(
            ItemType: itemType, ItemId: matId, ItemName: old.ItemName, ItemCode: old.ItemCode,
            MovementType: "IN", Quantity: accepted, BeforeQty: oldStock, AfterQty: oldStock + accepted,
            Unit: null, ReferenceNo: receiptNo, Notes: "采购入库", CreatedBy: operatorName ?? "SYSTEM"), ct);
    }

    private sealed class StockCostRow
    {
        public decimal OldStock { get; set; }
        public decimal OldCost { get; set; }
        public string ItemName { get; set; } = "";
        public string ItemCode { get; set; } = "";
    }

    /// <summary>获取收货单明细</summary>
    public async Task<List<PurchaseReceiptDetail>> GetReceiptDetailsAsync(int receiptId, CancellationToken ct = default)
        => await _db.PurchaseReceiptDetails.AsNoTracking().Where(d => d.ReceiptId == receiptId).ToListAsync(ct);

    // ==================== 供应商管理 ====================

    /// <summary>获取供应商列表</summary>
    public async Task<List<Supplier>> GetSuppliersAsync(string? category = null, bool? activeOnly = null, CancellationToken ct = default)
    {
        var q = _db.Suppliers.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(category)) q = q.Where(s => s.Category == category);
        if (activeOnly == true) q = q.Where(s => s.IsActive == true);
        return await q.OrderBy(s => s.SupplierName).ToListAsync(ct);
    }

    /// <summary>获取单个供应商</summary>
    public async Task<Supplier?> GetSupplierAsync(int id, CancellationToken ct = default)
        => await _db.Suppliers.FirstOrDefaultAsync(s => s.SupplierId == id, ct);

    /// <summary>创建/更新供应商</summary>
    public async Task<Supplier> UpsertSupplierAsync(Supplier supplier, CancellationToken ct = default)
    {
        if (supplier.SupplierId == 0)
        {
            supplier.CreatedAt = DateTime.Now;
            var entry = await _db.Suppliers.AddAsync(supplier, ct);
            await _db.SaveChangesAsync(ct);
            return entry.Entity;
        }
        _db.Suppliers.Update(supplier);
        await _db.SaveChangesAsync(ct);
        return supplier;
    }

    /// <summary>获取供应商评估记录</summary>
    public async Task<List<SupplierEvaluation>> GetSupplierEvaluationsAsync(int supplierId, CancellationToken ct = default)
        => await _db.SupplierEvaluations.AsNoTracking().Where(e => e.SupplierId == supplierId)
            .OrderByDescending(e => e.EvaluationDate).ToListAsync(ct);

    /// <summary>获取供应商合同</summary>
    public async Task<List<SupplierContract>> GetSupplierContractsAsync(int supplierId, CancellationToken ct = default)
        => await _db.SupplierContracts.AsNoTracking().Where(c => c.SupplierId == supplierId)
            .OrderByDescending(c => c.CreatedAt).ToListAsync(ct);

    // ==================== 价格管理 ====================

    /// <summary>获取供应商报价列表</summary>
    public async Task<List<SupplierPrice>> GetSupplierPricesAsync(int? supplierId = null, CancellationToken ct = default)
    {
        var q = _db.SupplierPrices.AsNoTracking().AsQueryable();
        if (supplierId.HasValue) q = q.Where(p => p.SupplierId == supplierId);
        return await q.OrderByDescending(p => p.CreatedAt).ToListAsync(ct);
    }

    /// <summary>更新供应商报价</summary>
    public async Task<SupplierPrice> UpsertPriceAsync(SupplierPrice price, CancellationToken ct = default)
    {
        if (price.PriceId == 0)
        {
            price.CreatedAt = DateTime.Now;
            var entry = await _db.SupplierPrices.AddAsync(price, ct);
            await _db.SaveChangesAsync(ct);
            return entry.Entity;
        }
        _db.SupplierPrices.Update(price);
        await _db.SaveChangesAsync(ct);
        return price;
    }

    // ==================== 采购批次 ====================

    /// <summary>获取采购批次列表</summary>
    public async Task<(List<PurchaseBatch> Items, int Total)> GetPurchaseBatchesAsync(
        int? supplierId = null, string? itemType = null, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        var q = _db.PurchaseBatches.AsNoTracking().AsQueryable();
        if (supplierId.HasValue) q = q.Where(b => b.SupplierId == supplierId);
        if (!string.IsNullOrWhiteSpace(itemType)) q = q.Where(b => b.ItemType == itemType);
        var total = await q.CountAsync(ct);
        var items = await q.OrderByDescending(b => b.CreatedAt).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    /// <summary>获取单个批次详情</summary>
    public async Task<PurchaseBatch?> GetBatchAsync(int id, CancellationToken ct = default)
        => await _db.PurchaseBatches.FirstOrDefaultAsync(b => b.BatchId == id, ct);

    // ==================== 品牌定香采购 ====================

    /// <summary>获取品牌定香产品列表</summary>
    public async Task<List<FixedBrandProduct>> GetFixedBrandProductsAsync(CancellationToken ct = default)
        => await _db.FixedBrandProducts.AsNoTracking().OrderBy(p => p.ProductName).ToListAsync(ct);

    /// <summary>获取品牌定香采购订单</summary>
    public async Task<(List<FixedBrandPurchaseOrder> Items, int Total)> GetFixedBrandOrdersAsync(
        string? status = null, int page = 1, int pageSize = 20, CancellationToken ct = default)
    {
        var q = _db.FixedBrandPurchaseOrders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(o => o.Status == status);
        var total = await q.CountAsync(ct);
        var items = await q.OrderByDescending(o => o.CreatedAt).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    /// <summary>获取品牌定香库存</summary>
    public async Task<List<FixedBrandInventory>> GetFixedBrandInventoriesAsync(CancellationToken ct = default)
        => await _db.FixedBrandInventories.AsNoTracking().ToListAsync(ct);

    /// <summary>获取品牌定香成本分配</summary>
    public async Task<List<FixedBrandCostAllocation>> GetFixedBrandCostAllocationsAsync(CancellationToken ct = default)
        => await _db.FixedBrandCostAllocations.AsNoTracking().OrderByDescending(a => a.AllocatedAt).ToListAsync(ct);

    // ==================== 成本与报表 ====================

    /// <summary>获取采购成本审查</summary>
    public async Task<List<PurchaseCostReview>> GetCostReviewsAsync(CancellationToken ct = default)
        => await _db.PurchaseCostReviews.AsNoTracking().OrderByDescending(r => r.CreatedAt).ToListAsync(ct);

    /// <summary>获取采购历史统计</summary>
    public async Task<List<PurchaseHistoryStat>> GetHistoryStatsAsync(CancellationToken ct = default)
        => await _db.PurchaseHistoryStats.AsNoTracking().ToListAsync(ct);

    /// <summary>获取采购分类</summary>
    public async Task<List<PurchaseCategory>> GetCategoriesAsync(CancellationToken ct = default)
        => await _db.PurchaseCategories.AsNoTracking().OrderBy(c => c.DisplayOrder).ToListAsync(ct);

    /// <summary>获取订单状态日志</summary>
    public async Task<List<PurchaseOrderStatusLog>> GetStatusLogsAsync(int purchaseId, CancellationToken ct = default)
        => await _db.PurchaseOrderStatusLogs.AsNoTracking().Where(l => l.PurchaseId == purchaseId)
            .OrderByDescending(l => l.ChangedAt).ToListAsync(ct);

    /// <summary>获取库存预警（低于安全库存的品类）</summary>
    public async Task<List<PackagingInventory>> GetLowStockPackagingAsync(CancellationToken ct = default)
        => await _db.PackagingInventories.AsNoTracking()
            .Where(p => p.StockQty < p.SafetyStock).ToListAsync(ct);

    /// <summary>获取瓶子库存列表</summary>
    public async Task<List<BottleInventory>> GetBottleInventoriesAsync(CancellationToken ct = default)
        => await _db.BottleInventories.AsNoTracking().ToListAsync(ct);
}
