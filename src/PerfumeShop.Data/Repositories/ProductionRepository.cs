using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 生产仓储实现 — V19 M4-C
/// </summary>
public class ProductionRepository : Repository<ProductionOrder>, IProductionRepository
{
    private readonly IInventoryLedger _ledger;
    public ProductionRepository(PerfumeShopContext context, IInventoryLedger ledger) : base(context)
    {
        _ledger = ledger;
    }

    public override async Task<ProductionOrder?> GetByIdAsync(int id, CancellationToken ct = default)
        => await _dbSet.FirstOrDefaultAsync(p => p.ProductionId == id, ct);

    public async Task<(IEnumerable<ProductionOrder> Items, int TotalCount)> GetProductionOrdersAsync(
        string? status, string? search, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _dbSet.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status))
            query = query.Where(p => p.Status == status);
        if (!string.IsNullOrWhiteSpace(search))
        {
            var s = search.Trim();
            query = query.Where(p => (p.WorkOrderNo != null && p.WorkOrderNo.Contains(s))
                                     || p.OrderId.ToString().Contains(s));
        }
        int total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(p => p.CreatedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<ProductionOrder?> GetProductionOrderDetailAsync(int productionId, CancellationToken ct = default)
        => await _dbSet.AsNoTracking().FirstOrDefaultAsync(p => p.ProductionId == productionId, ct);

    public async Task<IEnumerable<ProductionLog>> GetProductionLogsAsync(int productionId, CancellationToken ct = default)
        => await _context.ProductionLogs.AsNoTracking()
            .Where(l => l.ProductionId == productionId)
            .OrderByDescending(l => l.CreatedAt).ToListAsync(ct);

    public async Task<(int Synced, int Errors, string Message)> SyncProductionOrdersAsync(CancellationToken ct = default)
    {
        using var transaction = await _context.Database.BeginTransactionAsync(ct);
        try
        {
            // 幂等检查：已付款/处理中且无工单的订单
            var syncedOrderIds = await _context.ProductionOrders.AsNoTracking()
                .Where(po => po.Status != "Cancelled")
                .Select(po => po.OrderId).Distinct().ToListAsync(ct);

            var pendingOrders = await _context.Orders.AsNoTracking()
                .Where(o => (o.Status == "Paid" || o.Status == "Processing")
                            && !syncedOrderIds.Contains(o.OrderId))
                .OrderBy(o => o.OrderId).ToListAsync(ct);

            if (pendingOrders.Count == 0)
            {
                await transaction.RollbackAsync(ct);
                return (0, 0, "所有订单都已同步，无需操作");
            }

            int synced = 0, errors = 0;
            foreach (var order in pendingOrders)
            {
                // V21: 排产分流 — standard(品牌定香)走成品库存不排产，仅为 custom/kol 明细建工单
                var details = await (from d in _context.OrderDetails.AsNoTracking()
                                     join p in _context.Products.AsNoTracking() on d.ProductId equals p.ProductId into pj
                                     from p in pj.DefaultIfEmpty()
                                     where d.OrderId == order.OrderId
                                           && (p == null || p.ProductType == null || p.ProductType != "standard")
                                     select d).ToListAsync(ct);
                if (details.Count == 0) continue;

                int bottleIndex = 0;
                int totalBottles = details.Sum(d => d.Quantity);
                string prefix = $"WO-{DateTime.Now:yyyyMMdd}-";

                foreach (var detail in details)
                {
                    for (int i = 1; i <= detail.Quantity; i++)
                    {
                        bottleIndex++;
                        var workOrderNo = $"{prefix}{bottleIndex:D4}";

                        // 幂等：工单号已存在则跳过
                        bool exists = await _dbSet.AnyAsync(p => p.WorkOrderNo == workOrderNo && p.OrderId == order.OrderId, ct);
                        if (exists) { synced++; continue; }

                        var po = new ProductionOrder
                        {
                            OrderId = order.OrderId,
                            DetailId = detail.DetailId,
                            WorkOrderNo = workOrderNo,
                            BottleIndex = bottleIndex,
                            TotalBottles = totalBottles,
                            Status = "Pending",
                            Priority = 0,
                            CreatedAt = DateTime.Now,
                            UpdatedAt = DateTime.Now
                        };
                        await AddAsync(po, ct);
                        await SaveChangesAsync(ct);

                        // 写入日志
                        var log = new ProductionLog
                        {
                            ProductionId = po.ProductionId,
                            Status = "Pending",
                            Notes = $"系统同步创建 (订单{order.OrderNo} 第{bottleIndex}瓶/共{detail.Quantity}瓶)",
                            CreatedBy = "SYSTEM_SYNC",
                            CreatedAt = DateTime.Now
                        };
                        await _context.ProductionLogs.AddAsync(log, ct);
                        await _context.SaveChangesAsync(ct);
                        synced++;
                    }
                }

                // 订单状态→Processing
                var ord = await _context.Orders.FirstOrDefaultAsync(o => o.OrderId == order.OrderId, ct);
                if (ord != null && ord.Status == "Paid") { ord.Status = "Processing"; ord.UpdatedAt = DateTime.Now; }
                await _context.SaveChangesAsync(ct);
            }

            await transaction.CommitAsync(ct);
            return (synced, errors, $"成功同步 {synced} 个生产工单");
        }
        catch (Exception ex)
        {
            await transaction.RollbackAsync(ct);
            return (0, 1, $"同步失败: {ex.Message}");
        }
    }

    public async Task<bool> UpdateProductionStatusAsync(int productionId, string newStatus, string? operatorName = null, CancellationToken ct = default)
    {
        var po = await _dbSet.FirstOrDefaultAsync(p => p.ProductionId == productionId, ct);
        if (po == null) return false;

        var oldStatus = po.Status;
        po.Status = newStatus;
        po.UpdatedAt = DateTime.Now;
        if (newStatus == "Completed") po.CompletedAt = DateTime.Now;
        if (newStatus == "InProgress") po.StartedAt = DateTime.Now;

        _context.ProductionLogs.Add(new ProductionLog
        {
            ProductionId = productionId,
            Status = newStatus,
            Notes = $"{oldStatus} → {newStatus}",
            CreatedBy = operatorName ?? "Admin",
            CreatedAt = DateTime.Now
        });

        await SaveChangesAsync(ct);
        return true;
    }

    public async Task<(int Updated, string Message)> FixProductionStatusAsync(CancellationToken ct = default)
    {
        var mappings = new Dictionary<string, string>
        {
            ["待排产"] = "Pending", ["生产中"] = "InProgress",
            ["已完成"] = "Completed", ["已取消"] = "Cancelled", ["已质检"] = "QC_Review"
        };
        int total = 0;
        foreach (var (cn, en) in mappings)
        {
            var items = await _dbSet.Where(p => p.Status == cn).ToListAsync(ct);
            foreach (var item in items) item.Status = en;
            total += items.Count;

            var logs = await _context.ProductionLogs.Where(l => l.Status == cn).ToListAsync(ct);
            foreach (var log in logs) log.Status = en;
            total += logs.Count;
        }
        await SaveChangesAsync(ct);
        return (total, $"已修复 {total} 条状态记录");
    }

    public async Task<(IEnumerable<ProductionOrder> Items, int TotalCount)> GetQualityChecksAsync(
        string? status, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _dbSet.AsNoTracking().Where(p => p.Status == "Completed" || p.Status == "QC_Review" || p.Status == "QC_Fail");
        if (!string.IsNullOrWhiteSpace(status)) query = query.Where(p => p.Status == status);
        int total = await query.CountAsync(ct);
        var items = await query.OrderByDescending(p => p.CompletedAt)
            .Skip((page - 1) * pageSize).Take(pageSize).ToListAsync(ct);
        return (items, total);
    }

    public async Task<IEnumerable<ProductionOrder>> GetProductionReportDataAsync(
        DateTime startDate, DateTime endDate, CancellationToken ct = default)
        => await _dbSet.AsNoTracking()
            .Where(p => p.CreatedAt >= startDate && p.CreatedAt <= endDate)
            .OrderBy(p => p.CreatedAt).ToListAsync(ct);

    public async Task<IEnumerable<ProductInventory>> GetProductInventoryAsync(CancellationToken ct = default)
        => await _context.ProductInventories.AsNoTracking().OrderBy(p => p.ProductId).ToListAsync(ct);

    public async Task<IEnumerable<BottleInventory>> GetBottleInventoryAsync(CancellationToken ct = default)
        => await _context.BottleInventories.AsNoTracking().OrderBy(b => b.BottleId).ToListAsync(ct);

    public async Task<IEnumerable<PackagingInventory>> GetPackagingInventoryAsync(CancellationToken ct = default)
        => await _context.PackagingInventories.AsNoTracking().OrderBy(p => p.PackagingId).ToListAsync(ct);

    // ===== V21: 成品入库（按配方扣香调 + 扣瓶身 + 流水）=====
    // 对标 V18 admin/prodcenter/prod_warehouse.asp warehouse_in

    public record WarehouseInResult(bool Success, string Message);

    /// <summary>
    /// 工单入库：经 DetailId→OrderDetails(ProductId,VolumeML)→最新 RecipeProducts/RecipeProductNotes，
    /// 按 消耗ml=(Pct/100)×VolumeML 扣 NoteInventory，扣 1 个绑定瓶身，写 OUT 流水，更新工单状态。
    /// </summary>
    public async Task<WarehouseInResult> WarehouseInAsync(int productionId, string? operatorName, CancellationToken ct = default)
    {
        if (productionId <= 0) return new(false, "无效工单");

        var po = await _dbSet.AsNoTracking().FirstOrDefaultAsync(p => p.ProductionId == productionId, ct);
        if (po == null) return new(false, "工单不存在");

        // 1) 定位工单对应产品与单瓶容量
        var info = await (from p in _context.ProductionOrders.AsNoTracking()
                          join od in _context.OrderDetails.AsNoTracking() on p.DetailId equals od.DetailId into oj
                          from od in oj.DefaultIfEmpty()
                          where p.ProductionId == productionId
                          select new { ProductId = od != null ? od.ProductId : 0, VolumeMl = od != null ? (od.VolumeMl ?? 0) : 0 })
                         .FirstOrDefaultAsync(ct);
        int productId = info?.ProductId ?? 0;
        decimal volumeMl = info != null && info.VolumeMl > 0 ? info.VolumeMl : 50;

        using var tx = await _context.Database.BeginTransactionAsync(ct);
        try
        {
            // 2) 取该产品最新已发布产品配方的香调配比 + 香调加权成本
            if (productId > 0)
            {
                var notes = await (from rp in _context.RecipeProducts.AsNoTracking()
                                   where rp.ProductId == productId && rp.Status == "Published"
                                   orderby rp.PublishedAt descending
                                   select rp.ProductRecipeId).Take(1)
                                  .Join(_context.RecipeProductNotes.AsNoTracking(),
                                        prid => prid, rpn => rpn.ProductRecipeId,
                                        (prid, rpn) => rpn).ToListAsync(ct);

                foreach (var n in notes)
                {
                    int noteId = n.NoteId ?? 0;
                    decimal pct = (decimal)(n.Percentage ?? 0);
                    var consume = (pct / 100m) * volumeMl;
                    if (noteId <= 0 || consume <= 0) continue;

                    // 取香调加权成本(用于流水成本)
                    var wuc = await _context.Database.SqlQueryRaw<decimal>(
                        "SELECT CAST(COALESCE(WeightedUnitCost,0) AS decimal(19,4)) AS Value FROM NoteInventory WHERE NoteID = {0}", noteId)
                        .ToListAsync(ct);
                    var noteCost = wuc.FirstOrDefault();

                    await _context.Database.ExecuteSqlInterpolatedAsync(
                        $"UPDATE NoteInventory SET StockQuantity = COALESCE(StockQuantity,0) - {consume}, UpdatedAt = {DateTime.Now} WHERE NoteID = {noteId}", ct);

                    await _ledger.WriteTransactionAsync(new InvTxn(
                        NoteId: noteId, MaterialId: null, ProductId: productId, Quantity: -consume,
                        TransactionType: "生产领用", Direction: "OUT", ReferenceType: "ProductionOrder",
                        ReferenceOrderId: productionId, UnitCost: noteCost,
                        Notes: $"工单入库消耗香调 PO#{productionId}", CreatedBy: operatorName ?? "SYSTEM"), ct);
                }

                // 3) 扣绑定瓶身库存（每瓶1个）
                await _context.Database.ExecuteSqlInterpolatedAsync(
                    $"UPDATE BottleStyles SET StockQty = COALESCE(StockQty,0) - 1, UpdatedAt = {DateTime.Now} WHERE BottleID IN (SELECT TOP 1 BottleID FROM ProductBottleStyles WHERE ProductID = {productId})", ct);
            }

            // 4) 更新工单状态 + 日志
            await _context.Database.ExecuteSqlInterpolatedAsync(
                $"UPDATE ProductionOrders SET Status = 'WarehouseIn', WarehouseInAt = {DateTime.Now}, UpdatedAt = {DateTime.Now} WHERE ProductionID = {productionId}", ct);
            _context.ProductionLogs.Add(new ProductionLog
            {
                ProductionId = productionId, Status = "WarehouseIn",
                Notes = "成品入库(已按配方扣香调/瓶身)",
                CreatedBy = operatorName ?? "SYSTEM", CreatedAt = DateTime.Now
            });
            await _context.SaveChangesAsync(ct);

            await tx.CommitAsync(ct);
            return new(true, "成品入库成功");
        }
        catch (Exception ex)
        {
            await tx.RollbackAsync(ct);
            return new(false, "入库失败: " + ex.Message);
        }
    }
}
