using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 生产仓储实现 — V19 M4-C
/// </summary>
public class ProductionRepository : Repository<ProductionOrder>, IProductionRepository
{
    public ProductionRepository(PerfumeShopContext context) : base(context) { }

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
                var details = await _context.OrderDetails.AsNoTracking()
                    .Where(d => d.OrderId == order.OrderId).ToListAsync(ct);

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
}
