using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Data.Services;

namespace PerfumeShop.Data.Repositories;

public class LogisticsRepository
{
    private readonly PerfumeShopContext _context;
    private readonly IInventoryLedger _ledger;
    public LogisticsRepository(PerfumeShopContext context, IInventoryLedger ledger)
    {
        _context = context;
        _ledger = ledger;
    }

    public async Task<(List<Order> Items, int Total)> GetShippingOrdersAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.Orders.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(o => o.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(o => o.OrderId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<Order?> GetOrderAsync(int id) => await _context.Orders.FindAsync(id);
    public async Task UpdateOrderStatusAsync(int orderId, string status)
    {
        await _context.Orders.Where(o => o.OrderId == orderId)
            .ExecuteUpdateAsync(s => s.SetProperty(o => o.Status, status));
    }

    // ===== V21: 发货（standard 扣成品库存 + OUT 流水）=====
    // 对标 V18 admin/logistics/shipping_orders.asp ship

    public record ShipResult(bool Success, string Message);

    /// <summary>
    /// 订单发货：更新物流字段；对订单中 standard(品牌定香) 产品扣 ProductInventory 并写 OUT 流水。
    /// custom/kol 不扣成品库（已在工单入库时按配方扣香调/瓶身）。
    /// </summary>
    public async Task<ShipResult> ShipOrderAsync(
        int orderId, string? company, string? trackingNo, string? notes, string? operatorName, CancellationToken ct = default)
    {
        if (orderId <= 0) return new(false, "无效订单");
        using var tx = await _context.Database.BeginTransactionAsync(ct);
        try
        {
            var nowTs = DateTime.Now;
            await _context.Database.ExecuteSqlInterpolatedAsync(
                $@"UPDATE Orders SET ShippingStatus = 'Shipped', ShippingCompany = {company}, TrackingNumber = {trackingNo},
                   ShippingNotes = {notes}, ShippedAt = {nowTs}, UpdatedAt = {nowTs} WHERE OrderID = {orderId}", ct);

            // 缓冲 standard 明细(ProductId:Quantity)，避免遍历中嵌套写
            var stdItems = await (from od in _context.OrderDetails.AsNoTracking()
                                  join p in _context.Products.AsNoTracking() on od.ProductId equals p.ProductId
                                  where od.OrderId == orderId && p.ProductType == "standard"
                                  select new { od.ProductId, od.Quantity }).ToListAsync(ct);

            foreach (var it in stdItems)
            {
                if (it.ProductId <= 0 || it.Quantity <= 0) continue;
                await _context.Database.ExecuteSqlInterpolatedAsync(
                    $"UPDATE ProductInventory SET StockQty = COALESCE(StockQty,0) - {it.Quantity}, UpdatedAt = {DateTime.Now} WHERE ProductID = {it.ProductId}", ct);

                await _ledger.WriteTransactionAsync(new InvTxn(
                    NoteId: 0, MaterialId: null, ProductId: it.ProductId, Quantity: -it.Quantity,
                    TransactionType: "销售出库", Direction: "OUT", ReferenceType: "Order",
                    ReferenceOrderId: orderId, UnitCost: null,
                    Notes: "品牌定香发货扣成品库存", CreatedBy: operatorName ?? "SYSTEM"), ct);
            }

            await tx.CommitAsync(ct);
            return new(true, "发货成功");
        }
        catch (Exception ex)
        {
            await tx.RollbackAsync(ct);
            return new(false, "发货失败: " + ex.Message);
        }
    }
    public async Task<(List<AfterSale> Items, int Total)> GetReturnsAsync(int page, int pageSize, string? status = null)
    {
        var q = _context.AfterSales.AsNoTracking().AsQueryable();
        if (!string.IsNullOrWhiteSpace(status)) q = q.Where(a => a.Status == status);
        var total = await q.CountAsync();
        var items = await q.OrderByDescending(a => a.AfterSalesId).Skip((page - 1) * pageSize).Take(pageSize).ToListAsync();
        return (items, total);
    }
    public async Task<AfterSale?> GetReturnAsync(int id) => await _context.AfterSales.FindAsync(id);
    public async Task<List<ShippingCompany>> GetCarriersAsync() =>
        await _context.ShippingCompanies.AsNoTracking().ToListAsync();
    public async Task<ShippingCompany?> GetCarrierAsync(int id) => await _context.ShippingCompanies.FindAsync(id);
    public async Task SaveCarrierAsync(ShippingCompany carrier)
    {
        if (carrier.CompanyId == 0) _context.ShippingCompanies.Add(carrier);
        else _context.ShippingCompanies.Update(carrier);
        await _context.SaveChangesAsync();
    }
    public async Task DeleteCarrierAsync(int id)
    {
        var c = await _context.ShippingCompanies.FindAsync(id);
        if (c != null) { _context.ShippingCompanies.Remove(c); await _context.SaveChangesAsync(); }
    }
}
