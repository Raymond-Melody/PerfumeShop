using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

public class LogisticsRepository
{
    private readonly PerfumeShopContext _context;
    public LogisticsRepository(PerfumeShopContext context) => _context = context;

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
