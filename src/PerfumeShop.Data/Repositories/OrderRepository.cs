using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Repositories;

/// <summary>
/// 订单仓储实现
/// </summary>
public class OrderRepository : Repository<Order>, IOrderRepository
{
    public OrderRepository(PerfumeShopContext context) : base(context) { }

    // ========== 重写 GetByIdAsync — Order 为 keyless 实体 ==========

    public override async Task<Order?> GetByIdAsync(int id, CancellationToken ct = default)
    {
        return await _dbSet.FirstOrDefaultAsync(o => o.OrderId == id, ct);
    }

    // ========== IOrderRepository 实现 ==========

    public async Task<IEnumerable<Order>> GetByUserIdAsync(int userId, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(o => o.UserId == userId)
            .OrderByDescending(o => o.CreatedAt)
            .ToListAsync(ct);
    }

    public async Task<Order?> GetByOrderNoAsync(string orderNo, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .FirstOrDefaultAsync(o => o.OrderNo == orderNo, ct);
    }

    public async Task<IEnumerable<Order>> GetByStatusAsync(string status, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(o => o.Status == status)
            .OrderByDescending(o => o.CreatedAt)
            .ToListAsync(ct);
    }

    public async Task<IEnumerable<Order>> GetRecentByUserIdAsync(int userId, int count = 10, CancellationToken ct = default)
    {
        return await _dbSet.AsNoTracking()
            .Where(o => o.UserId == userId)
            .OrderByDescending(o => o.CreatedAt)
            .Take(count)
            .ToListAsync(ct);
    }

    public async Task<int> GetOrderCountByUserIdAsync(int userId, CancellationToken ct = default)
    {
        return await _dbSet.CountAsync(o => o.UserId == userId, ct);
    }

    public async Task<(IEnumerable<Order> Items, int TotalCount)> GetPagedByUserIdAsync(
        int userId, int page, int pageSize, CancellationToken ct = default)
    {
        var query = _dbSet.AsNoTracking().Where(o => o.UserId == userId);
        int totalCount = await query.CountAsync(ct);
        var items = await query
            .OrderByDescending(o => o.CreatedAt)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .ToListAsync(ct);
        return (items, totalCount);
    }
}
