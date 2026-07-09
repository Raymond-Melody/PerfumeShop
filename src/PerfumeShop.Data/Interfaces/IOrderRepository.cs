using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 订单仓储接口 — 订单相关的业务查询
/// </summary>
public interface IOrderRepository : IRepository<Order>
{
    /// <summary>按用户ID获取订单列表</summary>
    Task<IEnumerable<Order>> GetByUserIdAsync(int userId, CancellationToken ct = default);

    /// <summary>按订单号获取订单</summary>
    Task<Order?> GetByOrderNoAsync(string orderNo, CancellationToken ct = default);

    /// <summary>按状态获取订单</summary>
    Task<IEnumerable<Order>> GetByStatusAsync(string status, CancellationToken ct = default);

    /// <summary>获取用户最近的订单</summary>
    Task<IEnumerable<Order>> GetRecentByUserIdAsync(int userId, int count = 10, CancellationToken ct = default);

    /// <summary>获取用户的订单数量</summary>
    Task<int> GetOrderCountByUserIdAsync(int userId, CancellationToken ct = default);

    /// <summary>分页获取用户订单</summary>
    Task<(IEnumerable<Order> Items, int TotalCount)> GetPagedByUserIdAsync(
        int userId, int page, int pageSize, CancellationToken ct = default);
}
