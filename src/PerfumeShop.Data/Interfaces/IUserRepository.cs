using PerfumeShop.Core.Interfaces;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Data.Interfaces;

/// <summary>
/// 用户仓储接口 — 用户认证与查询
/// </summary>
public interface IUserRepository : IRepository<User>
{
    /// <summary>按用户名获取用户</summary>
    Task<User?> GetByUsernameAsync(string username, CancellationToken ct = default);

    /// <summary>按邮箱获取用户</summary>
    Task<User?> GetByEmailAsync(string email, CancellationToken ct = default);

    /// <summary>验证用户登录凭据</summary>
    Task<User?> AuthenticateAsync(string username, string password, CancellationToken ct = default);

    /// <summary>判断用户名是否已存在</summary>
    Task<bool> UsernameExistsAsync(string username, CancellationToken ct = default);

    /// <summary>判断邮箱是否已存在</summary>
    Task<bool> EmailExistsAsync(string email, CancellationToken ct = default);

    /// <summary>获取用户含优惠券信息</summary>
    Task<User?> GetWithCouponsAsync(int userId, CancellationToken ct = default);
}
