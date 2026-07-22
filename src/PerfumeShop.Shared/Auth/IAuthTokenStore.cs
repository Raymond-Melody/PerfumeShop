using System.Security.Claims;

namespace PerfumeShop.Shared.Auth;

/// <summary>
/// Token 验证存储接口 — 解耦 Shared 中间件与 Data 层
/// 实现方（Data/Api 项目）负责查询 AuthTokens 表并返回 ClaimsPrincipal
/// </summary>
public interface IAuthTokenStore
{
    /// <summary>
    /// 根据 Token 哈希值验证并返回对应的 ClaimsPrincipal
    /// </summary>
    /// <param name="tokenHash">SHA-256 哈希后的 Token（64 位十六进制）</param>
    /// <returns>验证通过返回 ClaimsPrincipal，失败返回 null</returns>
    Task<ClaimsPrincipal?> ValidateTokenAsync(string tokenHash);
}
