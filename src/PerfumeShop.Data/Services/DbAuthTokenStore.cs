using Microsoft.EntityFrameworkCore;
using PerfumeShop.Data.Models;
using PerfumeShop.Shared.Auth;
using System.Security.Claims;

namespace PerfumeShop.Data.Services;

/// <summary>
/// 基于 AuthTokens 数据库表的 Token 验证实现
/// 查询 AuthTokens 表验证 Token，返回对应的 ClaimsPrincipal
/// </summary>
public class DbAuthTokenStore : IAuthTokenStore
{
    private readonly PerfumeShopContext _db;

    public DbAuthTokenStore(PerfumeShopContext db)
    {
        _db = db;
    }

    public async Task<ClaimsPrincipal?> ValidateTokenAsync(string tokenHash)
    {
        var now = DateTime.UtcNow;

        var authToken = await _db.AuthTokens
            .AsNoTracking()
            .FirstOrDefaultAsync(t =>
                t.Token == tokenHash &&
                t.IsActive &&
                t.ExpiresAt > now);

        if (authToken == null)
            return null;

        // 查询关联用户信息
        var user = await _db.Users
            .AsNoTracking()
            .FirstOrDefaultAsync(u => u.UserId == authToken.UserId);

        if (user == null)
            return null;

        // 构建 ClaimsPrincipal
        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, user.UserId.ToString()),
            new(ClaimTypes.Name, user.Username),
            new(ClaimTypes.Email, user.Email),
            new("AuthTokenSource", authToken.Source),
            new("AuthTokenId", authToken.TokenId.ToString())
        };

        // 角色映射
        if (!string.IsNullOrEmpty(user.UserRole))
        {
            claims.Add(new Claim(ClaimTypes.Role, user.UserRole));
        }
        else if (authToken.Source == "AdminLogin")
        {
            claims.Add(new Claim(ClaimTypes.Role, "Admin"));
        }
        else
        {
            claims.Add(new Claim(ClaimTypes.Role, "User"));
        }

        var identity = new ClaimsIdentity(claims, "AuthBridge");
        return new ClaimsPrincipal(identity);
    }
}
