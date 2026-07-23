using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using PerfumeShop.Data.Models;

namespace PerfumeShop.Admin.Middleware;

/// <summary>
/// IP 黑名单拦截中间件 — 对齐 V18 IPBlacklist 的访问拦截语义（V18 只存不拦，V19 补齐运行时拦截）。
/// 命中活跃且未过期的封禁记录 → 累加 HitCount/LastHitAt 并返回 403。
/// 活跃名单以 30s TTL 缓存，避免每请求查库。
/// </summary>
public class IpBlacklistMiddleware
{
    private readonly RequestDelegate _next;
    private readonly IMemoryCache _cache;
    private const string CacheKey = "active_ip_blacklist_set";
    private static readonly TimeSpan CacheTtl = TimeSpan.FromSeconds(30);

    public IpBlacklistMiddleware(RequestDelegate next, IMemoryCache cache)
    {
        _next = next;
        _cache = cache;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        var ip = GetClientIp(context);
        if (!string.IsNullOrEmpty(ip))
        {
            var blocked = await GetActiveSetAsync(context);
            if (blocked.Contains(ip))
            {
                await RecordHitAsync(context, ip);
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                context.Response.ContentType = "text/html; charset=utf-8";
                await context.Response.WriteAsync(
                    "<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\"><title>403</title></head>" +
                    "<body style=\"font-family:sans-serif;background:#1a1a2e;color:#e0e0e0;text-align:center;padding:80px;\">" +
                    "<h1 style=\"color:#F44336;\">403 - 访问被拒绝</h1>" +
                    "<p>您的 IP 地址已被列入黑名单，暂时无法访问本系统。如有疑问请联系管理员。</p></body></html>");
                return;
            }
        }
        await _next(context);
    }

    private async Task<HashSet<string>> GetActiveSetAsync(HttpContext context)
    {
        if (_cache.TryGetValue(CacheKey, out HashSet<string>? cached) && cached != null)
            return cached;

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var options = context.RequestServices.GetRequiredService<DbContextOptions<PerfumeShopContext>>();
            await using var db = new PerfumeShopContext(options);
            var now = DateTime.Now;
            var ips = await db.Ipblacklists.AsNoTracking()
                .Where(b => b.IsActive == true && (b.ExpiresAt == null || b.ExpiresAt > now))
                .Select(b => b.Ipaddress)
                .ToListAsync();
            foreach (var ip in ips) if (!string.IsNullOrEmpty(ip)) set.Add(ip);
        }
        catch { /* 查询失败则放行，避免误伤 */ }

        _cache.Set(CacheKey, set, new MemoryCacheEntryOptions
        {
            AbsoluteExpirationRelativeToNow = CacheTtl,
            Size = 1
        });
        return set;
    }

    private static async Task RecordHitAsync(HttpContext context, string ip)
    {
        try
        {
            var options = context.RequestServices.GetRequiredService<DbContextOptions<PerfumeShopContext>>();
            await using var db = new PerfumeShopContext(options);
            var now = DateTime.Now;
            var row = await db.Ipblacklists.FirstOrDefaultAsync(b =>
                b.Ipaddress == ip && b.IsActive == true && (b.ExpiresAt == null || b.ExpiresAt > now));
            if (row != null)
            {
                row.HitCount = (row.HitCount ?? 0) + 1;
                row.LastHitAt = now;
                await db.SaveChangesAsync();
            }
        }
        catch { }
    }

    private static string GetClientIp(HttpContext context)
    {
        var fwd = context.Request.Headers["X-Forwarded-For"].FirstOrDefault();
        if (!string.IsNullOrEmpty(fwd)) return fwd.Split(',')[0].Trim();
        return context.Connection.RemoteIpAddress?.ToString() ?? "";
    }
}

public static class IpBlacklistMiddlewareExtensions
{
    public static IApplicationBuilder UseIpBlacklist(this IApplicationBuilder app) =>
        app.UseMiddleware<IpBlacklistMiddleware>();
}
