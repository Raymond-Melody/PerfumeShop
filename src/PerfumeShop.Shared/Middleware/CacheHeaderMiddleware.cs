using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;

namespace PerfumeShop.Shared;

/// <summary>
/// V19 缓存响应头中间件 — 对应 V18 cache_manager.asp CM_RecordCacheHeaders
/// 读取 HttpContext.Items["CacheStatus"] 并注入 X-Cache 响应头
/// 同时添加 Cache-Version 和 stale-while-revalidate 扩展
/// </summary>
public class CacheHeaderMiddleware
{
    private readonly RequestDelegate _next;

    public CacheHeaderMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        await _next(context);

        // 如果响应已开始流式传输（如 IResult 直接写入），无法修改响应头
        if (context.Response.HasStarted)
            return;

        // 检查缓存状态（由 ICacheService.GetAsync/GetOrSetAsync 设置）
        var cacheStatus = context.Items["CacheStatus"] as string;
        if (!string.IsNullOrEmpty(cacheStatus))
        {
            context.Response.Headers["X-Cache"] = cacheStatus;
        }
        else
        {
            // 默认 BYPASS（未经过缓存层的请求）
            context.Response.Headers["X-Cache"] = "BYPASS";
        }

        // 缓存版本标识
        context.Response.Headers["X-Cache-Version"] = "v19";
    }
}

/// <summary>扩展方法</summary>
public static class CacheHeaderMiddlewareExtensions
{
    public static IApplicationBuilder UseCacheHeaders(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<CacheHeaderMiddleware>();
    }
}
