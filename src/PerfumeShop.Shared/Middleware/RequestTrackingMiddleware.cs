using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using System.Diagnostics;

namespace PerfumeShop.Shared;

/// <summary>
/// V19 请求追踪中间件 — 分布式追踪头部注入
/// 对应 V18 includes/api_guard.asp 的 X-Request-ID 头部
/// - 为每个请求生成 X-Request-ID (UUID v4)
/// - 透传客户端 X-API-Version 请求头
/// - 注入 X-Server + X-Response-Time 响应头
/// </summary>
public class RequestTrackingMiddleware
{
    private readonly RequestDelegate _next;

    public RequestTrackingMiddleware(RequestDelegate next)
    {
        _next = next;
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // 1. X-Request-ID：优先使用客户端传入，否则生成新 UUID
        var requestId = context.Request.Headers["X-Request-ID"].FirstOrDefault();
        if (string.IsNullOrEmpty(requestId))
        {
            requestId = Guid.NewGuid().ToString("N");
        }
        context.Items["RequestId"] = requestId;

        // 2. X-API-Version：记录客户端版本标识
        var apiVersion = context.Request.Headers["X-API-Version"].FirstOrDefault();
        if (!string.IsNullOrEmpty(apiVersion))
        {
            context.Items["ApiVersion"] = apiVersion;
        }

        // 3. 计时
        var sw = Stopwatch.StartNew();
        await _next(context);
        sw.Stop();

        // 如果响应已开始流式传输，无法修改响应头
        if (context.Response.HasStarted)
            return;

        context.Response.Headers["X-Request-ID"] = requestId;
        if (!string.IsNullOrEmpty(apiVersion))
            context.Response.Headers["X-API-Version"] = apiVersion;
        context.Response.Headers["X-Server"] = "PerfumeShop-V19";
        context.Response.Headers["X-Response-Time"] = $"{sw.ElapsedMilliseconds}ms";
    }
}

/// <summary>扩展方法</summary>
public static class RequestTrackingMiddlewareExtensions
{
    public static IApplicationBuilder UseRequestTracking(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<RequestTrackingMiddleware>();
    }
}
