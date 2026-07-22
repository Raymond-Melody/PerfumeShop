using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Security.Cryptography;
using System.Text;

namespace PerfumeShop.Shared;

/// <summary>
/// API Key + HMAC-SHA256 认证中间件（双模式）
/// 对应 V18 includes/api_auth.asp + includes/api_guard.asp
/// 双模式认证：
///   1. Cookie 模式：V19_AUTH_TOKEN Cookie → 委托给 AuthBridgeMiddleware 处理
///   2. API Key 模式：X-Api-Key Header + HMAC-SHA256 签名
/// </summary>
public class ApiKeyAuthMiddleware
{
    private readonly RequestDelegate _next;
    private readonly string? _apiKey;
    private readonly string? _apiSecret;
    private readonly bool _enableApiKey;
    private readonly bool _enableCookieBridge;
    private readonly string _cookieName;
    private readonly HashSet<string> _publicPaths;
    private readonly ILogger<ApiKeyAuthMiddleware> _logger;

    public ApiKeyAuthMiddleware(RequestDelegate next, IConfiguration config, ILogger<ApiKeyAuthMiddleware> logger)
    {
        _next = next;
        _logger = logger;

        _apiKey = config["ApiAuth:ApiKey"] ?? config["API_KEY"];
        _apiSecret = config["ApiAuth:ApiSecret"] ?? config["API_SECRET"];

        // 双模式认证配置
        _enableApiKey = config.GetValue<bool>("Auth:EnableApiKey", true);
        _enableCookieBridge = config.GetValue<bool>("Auth:EnableCookieBridge", true);
        _cookieName = config["Auth:CookieName"] ?? "V19_AUTH_TOKEN";

        _publicPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "/api/health",
            "/api/metrics",
            "/swagger",
            "/api/notifications/stream"
        };
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // 只拦截 /api 路径
        if (!context.Request.Path.StartsWithSegments("/api"))
        {
            await _next(context);
            return;
        }

        // 公开端点跳过认证
        foreach (var pp in _publicPaths)
        {
            if (context.Request.Path.StartsWithSegments(pp))
            {
                await _next(context);
                return;
            }
        }

        // === 模式1: Cookie Bridge 认证（V18→V19 Session 互通）===
        if (_enableCookieBridge)
        {
            var hasCookie = context.Request.Cookies.ContainsKey(_cookieName);
            if (hasCookie)
            {
                // Cookie 存在 → 由 AuthBridgeMiddleware 负责验证和注入身份
                // 此处只检查 AuthBridge 是否已经设置了 User
                if (context.User?.Identity?.IsAuthenticated == true)
                {
                    _logger.LogDebug("ApiKeyAuth: Cookie Bridge 认证通过, UserId={UserId}",
                        context.User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value);
                    await _next(context);
                    return;
                }
                // Cookie 存在但 AuthBridge 未认证（Token 可能无效），继续检查 API Key
            }
        }

        // === 模式2: API Key + HMAC 认证 ===
        if (!_enableApiKey)
        {
            // API Key 模式被禁用
            if (context.User?.Identity?.IsAuthenticated != true)
            {
                context.Response.StatusCode = 401;
                context.Response.ContentType = "application/json";
                await context.Response.WriteAsync(
                    "{\"success\":false,\"message\":\"Authentication required\",\"code\":\"UNAUTHORIZED\"}");
                return;
            }
            await _next(context);
            return;
        }

        // 如果未配置密钥，跳过认证（开发模式）
        if (string.IsNullOrEmpty(_apiKey) || string.IsNullOrEmpty(_apiSecret))
        {
            await _next(context);
            return;
        }

        // 验证 API Key
        var providedKey = context.Request.Headers["X-Api-Key"].FirstOrDefault();
        if (string.IsNullOrEmpty(providedKey) || providedKey != _apiKey)
        {
            context.Response.StatusCode = 401;
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsync(
                "{\"success\":false,\"message\":\"Invalid or missing API Key\",\"code\":\"UNAUTHORIZED\"}");
            return;
        }

        // 验证 HMAC 签名（对写操作强校验）
        if (IsWriteMethod(context.Request.Method))
        {
            var timestamp = context.Request.Headers["X-Timestamp"].FirstOrDefault();
            var signature = context.Request.Headers["X-Signature"].FirstOrDefault();

            if (string.IsNullOrEmpty(timestamp) || string.IsNullOrEmpty(signature))
            {
                context.Response.StatusCode = 401;
                context.Response.ContentType = "application/json";
                await context.Response.WriteAsync(
                    "{\"success\":false,\"message\":\"HMAC signature required for write operations\",\"code\":\"HMAC_REQUIRED\"}");
                return;
            }

            // 防重放攻击：时间戳偏差不超过 5 分钟
            if (long.TryParse(timestamp, out var ts))
            {
                var requestTime = DateTimeOffset.FromUnixTimeSeconds(ts);
                if (Math.Abs((DateTimeOffset.UtcNow - requestTime).TotalMinutes) > 5)
                {
                    context.Response.StatusCode = 401;
                    context.Response.ContentType = "application/json";
                    await context.Response.WriteAsync(
                        "{\"success\":false,\"message\":\"Request timestamp expired\",\"code\":\"TIMESTAMP_EXPIRED\"}");
                    return;
                }
            }

            // 计算期望签名
            var body = "";
            if (context.Request.ContentLength > 0)
            {
                context.Request.EnableBuffering();
                using var reader = new StreamReader(context.Request.Body, Encoding.UTF8, leaveOpen: true);
                body = await reader.ReadToEndAsync();
                context.Request.Body.Position = 0;
            }

            var expectedSig = ComputeHmac(_apiSecret, $"{context.Request.Method}|{context.Request.Path}|{timestamp}|{body}");
            if (!string.Equals(signature, expectedSig, StringComparison.OrdinalIgnoreCase))
            {
                context.Response.StatusCode = 401;
                context.Response.ContentType = "application/json";
                await context.Response.WriteAsync(
                    "{\"success\":false,\"message\":\"Invalid HMAC signature\",\"code\":\"INVALID_SIGNATURE\"}");
                return;
            }
        }

        await _next(context);
    }

    private static bool IsWriteMethod(string method) =>
        method is "POST" or "PUT" or "PATCH" or "DELETE";

    public static string ComputeHmac(string secret, string message)
    {
        using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
        var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(message));
        return Convert.ToBase64String(hash);
    }
}

/// <summary>扩展方法：注册 API 认证中间件</summary>
public static class ApiKeyAuthExtensions
{
    public static IApplicationBuilder UseApiKeyAuth(this IApplicationBuilder builder)
    {
        return builder.UseMiddleware<ApiKeyAuthMiddleware>();
    }
}
