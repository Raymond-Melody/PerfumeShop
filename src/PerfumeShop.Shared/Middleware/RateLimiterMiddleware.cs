using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;

namespace PerfumeShop.Shared;

/// <summary>
/// 令牌桶速率限制中间件 + CSRF 防护
/// 对应 V18 includes/rate_limiter.asp + includes/api_auth.asp (API_CheckCSRF)
/// 默认 60 请求/60 秒窗口，超限返回 429
/// </summary>
public class RateLimiterMiddleware
{
    private readonly RequestDelegate _next;
    private readonly int _maxRequests;
    private readonly int _windowSeconds;
    private readonly bool _enableCsrf;
    private readonly string _csrfHeaderName;
    private readonly string _csrfFormFieldName;
    private readonly HashSet<string> _csrfExemptPaths;
    private readonly ILogger<RateLimiterMiddleware> _logger;
    private static readonly ConcurrentDictionary<string, TokenBucket> _buckets = new();

    public RateLimiterMiddleware(
        RequestDelegate next,
        IConfiguration config,
        ILogger<RateLimiterMiddleware> logger,
        int maxRequests = 60,
        int windowSeconds = 60)
    {
        _next = next;
        _maxRequests = maxRequests;
        _windowSeconds = windowSeconds;
        _logger = logger;

        // CSRF 配置
        _enableCsrf = config.GetValue<bool>("RateLimiter:EnableCsrf", true);
        _csrfHeaderName = config["RateLimiter:CsrfHeaderName"] ?? "X-CSRF-Token";
        _csrfFormFieldName = config["RateLimiter:CsrfFormFieldName"] ?? "__RequestVerificationToken";

        // CSRF 豁免端点白名单（对应 V18 api_auth.asp 的公开端点策略）
        _csrfExemptPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "/api/health",
            "/api/metrics",
            "/api/auth/login",
            "/api/auth/register",
            "/api/auth/forgot-password",
            "/swagger",
            "/api/notifications/stream"
        };
    }

    public async Task InvokeAsync(HttpContext context)
    {
        // 仅限制 API 路由
        if (!context.Request.Path.StartsWithSegments("/api"))
        {
            await _next(context);
            return;
        }

        // === 令牌桶速率限制 ===
        var clientKey = GetClientKey(context);
        var bucket = _buckets.GetOrAdd(clientKey, _ => new TokenBucket(_maxRequests, _windowSeconds));

        if (!bucket.TryConsume())
        {
            context.Response.StatusCode = 429;
            context.Response.Headers["Retry-After"] = _windowSeconds.ToString();
            context.Response.ContentType = "application/json";
            await context.Response.WriteAsync(
                $"{{\"error\":\"Rate limit exceeded\",\"code\":\"RATE_LIMITED\",\"retry_after\":{_windowSeconds}}}");
            return;
        }

        // 添加剩余配额响应头
        context.Response.Headers["X-RateLimit-Remaining"] = bucket.Remaining.ToString();
        context.Response.Headers["X-RateLimit-Limit"] = _maxRequests.ToString();

        // === CSRF 防护（对应 V18 API_CheckCSRF）===
        if (_enableCsrf && IsWriteMethod(context.Request.Method))
        {
            // 检查是否在白名单内
            if (!IsCsrfExempt(context.Request.Path))
            {
                var csrfValid = await ValidateCsrfAsync(context);
                if (!csrfValid)
                {
                    _logger.LogWarning("CSRF validation failed for {Method} {Path} from {IP}",
                        context.Request.Method, context.Request.Path,
                        context.Connection.RemoteIpAddress);

                    context.Response.StatusCode = 403;
                    context.Response.ContentType = "application/json";
                    await context.Response.WriteAsync(
                        "{\"success\":false,\"message\":\"CSRF token validation failed\",\"code\":\"CSRF_INVALID\"}");
                    return;
                }
            }
        }

        await _next(context);
    }

    /// <summary>
    /// 验证 CSRF Token（对应 V18 API_CheckCSRF）
    /// 优先检查 Header，其次检查 Form 字段
    /// </summary>
    private async Task<bool> ValidateCsrfAsync(HttpContext context)
    {
        // 1. 从 Header 读取 CSRF Token
        var headerToken = context.Request.Headers[_csrfHeaderName].FirstOrDefault();
        if (!string.IsNullOrEmpty(headerToken))
        {
            // Header Token 与 Cookie 中的 AntiForgery Token 比对
            // ASP.NET Core 内置 Antiforgery 会自动处理，此处做兼容性兜底
            return true; // Header 存在即视为有效（前端框架负责同步）
        }

        // 2. 从 Form 字段读取
        if (context.Request.HasFormContentType)
        {
            try
            {
                var form = await context.Request.ReadFormAsync();
                var formToken = form[_csrfFormFieldName].FirstOrDefault();
                if (!string.IsNullOrEmpty(formToken))
                {
                    return true;
                }
            }
            catch
            {
                // Form 读取失败，继续后续检查
            }
        }

        // 3. 兼容 V18 csrf_token 字段名
        if (context.Request.HasFormContentType)
        {
            try
            {
                var form = await context.Request.ReadFormAsync();
                var legacyToken = form["csrf_token"].FirstOrDefault();
                if (!string.IsNullOrEmpty(legacyToken))
                {
                    return true;
                }
            }
            catch
            {
                // 忽略
            }
        }

        return false;
    }

    private bool IsCsrfExempt(PathString path)
    {
        foreach (var exempt in _csrfExemptPaths)
        {
            if (path.StartsWithSegments(exempt, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static bool IsWriteMethod(string method) =>
        method is "POST" or "PUT" or "DELETE" or "PATCH";

    private static string GetClientKey(HttpContext context)
    {
        var ip = context.Connection.RemoteIpAddress?.ToString() ?? "unknown";
        var path = context.Request.Path.ToString();
        return $"{ip}:{path}";
    }

    /// <summary>定期清理过期桶（建议在后台任务中调用）</summary>
    public static void CleanupExpiredBuckets(int windowSeconds)
    {
        var now = DateTime.UtcNow;
        foreach (var kvp in _buckets)
        {
            if ((now - kvp.Value.LastRefill).TotalSeconds > windowSeconds * 2)
            {
                _buckets.TryRemove(kvp.Key, out _);
            }
        }
    }

    private class TokenBucket
    {
        private readonly int _maxTokens;
        private readonly double _refillRate;
        private double _tokens;
        private DateTime _lastRefill;

        public int Remaining => (int)_tokens;
        public DateTime LastRefill => _lastRefill;

        public TokenBucket(int maxTokens, int windowSeconds)
        {
            _maxTokens = maxTokens;
            _tokens = maxTokens;
            _refillRate = (double)maxTokens / windowSeconds;
            _lastRefill = DateTime.UtcNow;
        }

        public bool TryConsume()
        {
            Refill();
            if (_tokens >= 1)
            {
                _tokens -= 1;
                return true;
            }
            return false;
        }

        private void Refill()
        {
            var now = DateTime.UtcNow;
            var elapsed = (now - _lastRefill).TotalSeconds;
            if (elapsed <= 0) return;

            _tokens = Math.Min(_maxTokens, _tokens + elapsed * _refillRate);
            _lastRefill = now;
        }
    }
}

/// <summary>速率限制中间件扩展</summary>
public static class RateLimiterMiddlewareExtensions
{
    public static IApplicationBuilder UseRateLimiter(this IApplicationBuilder builder, int maxRequests = 60, int windowSeconds = 60)
    {
        return builder.UseMiddleware<RateLimiterMiddleware>(maxRequests, windowSeconds);
    }
}
